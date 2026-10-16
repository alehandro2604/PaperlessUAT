# Paperless – How the Redis cache works today

This document describes what is **already built** in the Paperless app (Angular `my-app` + Node `paperless-backend`). It does not propose changes.

Sources read: `paperless-backend/cacheRoutes.js`, `permissionCheck.js`, `redisClient.js`, and in `my-app/src/app/services/`: `cache.ts`, `cache-session.service.ts`, `file-crawl.service.ts`, `todo-task-store.service.ts`, plus `components/hr-files/hr-files.component.ts`.

**Not read / not verified:** `app.ts` (where the refresh and load flow is wired up), `subordinaryTask.service.ts`, and the other services. Anything about what happens exactly on a browser refresh is partly inferred from the cache code. The app was not run.

---



## 1. Redis in simple terms

Redis is a very fast notepad that lives on the **server**, not in the browser. When the app needs something slow to get (SharePoint folders, comments, tasks), it saves a copy in Redis. Next time, it reads the saved copy instead of asking SharePoint again.

Two important things follow from that:

- Redis is **shared by default**. It is not tied to one browser or device. Whether data is private or shared depends on how the key is named.
- The browser still has to **call your backend** to read from Redis. Redis reduces calls to SharePoint/Graph, not calls from the browser to your server.

---



## 2. The pieces


| Piece                                                 | Where   | Job                                                                                             |
| ----------------------------------------------------- | ------- | ----------------------------------------------------------------------------------------------- |
| `redisClient.js`                                      | backend | Connects to Redis (`REDIS_URL`, default `redis://localhost:6379`)                               |
| `cacheRoutes.js`                                      | backend | The `/api/cache/...` endpoints the browser calls to read/write/delete cache entries             |
| `permissionCheck.js`                                  | backend | Decides whether a user may read a **shared** cache entry                                        |
| `cache.ts` (`CacheService`)                           | browser | Thin client for the `/api/cache` endpoints (get/set/clear/by prefix)                            |
| `file-crawl.service.ts` (`FileCrawlCacheService`)     | browser | In-memory cache in front of Redis for files/comments; decides fresh vs stale                    |
| `todo-task-store.service.ts` (`TodoTaskStoreService`) | browser | Separate To Do cache, one Redis entry per task                                                  |
| `cache-session.service.ts` (`CacheSessionService`)    | browser | Tells the backend which tabs are open, so personal entries can expire after the last tab closes |


There are therefore **two layers**: an in-memory copy in the browser (lost on refresh) and Redis on the server (survives refresh).

---



## 3. Shared vs personal entries

Every key is stored in one of two namespaces:

- **Shared:** `shared:<key>`, one copy for all users and browsers.
- **Personal:** `cache:<userOid>:<key>`, where `userOid` is the user's Entra object id. Only that user can read or wipe it.

The backend treats a key as shared only if it starts with one of these prefixes (`SHARED_KEY_PREFIXES`):


| Prefix                | What it holds                                              |
| --------------------- | ---------------------------------------------------------- |
| `fc:files:folder:`    | All Files folder listings                                  |
| `fc:files:lib:`       | All Files library root listings (one per document library) |
| `fc:tasks:lib:v7:`    | All Files folder comments                                  |
| `fc:tasks:hr-folder:` | HR Files per-person comments                               |


Everything else is personal, for example the user's To Do tasks and `fc:tasks:hr-user:…` (My Tasks, per assignee).

**Why shared?** One person's crawl warms the cache for everybody, so the second person does not have to wait for SharePoint.

---



## 4. How permissions are protected on shared entries

Because shared data is visible to everyone in Redis, the backend checks SharePoint permissions **every time** a shared key is read, written or deleted (`canUseKey` → `canAccessKey` in `permissionCheck.js`).

How the check works:

1. The backend calls Microsoft Graph **as the signed-in user** (on-behalf-of token).
2. Each shared key is mapped to the SharePoint object that guards it:
  - folder listing → that drive item
  - library root listing → that document library (must be one of the user's visible drives)
  - folder comments (`fc:tasks:lib:…`) → that SharePoint list
  - HR person comments (`fc:tasks:hr-folder:…`) → that person's folder in the HR library
3. Graph is asked (batched, up to 20 at a time) whether the user can open that object. Success means allowed; 401/403/404 means denied.
4. The answer is remembered in Redis: **15 minutes** for "allowed", **5 minutes** for "denied" (`perm:<userOid>:<resource>`).
5. If the backend cannot verify (Graph error, unknown key shape, rate limit), it **denies** (fail closed). Temporary errors (429/5xx) are denied but not cached.

Results for list endpoints (`GET /entries`) are filtered, so a user only receives entries for folders and lists they can open.

Consequence: if someone loses access, they can still read the shared snapshot for up to about 15 minutes (the cached "allowed" answer).

The browser side also remembers a 403 for a key for 5 minutes (`DENIED_RETRY_MS`) so it does not keep asking.

---



## 5. Backend endpoints (`/api/cache`)


| Endpoint                  | What it does                                                                                                                                                 |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `GET /entry?key=`         | Read one entry (`{ data, timestamp }` or null). Shared keys: permission-checked. Falls back to an old personal copy if the shared one is missing             |
| `PUT /entry`              | Write one entry. Shared keys: permission-checked, 7-day expiry. Personal keys: 12-hour expiry. Removes the old personal copy of a key once it becomes shared |
| `DELETE /entry?key=`      | Delete one entry                                                                                                                                             |
| `GET /entries?prefix=`    | All entries starting with a prefix (shared ones filtered by permission)                                                                                      |
| `DELETE /entries?prefix=` | Delete by prefix. **Empty prefix wipes only this user's personal keys** (used on logout)                                                                     |
| `POST /session/start`     | A tab announces itself; restores the normal expiry on the user's personal entries                                                                            |
| `POST /session/end`       | A tab says goodbye; if none are left, personal entries get a short expiry                                                                                    |
| `GET /todo`               | This user's whole To Do cache                                                                                                                                |
| `PUT /tasks`              | Save changed To Do tasks and remove deleted ones                                                                                                             |
| `DELETE /todo`            | Forget this user's whole To Do cache                                                                                                                         |
| `GET /delta-test?list=`   | Experimental, read-only test of whether a SharePoint list supports Graph delta queries. The code comment says to remove it once the delta step is built      |


---



## 6. Expiry times at a glance


| What                                       | Time              | Where it is set             |
| ------------------------------------------ | ----------------- | --------------------------- |
| Shared entries in Redis                    | 7 days            | `ENTRY_TTL_SECONDS`         |
| Personal entries in Redis                  | 12 hours          | `PERSONAL_TTL_SECONDS`      |
| Personal entries after the last tab closes | 90 seconds        | `SESSION_END_GRACE_SECONDS` |
| Permission answer "allowed"                | 15 minutes        | `ALLOW_TTL`                 |
| Permission answer "denied"                 | 5 minutes         | `DENY_TTL`                  |
| Browser treats a snapshot as "fresh"       | 60 minutes        | `CACHE_TTL` (file crawl)    |
| Browser drops a saved snapshot at sign-in  | older than 7 days | `PERSIST_MAX_AGE`           |
| To Do counts as "fresh"                    | 60 minutes        | `FRESH_FOR_MS`              |
| Browser remembers a 403                    | 5 minutes         | `DENIED_RETRY_MS`           |


The first three can be changed with environment variables (`CACHE_ENTRY_TTL_SECONDS`, `CACHE_PERSONAL_TTL_SECONDS`, `CACHE_SESSION_GRACE_SECONDS`).

---



## 7. Files and comments (All Files, HR Files)

Handled by `FileCrawlCacheService`, which keeps an in-memory `Map` in front of Redis.

**Reading**

- `get(key)` returns data only if it is less than 60 minutes old. Otherwise it returns null, which tells the caller to re-fetch from SharePoint.
- `getStale(key)` returns data of any age, for an instant first paint while a fresh fetch runs.
- `hydrateFromPersistent(key)` pulls **one key** from Redis into memory when memory misses. It returns the data plus whether it is still fresh. This is how HR person snapshots and shared folder listings load, one at a time, when a screen needs them.

**Writing**

- `set(key, data)` writes to memory and queues a write to Redis (writes run one after another, after sign-in loading has finished).
- `CacheService.set` skips the upload if the content is identical to the last one stored for that key (a fingerprint check), and gzips bodies over 8 KB.
- `setMemoryOnly` is used for payloads too big for Redis. The All Files comments search index is never saved to Redis because it is too large (HTTP 413). It lives in memory only and is rebuilt after a refresh.
- If a save fails with 413, the old Redis copy is deleted so a stale copy is not shown on the next visit.

**Invalidating**

- `invalidate(key)`: delete from memory and Redis (after upload/delete/rename).
- `expire(key)`: mark as needing a refresh without deleting. `get()` misses, but `getStale()` still shows the old rows meanwhile, and the shared Redis copy stays for other users until the fresh `set()` overwrites it.
- `invalidatePrefix(prefix)`: clear a whole family (e.g. after a manual All Files refresh).
- `clearAll()`: used on logout.

**Sign-in (**`bindToUser`**)**

1. Checks which user the saved personal entries belong to (`fc-meta:owner`).
2. If it is a different user, the personal `fc:` entries are deleted. Shared ones are left alone.
3. If it is the same user, all personal `fc:` entries are loaded into memory (this is roughly a 3 MB download, per a code comment). Entries older than 7 days are dropped.
4. The shared `fc:tasks:lib:v7:` entries are then loaded in the background, without blocking the first paint.
5. HR comments (`fc:tasks:hr-folder:`) are **not** bulk-loaded. A code comment says they total about 200 MB across all HR people, and bulk-loading blocked sign-in for 10+ seconds and used about 500 MB of memory. They are pulled one person at a time.

**HR Files screen (**`hr-files.component.ts`**)**
This component only displays the list it is given (search, sort, paging 100 at a time). It paints names from the last saved copy of the domain-user list first (`peekUsers`), then updates when fresh data arrives. The actual loading of HR folders is done by the parent (`app.ts`).

---



## 8. To Do cache

Handled by `TodoTaskStoreService` and the `/todo` and `/tasks` endpoints. It is separate from the file cache because the old single To Do snapshot grew too large for the upload limit and had to be re-sent in full after every change.

**How it is stored (personal, never shared)**

- `cache:<oid>:todo:task:<list>:<id>`: one entry per task
- `cache:<oid>:todo:index`: a Redis set listing this user's task keys
- `cache:<oid>:todo:updated`: when the To Do was last saved
- `cache:<oid>:todo:meta`: bookkeeping for top-ups (see below)

All of these sit under `cache:<oid>:`, so logout's "clear all" wipes them too. They are personal on purpose, because the browser uploads these rows.

**Only changed tasks are uploaded**
`replaceAll(items)` fingerprints every task and compares it to what Redis already has. Only tasks that are new or changed are queued as `upsert`, and tasks that disappeared as `remove`. Saves are debounced (1.5 seconds), sent in chunks of 200 tasks (the backend allows at most 500 per call), and run one after another so a delete can never overtake an earlier upload. If a save fails, the changes are put back and retried with the next change.

So if child 1 changes, only child 1 is uploaded. The parent and child 2 are untouched. (Whether a parent has values calculated from its children was not checked.)

**Top-ups instead of full reloads (**`meta`**)**
The saved `meta` holds:

- `watermarks`: the newest `lastModifiedDateTime` seen for each SharePoint list
- `complete`: whether the last full crawl finished, so the saved tasks are the whole set
- `crawledAt`: when that crawl finished

On reload, the app can ask SharePoint only for items newer than each watermark, instead of paging every list again. Old crawls are redone, because a top-up cannot see lists that were empty at the time. A top-up also does not notice deleted items.

**Loading**
`load()` pulls the user's To Do from Redis once per session and is safe to call repeatedly. `getAll()` returns tasks of any age, and `getFresh()` returns them only if saved within the last hour. If SharePoint answers before Redis, the SharePoint list is kept and the store works out what Redis needs.

**Forgetting**
`invalidate()` clears the in-memory copy and calls `DELETE /todo`, which removes this user's whole To Do cache.

Each task is its own Redis key (`cache:<oid>:todo:task:<list>:<id>`). In `todo-task-store.service.ts`, `replaceAll()` fingerprints every task and uploads only the ones that changed (`PUT /tasks` with `upsert` and `remove`). Updating child 1 uploads child 1 only.

---



## 9. Tab lifetime of personal entries (`CacheSessionService`)

Goal: personal data should not sit in Redis long after the user has left.

- Each tab gets an id stored in `sessionStorage`. A **refresh keeps the same id**; a new tab gets a new one.
- On load, the tab calls `POST /session/start`. The backend adds the tab id to a set and puts the normal 12-hour expiry back on all the user's personal entries.
- When the page goes away (`pagehide`, unless it is going into the back/forward cache), the tab calls `POST /session/end` using `keepalive` so the request outlives the page.
- If **no tabs are left**, all the user's personal entries are set to expire in **90 seconds**. That grace period is long enough for a refresh (the old page says goodbye, the reloaded page says hello) not to lose anything.
- If a tab crashes without saying goodbye, the entries still expire at the 12-hour limit at the latest.
- **Shared entries are never touched** by this logic.

Result: a refresh keeps personal data in Redis. Closing every tab and coming back after more than about 90 seconds starts personal data from scratch. Shared data stays for up to 7 days.

---



## 10. Typical flows

**First visit of the day**

1. Sign in. `bindToUser` loads personal `fc:` entries from Redis into memory (or wipes them if a different user last used this account).
2. Shared comments for lists load in the background.
3. When the user opens a folder or HR person, that one key is fetched from Redis via the backend (permission-checked). A hit paints immediately. If the snapshot is older than 60 minutes, SharePoint is queried too and the result is saved back.
4. On a Redis miss, the data comes from SharePoint and is saved to Redis for everyone who may open it.

**Browser refresh**

1. The in-memory copy disappears.
2. The tab keeps its id, so personal entries in Redis stay.
3. Sign-in loading runs again: personal entries are re-downloaded and shared keys are re-fetched one at a time as screens need them. (Inferred. The exact order in `app.ts` was not read.)
4. Data older than 60 minutes is re-fetched from SharePoint, while old rows are shown meanwhile.

**Editing something (upload, delete, rename, update a task)**

- Files: the affected key is invalidated or expired, and the next load repopulates it.
- To Do: only the changed task is uploaded.

**Logout**

- Memory is cleared, and `DELETE /entries` with an empty prefix removes the user's personal Redis entries. Shared entries remain.

---



## 11. Summary of what exists vs what does not

**Exists**

- Shared, permission-checked cache for All Files and HR Files, with fail-closed permission checks.
- Personal, per-user cache for To Do and similar.
- Per-task To Do storage with changed-only uploads.
- To Do top-ups using per-list watermarks.
- Stale-while-refresh painting (`getStale`), 60-minute freshness, and 7-day maximum age.
- Tab-based expiry of personal entries with a 90-second grace period.
- Upload savings: identical-content skip, gzip over 8 KB, denied-key memory.
- An experimental Graph delta test route.

**Does not exist (yet)**

- A browser-side persistent cache (the app used IndexedDB before, then moved to Redis).
- Delta updates for files and comments; they are replaced as whole snapshots.
- A way for deletions to be noticed by To Do top-ups.
- Persistence of the All Files comments search index.

