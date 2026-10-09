# Cache update: fewer repeated downloads, uploads and SharePoint calls

Date: 2026-10-08 · Branch: `cache-per-task-todo`

## Why this was done

Darren reported two problems:

1. `/api/cache/entry` was receiving payloads over 25 MB (he had to raise the body limit in `server.js`).
2. With the site idle, the app kept sending requests to SharePoint / Graph. That is fine for one
   person, but hundreds of users could hit the Graph API limits.

I checked how Redis caching works for To Do, HR Files and All Files, measured real traffic in the
browser (F12 Network) and the real sizes inside Redis, and fixed what the measurements showed.

## What the measurements showed

| Finding | Evidence |
|---|---|
| The same ~3 MB was downloaded twice at every sign-in | `entries?prefix=fc:` 3,081 kB and `entries?prefix=fc:files:` 3,066 kB |
| An idle tab kept calling SharePoint for 5 minutes after the last click, and polled every 5 / 10 min | `app.constants.ts`, `app.ts` |
| The All Files listing was saved **per user** (14 MB each) plus a second copy per library | Redis: `cache:<user>:fc:files:all-root:v5-paged` 14 MB, `…lib:v5-paged:procurement…` 8.5 MB |
| Normal HR people are small (about 4 KB), a few entries are huge | One HR account is 16.6 MB, see "Still open" |

## How caching works (simple version)

- **Redis** is a fast shared store on the server. The app saves the results of slow SharePoint
  crawls there, so the next screen opens instantly instead of crawling again.
- A cache entry is one **key** (a name) holding one **blob** (a big piece of JSON).
- **Shared** keys are stored once for all users. Before anyone reads or writes one, the backend
  checks they can open that folder, list or library in SharePoint.
- **Personal** keys belong to one user only (for example their To Do list).

| Area | How it is stored |
|---|---|
| To Do | One Redis key **per task**. Only changed tasks are sent. (Already this way.) |
| HR Files | One blob per person, shared |
| All Files comments | One blob per folder, shared |
| All Files file listings | **Changed in this update**: one blob per library, shared |

## What changed

### 1. Removed a duplicate download at sign-in
- **Before:** at sign-in the app downloaded all personal cache entries (`fc:`) and then asked again
  for `fc:files:`, which returned the same entries. About 3 MB twice.
- **After:** `fc:files:` is no longer requested separately. About 3 MB saved per sign-in.
- **File:** `my-app/src/app/services/file-crawl.service.ts` (`SHARED_HYDRATE_PREFIXES`, and its comment)

### 2. Idle tabs stop calling SharePoint sooner, and polling is less frequent
- **Before:** background work stopped 5 min after the last input. To Do was checked every 5 min,
  the open folder's comments every 10 min.
- **After:** stops after 2 min idle. To Do checked every 10 min, open folder comments every 20 min.
- **Why:** fewer Graph requests from tabs nobody is using, which lowers the risk of 429 throttling.
- **Files:**
  - `my-app/src/app/app.constants.ts` (`backgroundIdleAfterMs`: 5 → 2 minutes)
  - `my-app/src/app/app.ts` (`SHAREPOINT_CACHE_POLL_MS`: 5 → 10 minutes, `FOLDER_COMMENTS_POLL_MS`: 10 → 20 minutes)
- **Effect on users:** none in normal use. An edit made directly in SharePoint by someone else can
  take up to 10 minutes (To Do) or 20 minutes (folder comments) to appear in an open tab.
  Changes made inside the app still show immediately.

### 3. All Files library listings are stored once for everyone (permission check kept)
- **Before:** each user saved their own copy of the merged all-libraries list (`all-root`, ~14 MB)
  and their own copy of every library (~8.6 MB for one library). 500 users could mean gigabytes
  of the same files, and repeated uploads and downloads.
- **After:**
  - Each library's listing (`fc:files:lib:…`) is a **shared** key: one copy for all users.
  - The backend allows a read, write or delete only if the library is in the user's own drive
    list **and** Graph confirms they can open its root. A user without access gets `403`.
  - The merged `all-root` list is kept in memory only (never sent to Redis).
  - Libraries load one at a time from Redis when needed, not one big 14 MB download.
  - A library that is already stored complete is not saved again (this removes a repeat upload).
- **Files:**
  - `paperless-backend/cacheRoutes.js`: added `fc:files:lib:` to `SHARED_KEY_PREFIXES`, and comments
  - `paperless-backend/permissionCheck.js`: access check for library keys (`describeKey`), and the user's drive list added to the check context
  - `my-app/src/app/components/All-files/all-files.ts`:
    - `loadAllLibrariesFiles`: loads each library with `hydrateFromPersistent`, and saves `all-root` to memory only (`setMemoryOnly`)
    - `seedLibraryCachesFromAllRoot`: skips a library that is already saved complete
    - the selected-library load: loads that library's shared entry first

### 4. Cleanup script for the old per-user copies
- **File:** `paperless-backend/scripts/cleanup-personal-file-listings.js` (new)
- Deletes `cache:<user>:fc:files:all-root:*` and `cache:<user>:fc:files:lib:*`, which nothing
  reads any more. The shared copies are rebuilt when someone opens a library.
- **Dry run by default** (lists what it would delete). Delete with `--delete`. Use `--env .env.uat` for UAT.
- Not run yet. Dry run on dev Redis: 1 key, about 13.8 MB.

## Results of testing (dev, one user, localhost)

| Test | Result |
|---|---|
| Sign-in | `entries?prefix=fc:files:` no longer requested |
| Idle for about 6 min after the first load finished | **0** requests to Graph and `/api/cache` |
| Opening a normal HR person (Aaron Schembri) | Two small uploads: 136 B and 4.1 KB (gzipped) |
| All Files, loading all libraries | 15 small `GET` requests, none for `all-root`, one 196 B upload, no errors |
| Redis after the change | Library listings are `shared:fc:files:lib:…` (one copy each) |

## Still open

1. **A second user without access** to a library was **not** tested. Test it: that user must get no listing for that library.
2. **One HR key is 16.6 MB:** `shared:fc:tasks:hr-folder:v9:0 pauline.bonavia@emoffice365.onmicrosoft.com`.
   Find out what this account is. It is the biggest single payload left.
3. **HR Files and All Files comments are still saved as one blob per person/folder.** A one-task change
   re-sends the whole list. Per-task storage (as To Do already has) would fix this.
4. **The first load after sign-in is still heavy** on Graph (the To Do crawl pages through many lists).
5. The access check is at library level. A library with items that have their own restricted
   permissions could show those item names to someone who can open the library.
6. Run the Redis size check on the UAT / live Redis. The numbers above are from the dev Redis.

## Rollout steps

1. Deploy the backend first (`cacheRoutes.js`, `permissionCheck.js`), then the Angular app.
2. Test with two accounts (one with access to a library, one without).
3. Dry run the cleanup script on UAT, review, then run it with `--delete`.
4. Darren can set the body limit back once large uploads are gone.
