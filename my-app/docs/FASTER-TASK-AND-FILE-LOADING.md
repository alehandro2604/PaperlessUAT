# Faster Task & File Loading

This document lists how the app loads **tasks** and **folder files** faster: in-memory caching, progressive first paint, idle prefetch, and Attachments-first UX.

Related correctness doc (UAT vs Live HR Comments algorithm): [`HR-FILES-LIVE-TASK-LOADING-FIX.md`](./HR-FILES-LIVE-TASK-LOADING-FIX.md).

---

## 1. Goals

| Goal | Approach |
|---|---|
| Reopen the same folder / tasks quickly | In-memory TTL cache (`FileCrawlCacheService`) |
| Feel fast on the **first** open | Progressive paint + Attachments first + idle prefetch |
| Keep data correct after changes | Manual refresh invalidates cache and re-crawls |
| Avoid Graph throttling (429) | Limited concurrency, smaller pages, no Live `$orderby` |

**Important:** Cache cannot make a true first Graph call faster. Prefetch and progressive UI make that first open *feel* faster; cache makes the 2nd+ open within TTL instant.

---

## 2. Cache service

**File:** `src/app/services/file-crawl.service.ts`

- In-memory `Map` with **15-minute TTL**
- API: `get` / `set` / `invalidate` / `invalidatePrefix` / `clearAll`
- Not persisted to `localStorage` (cleared on full page reload)
- Last opened HR Files person is remembered across To Do / All Files tab switches and restored from `tasks:hr-folder:...` instantly

### Cache keys

| Key pattern | Used by | Data |
|---|---|---|
| `tasks:hr-user` | Main To Do / Comments recent HR load | Mapped HR tasks for signed-in user |
| `tasks:hr-folder:{folderName}` | HRPersonal person folder Comments | Tasks for that person folder |
| `tasks:lib:{listName}:{folderName}` | Document-library folder Comments | Tasks for that library folder |
| `files:lib:{libraryName}` | All Files → Files filter | Root children of one library |
| `files:all-root` | All Files → All filter | Root folders/files across libraries |
| `files:deep:{libraryName}` | Attachments search deep crawl | Recursive files in one library |
| `files:folder:{driveId}:{folderId}` | Attachments panel folder browse | One-level children of a drive folder |

Names are normalized (`trim` + lower case) before building keys.

---

## 3. How tasks load faster

### 3.1 Cache hit path (2nd+ open within TTL)

```
User opens same folder / reloads same HR task set
        │
        ▼
  fileCrawlCache.get(key)
        │
   ┌────┴────┐
   │ hit     │ miss
   ▼         ▼
 Show UI   Crawl Graph → set(key) → Show UI
 instantly
```

Wired in:

- `AppComponent.loadAllFilesTasksForLibrary` — folder Comments
- `AppComponent.loadHrTasks` / `loadHrTasksFromLists` — main user HR tasks
- Successful crawls call `cache.set(...)`

Silent background refresh uses `skipCache: true` so it can still refresh Graph and update the cache.

### 3.2 Progressive Comments paint (first open)

**Document-library folders** (`loadDocumentLibraryTasksForFolder`) — All Files only:

- Try a fast Graph `$filter=fields/Title eq '{folder}'` first (skipped if Live rejects it).
- Then progressive scan: **200 items/page**, up to **50 pages** (~10k), so matches past early pages are not missed.
- Show matches as soon as found; later pages merge + refresh.
- Final result is written to `tasks:lib:...`.
- **Does not change HRPersonal loading** (separate code path + constants).

**HRPersonal person folders** (`loadHrPersonalTasksForFolder` / `collectHrPersonalTasksFromLists`):

- Lists are queried in small parallel batches (`hrFilesTaskListConcurrency`).
- Pass 1 (own submissions) and Pass 2 (related same-ID / same-name steps) both run before Comments is marked **done**.
- The loading message stays on (`Loading tasks for … (including related steps)...`) until Pass 2 finishes, so related tasks do not appear to “pop in” after the UI looks finished.
- Final complete result is stored in `tasks:hr-folder:...` (15 min TTL) and restored when returning from To Do / All Files.

### 3.3 Attachments first on folder click

When a folder is clicked (All Files or HR Files):

1. Switch to **Attachments** immediately (`getAttachmentsTabMobileState`).
2. Load folder children (often already prefetched/cached).
3. Load Comments/tasks in the **background**.

User sees files first; tasks fill in without blocking the Attachments panel.

### 3.4 Manual refresh

| Section | What is invalidated |
|---|---|
| All Files | Folder task keys for selection + `files:folder:` prefix + All Files file keys |
| To Do | `tasks:hr-user` |
| HR Files | HR folder task key + current drive-folder contents key |

Refresh then re-crawls with `skipCache: true` where needed.

---

## 4. How folder files load faster

### 4.1 Per-folder contents cache

`loadDriveFolderContents` uses `files:folder:{driveId}:{folderId}`:

- Cache hit → Attachments UI updates immediately.
- Cache miss → one Graph `/children` call → `set` → UI.

### 4.2 Idle prefetch (warms cache before click)

1. **Visible All Files folders** — when the list paints, parent receives `foldersForPrefetch` and silently fetches up to ~40 folders (concurrency 5).
2. **Child folders** — after a folder opens, its subfolders are prefetched into the same cache.

Prefetch never updates the Attachments UI; it only fills the cache. A later click often hits cache on first try.

### 4.3 All Files library listings

| Load | Cache key | Notes |
|---|---|---|
| One library root (`loadFiles`) | `files:lib:...` | Instant revisit |
| All libraries roots | `files:all-root` | Progressive batches; `force` invalidates |
| Deep attachment search | `files:deep:...` | Per library |

### 4.4 HR Files people list (progressive)

For users who can list the full HRPersonal root:

- Old: wait for full root listing (`$top=999`, all pages).
- Now: pages of **200** folders; list paints after **page 1**; later pages append.

---

## 5. HR Files Comments correctness (related tasks)

When opening an HR person folder, Comments loads in two passes:

### Pass 1 — Seed submissions

Tasks submitted/created by that folder person (LookupId filters when possible).

### Pass 2 — Same ID + same task name

Prefer seeds submitted by the **logged-in user** when present; otherwise use the folder person’s seeds.

Include other workflow steps (e.g. manager Finalised) only when they match:

1. **Same eForm ID** (`eFormListId` / `ID=9699` in text), **and**
2. **Same task/form name** after normalizing titles  
   (e.g. `Please Approve Missing Punch…` and `Missing Punch Finalised` → `missingpunch`)

Unrelated forms that only mention the same ID are excluded.

Details of Live LookupId / 429 / 422 behaviour: [`HR-FILES-LIVE-TASK-LOADING-FIX.md`](./HR-FILES-LIVE-TASK-LOADING-FIX.md).

---

## 6. Main code locations

| Area | File / symbol |
|---|---|
| Cache service | `src/app/services/file-crawl.service.ts` |
| Task + folder cache / prefetch | `AppComponent` in `src/app/app.ts` |
| All Files listing cache + prefetch emit | `src/app/components/All-files/all-files.ts` |
| Prefetch binding | `src/app/app.html` → `(foldersForPrefetch)` |
| Concurrency / page limits | `src/app/app.constants.ts` (`hrFilesTaskList*`, `hrTasksFastLoad*`) |
| Graph retry / Prefer header | `src/app/microsoft-graph.ts` |

---

## 7. What to expect when testing

| Action | Expected |
|---|---|
| Open All Files, wait ~1s, click a visible folder | Attachments often instant (prefetch) |
| Open same folder again within 15 min | Attachments + Comments from cache |
| Return to HR Files after To Do / All Files | Last person’s Comments restore from cache |
| First Comments open for a library folder | Spinner clears after first page; list grows |
| First HR person Comments | Spinner clears after first list batch; siblings merge later |
| Manual refresh | Fresh Graph crawl; cache rebuilt |
| Full browser reload | In-memory cache empty; first load is cold again |

---

## 8. Limits / non-goals

- Cache is **session memory only** (not across refreshes/devices).
- Prefetch is best-effort; failures are silent and the click path retries.
- Progressive load may briefly show empty Comments until a matching page/batch arrives.
- Deep full-library crawl on login is intentionally **not** done (avoids Live throttling).
