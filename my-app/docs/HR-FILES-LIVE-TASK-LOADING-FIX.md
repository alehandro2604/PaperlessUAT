# HR Files Comments / Task Loading Fix (UAT vs Live)

This document explains why HR Files Comments worked differently on SharePoint **UAT** vs **Live**, what was broken, and how the fix works.

---

## 1. Difference between UAT and Live

| Area | UAT | Live (`PaperlessLive`) |
|---|---|---|
| Site | Smaller UAT SharePoint site | `sites/PaperlessLive` (see `sharepoint.config.ts`) |
| Task lists | Fewer lists, fewer items | Many eForm/HR task lists, large item counts |
| Attachments | Fast (one folder `children` call) | Also fast — same pattern |
| Comments / tasks for a person | Often “worked” by accident | Slow, incomplete, or empty |
| Graph `$orderby=createdDateTime` | Sometimes tolerated | Rejected with **422 Unprocessable Entity** |
| Graph load volume | Low enough not to throttle | Easy to hit **429 Too Many Requests** |

**Important:** Attachments and Comments use different APIs.

- **Attachments** = drive folder children (`/drives/.../items/{id}/children`) — one cheap call.
- **Comments (HR Files person)** = SharePoint **list items** across many task lists (`/lists/{id}/items?...`).

So attachments looking “instant” while Comments looked broken did **not** mean the library name was wrong. It meant Comments was doing much heavier work.

Current Live config:

```ts
sitePath: 'sites/PaperlessLive'
hrPersonalListDisplayName: 'HRPersonal'
```

(`src/app/sharepoint.config.ts`)

---

## 2. Why it didn’t work

### 2.1 Old approach (broken on Live)

When you opened a person in **HR Files**, the app roughly did:

1. Call Graph for many task lists in parallel.
2. Download only a **small page slice** of each list (for example first 100 items).
3. On the client, keep items whose submitter “looked like” the folder name.

That worked on UAT because lists were small — the person’s tasks were often inside that first page.

On Live it failed because:

1. **Lists are huge** — the person’s June tasks may not be in the first page.
2. **Cannot sort newest-first** — `$orderby=createdDateTime desc` returns **422** on these Live lists.
3. **Too many parallel calls** — scanning dozens of lists caused **429** throttling.
4. **Submitter-only matching** — a later workflow step (manager “Finalised”) is submitted by the manager, not by you, so it was dropped even when it belongs to the **same eForm ID**.

### 2.2 Real-world example (Missing Punch ID 9699)

Two related SharePoint task rows:

| Row | Submitted by | Meaning |
|---|---|---|
| Original | Alehandro | “Please Approve Missing Punch…” — `ID=9699` |
| Later step | Klaus (manager) | “Missing Punch … Finalised” — same `ID: 9699` |

Alehandro must see **both** in his HR Files Comments:

- his own submission, and  
- the manager’s Finalised step for the **same eForm ID**.

Old logic only kept “submitted by Alehandro”, so Klaus’s Finalised row disappeared.

### 2.3 What was *not* the root cause

It was **not** mainly an `HRPersonal` naming mismatch between UAT and Live.

Network evidence showed successful `/lists/.../items` calls (and later 422/429), not “drive not found” failures for attachments.

---

## 3. The fix (high level)

HR Files Comments now loads in **two controlled passes**:

### Pass 1 — Your own submissions

1. Read the email from the HR folder name (e.g. `12345 you@enemalta.com`).
2. Resolve that email to a SharePoint **LookupId** (site-local user id).
3. Query each task list with filters like:
   - `fields/RequestorLookupId eq '{id}'`
   - `fields/AuthorLookupId eq '{id}'`
   - `fields/SubmittedByLookupId eq '{id}'`
4. That returns **your** submitted/created tasks for that list (full history for that person field, not a random page).

### Pass 2 — Same eForm / same task chain

1. From Pass 1 results, collect eForm keys:
   - `eFormListId` field
   - numeric item id when relevant
   - IDs parsed from comment text (`ID=9699`, `ID: 9699`)
2. On lists where you already had hits, fetch siblings that reference those same keys.
3. Merge them into Comments.

Result for ID 9699:

- Alehandro’s original request ✅  
- Klaus’s Finalised step for the same ID ✅  
- Unrelated tasks that only mention a name ❌ (not included)

### Extra stability

- **No `$orderby=createdDateTime`** on these lists (avoids 422).
- **Lower concurrency** (3 parallel list workers).
- **`graphGetWithRetry`** retries on HTTP 429 with backoff.
- UI updates as batches finish so Comments does not stay empty forever.

---

## 4. Files that needed updates

| File | Role |
|---|---|
| `src/app/app.ts` | Main HR Files Comments load logic, LookupId + eForm sibling matching |
| `src/app/microsoft-graph.ts` | Extra headers on Graph GET + 429 retry helper |
| `src/app/app.constants.ts` | Concurrency / page limits for HR Files loads |
| `src/app/services/form-configuration.service.ts` | Optional helper `getHrTaskListNames()` (HR workflow lists) |
| `src/app/sharepoint.config.ts` | Live site target (`PaperlessLive`, `HRPersonal`) — environment config, not the core algorithm |

---

## 5. Why each update

### `app.ts`

- Replace “download random pages + name match” with **person LookupId filters**.
- Add **same-eForm sibling** expansion so manager Finalised steps appear.
- Keep scope strict: seed = logged-in folder person’s submissions; extras = same eForm ID only.
- Improve person-field stringification (`Email` / `DisplayName`) for matching fallbacks.
- Avoid Live `$orderby` 422s; remember when orderby is unsupported.

### `microsoft-graph.ts`

- Allow `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly` (needed for some LookupId filters).
- Retry throttled Graph calls instead of failing the whole Comments load.

### `app.constants.ts`

- Cap parallel HR Files list traffic to reduce 429s.
- Keep fallback page scans small when LookupId filtering is unavailable.

### `form-configuration.service.ts`

- Expose HR-only task list names if needed for narrower queries later.

---

## 6. Where the new / updated code is (line map)

> Line numbers refer to the codebase at the time this doc was written. If files shift, search for the symbol names below.

### `src/app/microsoft-graph.ts`

| Lines | What |
|---|---|
| ~15–27 | `graphGet(..., extraHeaders)` — supports Prefer header |
| ~30–55 | `graphGetWithRetry` — retries HTTP 429 with backoff |

### `src/app/app.constants.ts`

| Lines | What |
|---|---|
| ~32–37 | `hrFilesTaskListConcurrency`, `hrFilesTaskListPageLimit`, `hrFilesTaskListBackfillPageLimit` |

### `src/app/app.ts`

| Lines | What |
|---|---|
| ~12 | Import `graphGetWithRetry` |
| ~588–590 | Cache: User Information List id + email→LookupId map |
| ~2609–2627 | `stringifyTaskFieldValue` — reads `Email` / `DisplayName` from person objects |
| ~2635–2655 | `getHrPersonalFolderMatchTokens` — tokens from folder name/email/employee id |
| ~2690–2750 | `loadHrPersonalTasksForFolder` — entry point when an HR Files person is opened |
| ~2754–2880 | `collectHrPersonalTasksFromLists` — Pass 1 (person) + Pass 2 (same eForm ID) |
| ~2889–2925 | `collectEFormKeysFromMappedTasks` / `extractEFormKeysFromMappedItem` |
| ~2933–2970 | `fetchSharePointListItemsForEFormKeys` — Graph filter by eForm id fields |
| ~2975–2988 | `doesRawItemReferenceEFormKeys` — also matches `ID=9699` in comment text |
| ~2992–2994 | `getEmailFromHrFolderName` |
| ~2997–3060 | `resolveSharePointUserLookupId` — email → SharePoint LookupId |
| ~3066–3110 | `fetchSharePointListItemsForPersonLookup` — Requestor/Author/SubmittedBy LookupId filters |
| ~3199–3235 | `fetchSharePointListPages` — no Live orderby; remembers if orderby is unsupported |

### `src/app/services/form-configuration.service.ts`

| Lines | What |
|---|---|
| ~184–191 | `getHrTaskListNames()` — HR workflow task list names from eForms |

### `src/app/sharepoint.config.ts`

| Lines | What |
|---|---|
| ~9–13 | Live site path / site id / `HRPersonal` library display name |

---

## 7. Code explained in plain English

### 7.1 Opening a person in HR Files

`loadHrPersonalTasksForFolder` runs when you click a person folder.

In plain English:

1. Get a Graph token.
2. Make sure eForms config and site/list metadata are loaded.
3. Ask `collectHrPersonalTasksFromLists` to gather tasks.
4. While that runs, show partial results in Comments as soon as batches arrive.
5. When finished, set `commentItems` and turn off the loading spinner.

### 7.2 Collecting tasks (`collectHrPersonalTasksFromLists`)

**Step A — find who this folder is**

- Pull email from folder name.
- Resolve SharePoint LookupId for that email.

**Step B — Pass 1 (your submissions)**

For each task list (in small batches of 3):

- Prefer Graph filters: “items where Requestor/Author/SubmittedBy LookupId = this person”.
- If those filters are unavailable, fall back to downloading a few pages and matching submitter locally.
- Keep non-comment eForm tasks only.

**Step C — Pass 2 (same eForm / same task)**

- From Pass 1 tasks, collect IDs like `9699`.
- On lists that already had your tasks, fetch other items that reference those IDs (fields or comment text).
- Merge them in.

That is how Klaus’s Finalised row appears next to your original Missing Punch.

### 7.3 Resolving LookupId (`resolveSharePointUserLookupId`)

SharePoint person columns do not filter well by email directly.

So the app:

1. Finds the site’s **User Information List**.
2. Looks up the user row for that email.
3. Uses that row’s numeric Id as `RequestorLookupId` (etc.).

Results are cached so the same email is not resolved repeatedly.

### 7.4 Person filter (`fetchSharePointListItemsForPersonLookup`)

For one list and one LookupId, it asks Graph:

- Give me items where this person is Requestor, or Author, or SubmittedBy.
- Page through all matching results.
- On 429, retry with delay.

If none of those filters work on a list, it returns `null` so the caller can fall back to page scanning.

### 7.5 Same-task siblings (`fetchSharePointListItemsForEFormKeys` + `doesRawItemReferenceEFormKeys`)

After we know eForm keys from **your** tasks:

- Try Graph filters on `eFormListId` / `ListId` / `field_11`.
- Also accept items whose comment/title contains `ID=9699` or `ID: 9699`.

Only items tied to those keys are added. This keeps Comments limited to:

1. tasks you submitted, and  
2. workflow siblings of those same eForms.

### 7.6 Why attachments stayed fast

`loadDriveFolderContents` still does one Graph call:

`/drives/{driveId}/items/{folderId}/children`

That path was never the bottleneck. The Comments path (many list queries + matching) was.

### 7.7 422 and 429 handling

- **422** from `$orderby=createdDateTime` → do not use orderby on these lists.
- **429** from too many Graph calls → lower concurrency + `graphGetWithRetry`.

---

## 8. End-to-end flow (diagram)

```text
User opens HR Files person folder
        │
        ▼
Extract email from folder name
        │
        ▼
Resolve SharePoint LookupId
        │
        ▼
Pass 1: filter each task list by Requestor/Author/SubmittedBy LookupId
        │
        ▼
Collect eForm IDs from those results (e.g. 9699)
        │
        ▼
Pass 2: fetch siblings with same eForm ID (manager Finalised, etc.)
        │
        ▼
Show merged tasks in Comments
```

---

## 9. How to verify

1. Point config at Live (`PaperlessLive`).
2. Sign in and open **your** HR Files folder.
3. Confirm:
   - Your submitted tasks appear.
   - Related Finalised/manager steps for the **same ID** also appear.
   - Unrelated other people’s tasks do not appear.
4. In Network tab, expect:
   - `items?$filter=fields/RequestorLookupId eq '...'` (and similar)
   - optionally `items?$filter=fields/eFormListId eq '9699'`
   - few or no `$orderby=createdDateTime` 422s
   - 429s should retry and recover rather than leave Comments empty

---

## 10. Summary

| Problem | Cause | Fix |
|---|---|---|
| Comments slow / incomplete on Live | Page-slice scan of huge lists | LookupId person filters |
| Empty after orderby attempts | Live rejects `$orderby=createdDateTime` (422) | Stop using orderby |
| Console 429 floods | Too many parallel Graph calls | Concurrency 3 + retry helper |
| Manager Finalised missing | Only “submitted by me” kept | Second pass by same eForm ID |
| Attachments fine, Comments not | Different APIs | Expected; fixed Comments path |

The durable rule for HR Files Comments:

> Show the folder person’s own submissions, plus other workflow tasks that share the same eForm ID.
