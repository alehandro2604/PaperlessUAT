## PaperlessApp (Angular)

PaperlessApp is an Angular app that connects to SharePoint/Microsoft Graph to:

- Show **HR tasks assigned to the logged-in user** (To Do)
- Show **HR tasks assigned to your subordinates** (Manager view)
- Browse **HR document libraries / personal folders** (All Files)
- Optionally **delegate** a task to another domain user (permission-based)

This folder is the Angular project: `my-app/`.

### Prerequisites

- **Node.js + npm** (see `package.json`)
- Access to the SharePoint site + task lists used by the app
- Azure AD app registration configured for Graph/SharePoint access (MSAL)

### Quick start (Windows / PowerShell)

From `my-app/`:

```powershell
npm install
npm run start
```

Then open the URL printed by Angular (typically `http://localhost:4200`).

### Configuration

- **SharePoint**: `src/app/sharepoint.config.ts`
- **Auth (MSAL / tokens)**: `src/app/services/auth.service.ts`

### Common scripts

- `npm run start`: run dev server
- `npm run build`: production build
- `npm run test`: unit tests

---

## Major feature: Task Delegation

Authorised users can reassign an HR task to another domain user. The new
assignee is written back to SharePoint so the change **persists across reloads**
— not just in the browser.

### User flow

1. A user who appears in the **`UsersWhoCanDelegate`** list sees a delegate
   button on pending tasks assigned to them.
2. They search for a colleague (sourced from the **`AllDomainUsers`** list).
3. On submit, the app writes the new assignee to the task's `AssignedTo`
   person column in SharePoint and verifies the change before updating the UI.

### Key files

| File | Responsibility |
|------|----------------|
| `src/app/components/delegate-component/delegate.ts` / `.html` | Delegate button, search box, and submit form. Emits the selected task + assignee. |
| `src/app/services/delegate.service.ts` | Core logic: permission checks, user search, and the SharePoint write (`updateAssignedTo`). |
| `src/app/services/auth.service.ts` | Acquires Graph and SharePoint REST write tokens. |
| `src/app/app.ts` | `onDelegateTask()` calls the service, awaits a verified write, then updates the view. |
| `src/app/microsoft-graph.ts` | `graphGet` / `graphPatch` helpers. |

### How the SharePoint write works

SharePoint **Person/Group** columns (like `Assigned To`) do **not** store a name
or email. They store an integer **lookup ID** that points to a row in the site's
hidden **User Information List**. Writing a name or email silently fails — the
API can even return `200 OK` without changing anything.

The delegation save therefore follows this chain:

1. **Resolve site + list** by Graph, and read the assignee **column metadata**
   to detect the real internal field name and whether it allows multiple values.
2. **`ensureuser`** (SharePoint REST `_api/web/ensureuser`) — provisions the new
   assignee into the site if needed and returns their **lookup ID**.
3. **Graph `PATCH`** the item using the correct person-field format:
   - single-value: `{ "AssignedToLookupId": 123 }`
   - multi-value: `{ "AssignedTo@odata.type": "Collection(Edm.Int32)", "AssignedToLookupId": [123] }`
4. **SharePoint REST `MERGE` fallback** (`AssignedToId: 123`) if Graph still
   won't apply the change.
5. **Read-back verification** — the write is only treated as successful when the
   item actually reflects the **new** assignee (not just any non-empty value, and
   not the previous assignee). The UI updates only after this passes.

> Glossary: **LookupId** is the user's numeric `ID` in the hidden
> *User Information List*. **`AssignedToLookupId`** (Graph) and **`AssignedToId`**
> (SharePoint REST) are two API names for writing that same number into the
> `AssignedTo` column. Neither appears as a visible column in the list UI.

### Understanding the lookup ID (`spUserId`)

A **lookup ID** is the integer that uniquely identifies a user *within a single
SharePoint site*. In this codebase it is the variable `spUserId`.

**What it is**
- It is the `ID` (row number) of the user in the site's hidden
  **User Information List**, e.g. user `Adrian Hedley` → `ID = 123`.
- A Person/Group column such as `AssignedTo` stores **only this number**, not a
  name or email. The displayed name is looked up from that list using the ID —
  hence "lookup ID".

**Where to get it**
- Primary: **`ensureuser`** (SharePoint REST `_api/web/ensureuser`). You pass the
  user's login claim (`i:0#.f|membership|<email>`); SharePoint returns their `Id`,
  creating the User Information List row first if it doesn't exist yet.
- Fallback: **scan the User Information List** via Graph and match on `EMail` /
  `Title` (`resolveSharePointUserId` in `delegate.service.ts`).
- You can also view it manually at
  `https://<tenant>.sharepoint.com/sites/<site>/_catalogs/users/simple.aspx`.

**Why and how it is used**
- It is the *only* value a person column accepts. Once we have `spUserId`, we
  write it to the task item:
  - Graph: `{ "AssignedToLookupId": spUserId }` (or an array for multi-value).
  - SharePoint REST fallback: `{ "AssignedToId": spUserId }`.
- The ID is **per-site**: the same person has different IDs on different sites,
  so it must always be resolved against the task's own site.

### Why the earlier `spUserId` handling was wrong

Getting `spUserId` was the right idea, but the old approach was unreliable:

1. **It only *read* the User Information List, never *ensured* it.** If the new
   assignee had never accessed the site, they had no row, the lookup returned no
   match, and `spUserId` ended up `null`.
2. **A `null` ID silently disabled the correct writes.** The lookup-ID PATCH
   formats were skipped whenever `spUserId` was missing, leaving only an
   unreliable login-name/claims string write.
3. **The lookup query itself was fragile** — it used an unsupported `$filter` and
   referenced the list by its display name (with spaces) in the URL, which
   returned `400 Bad Request`, and the error was swallowed as "no ID".
4. **Wrong field key when an ID was found.** Multi-value writes used
   `AssignedToLookupIds` (plural), which is not a valid field. The correct key is
   `AssignedToLookupId` (singular), passed an **array** for multi-value columns.

**The fix:** resolve `spUserId` with `ensureuser` first (guaranteeing a valid ID,
even for first-time users), fall back to a robust paginated User Information List
scan resolved by **GUID**, and write it with the correct `AssignedToLookupId` /
`AssignedToId` formats.

### Why it used to "save" but revert on refresh (fixed)

The earlier version appeared to delegate successfully, but the task reverted to
the original assignee after a page reload. This was caused by the change living
**only in the browser's memory**, never reaching SharePoint:

1. On submit, the app updated the local task object
   (`commentItems[i].assignedTo`) and re-rendered. The UI instantly showed the new assignee, so it looked done.
2. But the actual SharePoint write either **never happened** (the service method was an unimplemented stub) or **silently failed** (wrong person-field format —
   Graph returned `200 OK` without changing the value).
3. The success check was too loose: it accepted *any* non-empty `AssignedTo`
   value as proof, so the **old** assignee still being there counted as success.
4. On refresh, the app **re-fetches tasks from SharePoint**, which still held the
   original assignee — so the in-memory change was overwritten and "disappeared".

In short: the screen was updated from a local variable, not from a confirmed
server write. The browser and SharePoint were out of sync.

**The fix** makes the UI update *follow* the server instead of leading it:

- `onDelegateTask()` in `app.ts` now **awaits** `delegateService.updateAssignedTo()`
  and only updates the view if it resolves successfully.
- The write is **verified by reading the item back** and confirming it shows the
  **new** assignee (see step 5 above) before success is reported.
- If the write cannot be verified, it **throws** and the UI shows a failure
  modal instead of a fake success — so memory and SharePoint can no longer drift
  apart.

### Required permissions

- Microsoft Graph delegated: `Sites.ReadWrite.All` (and `Sites.Read.All`, `User.Read`)
- SharePoint REST: a token for `https://<tenant>.sharepoint.com/.default`, used by
  `ensureuser` and the REST `MERGE` fallback.

### Relevant SharePoint lists

| List | Purpose |
|------|---------|
| HR task lists (e.g. `HRTaskMissingPunch`, `HRTaskTelework`, `ProcTasks1`) | Hold the tasks and their `AssignedTo` column. |
| `UsersWhoCanDelegate` | Controls who is allowed to delegate. |
| `AllDomainUsers` | Source for the assignee search. |
| `User Information List` (hidden, system) | Maps lookup IDs to real users. |

## Subordinate Tasks (Manager View)

Managers can tick **Subordinate Tasks** in the To Do panel to see pending HR tasks
**assigned to employees who report to them**, grouped by employee name. Only
employees with at least one active task are listed (not all direct reports).

Data comes from:

- **`AllDomainUsers`** — resolves who reports to the logged-in manager (`Manager` /
  `ManagerEmail` columns).
- **HR task lists** — tasks whose assignee matches one of those subordinates.

### User flow

1. Open **To Do** and check **Subordinate Tasks**.
2. The app resolves subordinates from the manager's login (email, UPN, display name).
3. It reloads HR tasks in **subordinate-only** mode (newest first).
4. The list shows each subordinate who has pending / in-progress / rejected tasks.
5. Clicking an employee filters the **Comments** panel to that person's assigned tasks.

### Key files

| File | Responsibility |
|------|----------------|
| `src/app/components/todo-list/todo-list.component.ts` / `.html` | Subordinate Tasks checkbox, employee grouping, empty states. |
| `src/app/services/subordinaryTask.service.ts` | Loads `AllDomainUsers`, builds manager search keys, caches subordinates, matches assignees. |
| `src/app/services/user.service.ts` | Shared assignee matching (name, email, email local part). |
| `src/app/app.ts` | `reloadTodoTasksWithSubordinates()`, `subordinateTasksOnly` fetch mode, assignee field collection. |
| `src/app/app.constants.ts` | `hrTasksSubordinateLoadPageLimit` (scans more list pages for subordinate view). |

### Why it did not work before

Several separate bugs stacked on top of each other. Fixing only one was not enough.

#### 1. Manager lookup used the wrong identifiers

SharePoint stores the manager in `AllDomainUsers.Manager` as a **short name**
(e.g. `alehandro` inside `Klaus Paul Camilleri`). The app only searched with the
full login email or display name, so **zero subordinates** were found automatically
—even though typing `alehandro` in the old debug search found 4 employees.

#### 2. Subordinate tasks were never loaded from SharePoint

HR tasks were filtered in `mapSharePointItemToHrTask()` to rows where the
**logged-in user** was submitter, assignee, or Superior 1/2. Tasks assigned to
**employees under the manager** were discarded before they reached the To Do list.

#### 3. The fetch returned oldest tasks first

List items were loaded without `$orderby`, so Graph returned the **oldest** pages
first (often only 100–500 items). A **newly assigned test task** at the top of
the list was never fetched, so the UI showed *"No pending tasks assigned to your
4 subordinate(s)"* even though subordinates were found correctly.

#### 4. Assignee matching was too narrow

Matching relied on a single `AssignedTo` value and sometimes fell back to
`field_7`, which on some HR lists is **Time In** (e.g. `09:00`), not a person.
SharePoint also stores assignees as full name, email, or email username only —
strict name matching missed valid rows.

#### 5. Stale empty cache

If the first subordinate lookup returned `[]`, that empty array was cached and
never refreshed when the user toggled the checkbox again.

### The major fixes

#### Fix A — Resolve the manager like SharePoint does (`buildManagerSearchKeys`)

Derive all useful lookup keys from the login, especially the **email local part**
(the same token that appears in `Manager`):

```ts
// subordinaryTask.service.ts
async buildManagerSearchKeys(rawIdentifiers: string[]): Promise<string[]> {
  const keys = new Set<string>();

  for (const raw of rawIdentifiers) {
    keys.add(trimmed);
    if (trimmed.includes('@')) {
      const localPart = trimmed.split('@')[0]?.trim();
      if (localPart) keys.add(localPart.toLowerCase());
    }
  }

  // Also enrich from the manager's own AllDomainUsers row (FullName, ADmail, …)
  const self = users.find(user => /* match by email / local part / name */);
  if (self?.FullName) keys.add(self.FullName);
  if (self.ADmail.includes('@')) keys.add(self.ADmail.split('@')[0]);

  return [...keys];
}
```

`ensureSubordinatesForManager()` searches with **every** key and deduplicates
results. Searching `alehandro@…` now also tries `alehandro`, matching the
SharePoint `Manager` column.

#### Fix B — Dedicated subordinate task reload (`subordinateTasksOnly`)

Checking **Subordinate Tasks** triggers a separate load that keeps only tasks
assigned to cached subordinates:

```ts
// app.ts — on checkbox on
await this.subordinaryTaskService.ensureSubordinatesForManager(identifiers, true);
await this.loadHrTasksFromLists(false, { createdSinceIso: sinceIso }, {
  updateTodo: true,
  subordinateTasksOnly: true,
  fastLoadPageLimit: AppConstants.hrTasksSubordinateLoadPageLimit, // 5 pages
});
```

Inside the mapper:

```ts
// app.ts — mapSharePointItemToHrTask()
const isAssignedToSubordinate =
  this.subordinaryTaskService.isTaskAssignedToAnyAssigneeValue(assigneeValues);

if (subordinateTasksOnly) {
  if (!isAssignedToSubordinate) return null;
}
```

Unchecking the box calls `loadTodoTasksForCurrentUser()` again so the normal
"tasks assigned to me" view is restored.

#### Fix C — Fetch newest tasks first

Subordinate mode orders list items by `createdDateTime desc` so recent assignments
(including test tasks) appear in the first page:

```ts
// app.ts — fetchSharePointListPages()
const orderQuery = newestFirst ? '&$orderby=createdDateTime desc' : '';
let nextPath =
  `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${top}${orderQuery}`;
```

#### Fix D — Collect every assignee field and match flexibly

All person-like columns are read; `field_7` is ignored when it looks like a time:

```ts
// app.ts — collectTaskAssigneeValues()
private collectTaskAssigneeValues(fields: Record<string, any>): string[] {
  const candidates = [
    fields['AssignedTo'],
    fields['Assigned'],
    fields['AssignedTo0'],
    fields['CurrentAssignee'],
    fields['TaskAssignee'],
  ];

  const field7 = String(fields['field_7'] ?? '').trim();
  if (field7 && !/^\d{1,2}:\d{2}/.test(field7)) {
    candidates.push(fields['field_7']);
  }

  return [...new Set(candidates.map(v => this.extractPersonName(v)).filter(Boolean))];
}
```

`SubordinaryTaskService.isTaskAssignedToEmployee()` uses `UserService.matchesAssigneeField`
**plus** normalized token overlap on name, email, PIN, and email local part.

`UserService` also matches the **email local part** explicitly:

```ts
// user.service.ts — matchesSingleAssignee()
const emailLocal = email.includes('@') ? email.split('@')[0].toLowerCase() : email;
if (emailLocal.length > 2 && assignedToLower.includes(emailLocal)) return true;
```

#### Fix E — UI shows employees with tasks only

`tasksAsSuperior` no longer pre-populates every subordinate with an empty group.
It groups only tasks that pass assignee + status checks, so the list shows
**employees who actually have work**, not all four names with zero counts.

### Relevant SharePoint lists

| List | Purpose |
|------|---------|
| `AllDomainUsers` | Employee hierarchy (`Manager`, `ManagerEmail`, `ADmail`, `FullName`). |
| HR task lists (e.g. `HRTaskMissingPunch`, `HRTaskTelework`) | Hold tasks and `AssignedTo` person columns. |

### Note on the removed debug panel

An earlier **Manager Employee Search** test panel (`ManagerTestComponent`) was
used to debug manager-name lookups. It has been **removed from the UI**; subordinate
resolution is fully automatic from the logged-in user. The component file remains
in the repo for ad-hoc testing if needed.

## HRPersonal Permission-Based Folder View Fix

The **HR Files** panel now follows the user's real SharePoint permissions for the
`HRPersonal` document library.

### What was the problem

The app originally loaded only one personal folder: the folder that matched the
logged-in user. That worked for normal users, but it was wrong for users who have
SharePoint permission to see all folders inside `HRPersonal`.

The old logic was effectively:

```ts
const userFolder = await this.findUserFolder(targetDriveId, token);
if (!userFolder?.id) {
  // Stop if the logged-in user's personal folder cannot be found.
  return;
}

this.userRootFolderId = userFolder.id;
this.hrPersonalFolderName = userFolder.name;

const driveLoadResult = await this.getAllDriveItems(
  targetDriveId,
  token,
  userFolder.id
);

this.applyDriveItems(driveLoadResult.items, userFolder.id);
```

That meant the app always forced the user into **one matched personal folder**,
even when SharePoint allowed that user to see more folders.

After changing the app to show all accessible folders, another issue appeared:
when clicking a different user's folder, the comments/tasks panel could still
show tasks for the logged-in user or unrelated tasks from other users.

This happened because:

1. The folder list and the attachment/task panels were sharing too much state.
2. Clicking an `HRPersonal` folder initially used the normal `loadHrTasks()` flow,
   which loads tasks for the **logged-in user**, not the selected folder.
3. The task matching was too broad. It checked fields like title, path, comments,
   or shared domain text, so unrelated tasks could accidentally match.
4. The silent background refresh could merge the logged-in user's tasks back into
   the comments panel while another user's folder was open.

### Earlier folder-loading fix

The `HRPersonal` flow was first split into two clear steps:

1. **Load folders first**
   - The app checks whether the signed-in user can list the `HRPersonal` root.
   - If yes, it loads the top-level folders returned by SharePoint.
   - If no, it falls back to the logged-in user's own folder.
   - It does **not** recursively load all folder contents at this stage.

2. **Load contents only when a folder is clicked**
   - Clicking a folder uses the existing attachment loader
     (`loadDriveFolderContents`) so files still appear in the attachments pane.
   - It also loads comments/tasks specifically for the selected folder/user.
   - If that selected user has no tasks, the comments panel is cleared and shows
     `No tasks found for <folder>`.

This fixed the original "only one folder loads" behavior and kept folder
contents loading on demand. However, it still trusted Graph root listing too
much, so the final access decision below was added.

### Final access-control fix

The working fix is **not** to guess permissions from Graph root listing.

Graph was still returning all `HRPersonal` folder names for users who should only
see their own folder. So checks like "can this user list the HRPersonal root?"
were not safe enough. The app now uses an explicit allow-list for users who are
allowed to browse all `HRPersonal` folders.

The wrong code was effectively:

```ts
if (await this.canListHrPersonalRoot(targetDriveId, token)) {
  this.hrPersonalRootFolders = await this.fetchHrPersonalRootFolders(
    targetDriveId,
    token
  );
  return;
}
```

The mistake was assuming this:

```ts
can list HRPersonal root = can view all HRPersonal folders
```

That assumption was false. In this SharePoint setup, a restricted user could
still receive the root folder list from Graph, so the app displayed everyone even
when the UI should only show that user's own folder.

In simple terms:

- Users in `hrPersonalAllAccessUsers` see **all** top-level folders.
- Users not in `hrPersonalAllAccessUsers` see **only their own** personal folder.
- Folder contents are still loaded only when a folder is clicked.
- Clicking a folder still uses the existing attachment loader
  (`loadDriveFolderContents`), so the attachment pane behavior stays the same.
- If the selected user has no tasks, the comments panel shows
  `No tasks found for <folder>`.

The allow-list lives in `src/app/sharepoint.config.ts`:

```ts
hrPersonalAllAccessUsers: [
  'alehandro-joaquin.camilleri@enemalta.com.mt',
] as string[],
```

The decision in `app.ts` is now effectively:

```ts
if (this.isConfiguredHrPersonalAllAccessUser()) {
  // Alehandro / configured HR users: show all HRPersonal folders.
  this.hrPersonalRootFolders = await this.fetchHrPersonalRootFolders(
    targetDriveId,
    token
  );
  return;
}

// Everyone else: show only the logged-in user's own folder.
const userFolder = await this.findUserFolder(targetDriveId, token);
this.hrPersonalRootFolders = [
  {
    ...userFolder,
    webUrl: userFolder.webUrl ?? '',
    isFolder: true,
  },
];
```

And the helper checks the logged-in user's email/UPN:

```ts
private isConfiguredHrPersonalAllAccessUser(): boolean {
  const allowedUsers = sharePointConfig.hrPersonalAllAccessUsers ?? [];
  const userEmail = (this.currentUser?.email ?? '').toLowerCase();
  const userUpn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();

  return allowedUsers.some(email => {
    const normalized = email.toLowerCase();
    return normalized === userEmail || normalized === userUpn;
  });
}
```

Task matching for selected `HRPersonal` folders is now stricter:

- It only checks identity/assignee-style fields such as `AssignedTo`,
  `Requestor`, `SubmittedBy`, `EmployeeName`, `Email`, and `ADmail`.
- It no longer matches broad fields like comments, paths, or generic text that
  can cause false positives.
- Background refresh no longer merges the logged-in user's tasks into the
  comments panel while a specific `HRPersonal` folder is selected.

In short: the app no longer trusts Graph root listing to decide who can see all
folders. It uses an explicit all-access list for broad HRPersonal browsing, and
everyone else is limited to their own folder in the UI.

### Version
6.5.2