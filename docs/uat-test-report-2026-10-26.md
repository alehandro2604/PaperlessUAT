# UAT test report

**Date:** 2026-10-26  **Commit:** `0f797c1` (plus an uncommitted wording change in `todo-list.component.html`)
**App:** `npm run start:uat` on http://localhost:4201 (PaperlessUAT site), backend `dev:uat` on :3001, Redis db 1
**Account:** Alehandro-Joaquin Camilleri. Read-only testing: nothing was submitted, completed, created or deleted.

Compared against the previous run (`whole-app-test-report.md`, 14 Oct, live).

## 1. Summary

| Area | Result | Numbers |
|---|---|---|
| Unit tests | Pass | 6/6 tests, 3 files. One teardown error (see 2.7) |
| Production build (`build:uat`) | Pass | 1.23 MB (255 KB gzipped). Targets UAT URL and site |
| Backend syntax | Pass | All 9 files |
| Backend security | Pass | Every route except `/api/health` returns 401 without a valid token; bad token → 401 "jwt malformed"; unknown origins get no CORS header; a 30 MB unauthenticated upload is rejected in 0.1 s |
| Backend speed | Pass | Health: 9 ms average; 50 at once: 8 ms average, all OK |
| Sign-in | Pass | Microsoft sign-in, back to app, header says "Paperless System UAT" |
| First load (empty cache) | Pass | Page ready 1.9 s. Background To Do scan done in ~20 s (125 requests, 0 errors), 1,637 tasks cached |
| Reload (cached) | Pass | Page ready 0.2 s |
| To Do search | Pass | Case-insensitive, trims spaces, special characters shown as text, clears correctly |
| Expand a group | Pass | 0 requests |
| Open a task | **Bug** | 0 requests (was 38 requests / ~38 s), but the Comments panel shows the wrong items (2.1) |
| Comments search | Pass | Wrong empty message (2.3) |
| Attachments tab | Pass | Label names the wrong person (2.4) |
| Subordinate Tasks | Pass | 9.9 s, 25 requests (was ~55 s, 77 requests) |
| Future Tasks | Pass | 1.5 s; cutoff today + 14 days is correct |
| Refresh | Pass | ~2 s |
| Load more tasks | Pass | 14 s, 26 requests, button hides when nothing is left |
| + New form dialog | Pass | 15 HR forms, WorkFlows list, × closes. Forms now open in a new tab (the old "leaves the app" issue is fixed). Escape does not close it |
| HR tab | Pass | 8 s, 43 requests, 98 shown of ~930 |
| HR search | Pass | Searches all ~930 people by name or staff number, 0 requests |
| HR person folder | Slow | 78 requests / 9 s to show 3 folders (2.5) |
| All Files search | Pass | First search 12.6 s (73 requests), then 0.4 s. Empty result message is clear |
| All Files folder | Pass | 2 s, 9 requests; its comments are correctly limited to that folder |
| Links | Pass | All point to `sites/PaperlessUAT` |
| Mobile (375 × 812) | Pass, 2 issues | No sideways scroll, no tiny text, no overflow on any tab |
| Desktop (1280 × 800) | Pass, 1 issue | 3 columns, no sideways scroll |
| Errors | None | 0 console errors, 0 HTTP 4xx/5xx, 0 throttling (429) all session |

## 2. Findings (most important first)

1. **Opening some To Do tasks shows 165 unrelated comments.** Example: Sylvan Abela → "ShiftDev-Test Registry-20". These "Monitor Workflow" tasks have no eForm list ID, so the Comments filter (`comments.store.ts` `filterCommentItems`) matches every comment that also has no ID. You see Thorsten Callus, Lorraine Smith etc., with "Load more (145 remaining)", and none are about the task. All Files filters correctly. Pre-existing, not caused by today's fix.
2. **Bottom menu covers the end of the "Fill out a new form" list on phones** (dialog ends at 812 px, menu starts at 742 px). Still there from the last report.
3. **Comments search with no match says "No HR Tasks found. Click refresh to load HR Tasks."** It should say no comments match the search.
4. **Attachments says "Showing items for Alehandro-Joaquin Camilleri"** (the signed-in user) when the task belongs to Sylvan Abela.
5. **Opening a person's HR folder makes 78 requests (~9 s) to show 3 folders.**
6. **Desktop: names in the To Do column are cut to ~8 letters** ("Alehan…"). The column is 326 px and the name gets 64 px.
7. **Unit test teardown error:** `app.spec.ts` starts the sign-in redirect after the test ends (NG0205 "Injector has already been destroyed"). Tests pass; the spec should mock sign-in.
8. **Accessibility:**
   - The bottom menu items (HR, To Do, All Files) are plain `div`s: keyboard and screen-reader users cannot reach them.
   - The Subordinate and Future checkboxes have no label ("on").
   - The menu icons have no alt text.
   - The "Add new comment" button is 14 × 18 px on desktop.
   - Escape does not close the new form dialog.
9. **HR list starts with a "? 9999" entry** with no name.
10. **Backend:** still sends `X-Powered-By: Express`. On start, the first Redis connection timed out because Redis was still starting; it reconnected by itself.
11. **Wording (uncommitted change):** "No subordinate tasks assigned to you." The tasks belong to your staff, not to you; "No pending tasks for your staff" would be clearer.

## 3. Improved since the last report
- Opening a task: 38 requests / ~38 s → 0 requests, instant.
- Subordinate Tasks: ~55 s / 77 requests → 9.9 s / 25 requests.
- Power Apps forms open in a new tab instead of replacing the app.
- Reload paints from cache in 0.2 s.

## 4. Not covered
Completing, delegating, claiming or commenting on tasks (all write to SharePoint); uploads and downloads; a non-manager or non-HR account; the deployed IIS site and its `/api` proxy; Android/desktop builds; offline and expired-token cases. Today's subordinate "Complete" fix could not be checked end to end because there were no open subordinate tasks.
