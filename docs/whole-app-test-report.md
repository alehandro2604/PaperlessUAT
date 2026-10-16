# Whole-app test report

**Date:** 2026-10-14  **Branch:** `cache-per-task-todo`  **App:** `ng serve --configuration live` on http://localhost:4200 (live tenant, read-only actions), backend on http://localhost:3000
**Scope:** every area reachable from the main UI: sign-in/startup, To Do (list, detail, Subordinate, Future, + new form), HR, All Files (covered in detail in `all-files-error-test-report.md`), account menu, mobile (375 px) and desktop (1280 px) layouts, and page-level quality checks. No source code was changed. Nothing was submitted, created, deleted or signed out.

Method: the same passive in-page logger as the All Files test (requests with status and timing, console errors, uncaught errors, last user action), plus the browser's own network log for cold start. Limits: the pane is small and scales emulated sizes, so visual layout judgements are approximate; no real outages or expired tokens were reproduced in this run; a regular (non-admin) account was not available, so permission differences were not tested.

## 1. Summary

| Area | Works | Notes |
|---|---|---|
| Cold start / sign-in | Yes | All 200/204, cache session started, no console errors |
| To Do list | Yes | Served from cache, 0 extra requests on expand |
| To Do task detail | Yes, **slow** | 38 requests / ~38 s in background; 32 SharePoint list queries |
| Subordinate Tasks | Yes, **very slow** | ~55 s, 77 requests, result "No employees found" |
| Future Tasks | Yes | Loads, "Showing tasks after 28/10/2026" |
| + New form dialog | Yes | Lists HR forms and WorkFlows; bottom nav overlaps the list |
| Form "Flexi Time" | **Leaves the app** | Same tab goes to `apps.powerapps.com` |
| HR tab | Yes | 100 employees, "Load more (750 remaining)", ~39 requests on open |
| All Files | Yes | See separate report |
| Account menu | Yes | "Signed in as … / Log Out" only |
| Desktop layout (1280) | Yes | 3 columns, no horizontal scroll |
| Uncaught JS errors | **None** in the whole session | 0 `error` / `unhandledrejection`, 0 `console.error` outside the induced All Files faults |

No HTTP error (4xx/5xx) occurred naturally anywhere in the app during this run.

## 2. Findings (ranked)

### Performance
1. **Subordinate Tasks: ~55 s and 77 requests for an empty answer.** Ticking "Subordinate Tasks" (action) kept "Loading subordinate tasks…" on screen until about 55 s, then showed "No employees found under your manager account." Every call was 200. It keeps crawling SharePoint lists even though there is nothing to show. A quick "do I have direct reports?" check first would avoid this.
2. **Opening a task fires 32 SharePoint list queries (38 requests, ~38 s).** Each is `lists/{id}/items?$expand=fields&$top=200&$orderby…`, two at a time, 1–8 s each (average ~2 s, max ~8 s). The first comment shows early, but background work keeps running and competes with anything else the user does (it shares the Graph throttle budget with All Files).
3. **HR tab fires ~39 requests on open** (18 `root/children`, 21 SharePoint list queries, 17 backend cache calls) just to show a list of 100 of ~850 employees.
4. **Background To Do traffic while on other tabs** (seen during the All Files test): SharePoint list calls every ~2 s. Worth checking it pauses when the tab is not visible.

### UX / behaviour
5. **"Flexi Time" (and probably other Power Apps entries) navigates the whole tab away** to `apps.powerapps.com/mobile/redirect…` ("Open the app / Continue in browser"). The user loses the app's state and has no obvious way back except the browser Back button. This may be the mobile-emulation redirect; check on a real phone and consider `target="_blank"`.
6. **Bottom navigation overlaps the "Fill out a new form" dialog** on mobile (dialog bottom = 812 px, nav top = 742 px, so the last ~70 px of the list hides under the nav). Items are still reachable by scrolling.
7. **Desktop: empty Comments column has no hint**, while Attachments says "Click a task in To Do to view its attachments."
8. **Desktop navigation shows only "To Do | All Files"**; the **HR** entry that exists on mobile is not visible. Confirm this is intended.
9. **Images without alt text:** `account.png`, `todo.png`, `allfiles.png` (nav icons). Buttons and inputs are all labelled (0 unlabeled).
10. HR directory shows names, staff numbers and email addresses for ~850 people. Confirm the HR tab is limited to HR staff on the server side, not just hidden in the UI (could not test with a non-HR account).

### Security / hygiene (all low or informational)
11. Backend answers **401 without a token** on all cache routes (`/api/cache/entries`, `/todo`, `/entry`) and `/`. `/api/health` is public (200, 30 bytes), which is normal.
12. Backend sends **`X-Powered-By: Express`**; remove with `app.disable('x-powered-by')` or helmet.
13. **No tokens in `localStorage`** (MSAL keeps them elsewhere). `localStorage` holds large cached filter lists (up to ~9.5 KB each) and `sessionStorage` holds 13 keys (cached folder/list data); check whether any of it should be per-user and cleared on sign-out.

## 3. What worked well
- Clean cold start (no errors, cache session created).
- Cached To Do list expands with zero network calls.
- Every request in every flow returned 200 or 204; no throttling (429) in normal use.
- No horizontal scroll at 375 px or 1280 px.
- All buttons and inputs have accessible names.
- The account menu is minimal and safe (no destructive options besides Log Out).

## 4. Not covered
Creating, editing, approving or deleting tasks and forms; file upload/download; opening attachments in Office; push notifications / Capacitor and Electron (desktop) builds; the UAT configuration; a non-admin or non-HR account; real network outages and token expiry (the All Files error behaviour is in the other report).

## 5. Suggested next steps (not implemented)
1. Short-circuit Subordinate Tasks when there are no direct reports; cap or batch the 32 list queries when opening a task.
2. Lazy-load HR (first page only; fetch details when a person is opened).
3. Open Power Apps links in a new tab.
4. Fix the nav/dialog overlap, add alt text, add the empty-state text for Comments on desktop.
5. Hide `X-Powered-By`; confirm server-side authorisation of HR data with a non-HR test account.
