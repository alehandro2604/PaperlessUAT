# All Files – HTTP error behaviour test report

**Date:** 2026-10-14  **Branch:** `cache-per-task-todo`  **App:** `ng serve --configuration live` on http://localhost:4200 (real SharePoint tenant, read-only actions)
**Scope:** test only. No source files were changed during testing. (The earlier first-search change in `all-files.ts` – search runs before the crawl, no gap between search pages, parallel root-folder lookups – was already in place and is part of what was tested.)

## 1. Method

- Signed in manually in the in-app browser; viewport was the mobile layout.
- Injected a passive logger into the page (not into the source). It records every `fetch`/XHR with time, method, URL (trimmed), status and duration, plus clicks, input and `console.error`, and tags each request with the **last user action**.
- **Natural run:** open All Files, first (cold) search, repeat search. No faults.
- **Induced faults:** the same logger returns a fake response (400, 401, 403, 404, 429, 502, 503, 500 or a network failure) for chosen URL patterns. Real servers were never sent faults; the fault exists only in the browser.
- "Cold" = page reloaded, so in-memory caches were empty (the shared Redis cache on `localhost:3000` still existed unless it was faulted too).
- Search term for all runs: `duncan` (returns ~100 folders, mostly via the **Assigned To** field).

Limits: real token expiry (true 401) and real outages were not reproduced; faults are simulated. One early cache-fault rule did not match the calls and was redone with a simpler rule (result below is from the redo).

## 2. Natural (no fault) results

| Action before | Requests | Statuses | Time | Result |
|---|---|---|---|---|
| Open All Files tab (cold) | none for the listing | – | instant | "Type a keyword to search folders in all libraries." |
| First search `report` | 38 | all 200 | ~11 s of Graph search, then cache hydrate + crawl | "No matching folders found." (valid, no such folder) |
| First search `duncan` (warm repeat) | 32 | all 200 | ~10 s | 100 cards |
| Console errors / page errors | 0 | – | – | – |

No 4xx/5xx happened on their own during the session.

## 3. Induced error results – Graph folder search (`/drives/{id}/root/search`)

Triggered by: typing `duncan` in the All Files search box (state: warm – in-memory list already built by an earlier search).

| Status | Graph search calls | Retries | User saw | Notes |
|---|---|---|---|---|
| 400 | 15 (one per library) | none | **100 cards, no message** | masked by in-memory list |
| 401 | 15 | none, **no token refresh** | 100 cards, no message | token not re-acquired |
| 403 | 15 | none | 100 cards, no message | |
| 404 | 15 | none | 100 cards, no message | |
| 429 | 16 | yes, backs off, shared cooldown | 100 cards, no message | calls spread over ~7 s |
| 502 | 19 | none by the app (4 extra were 429 retries left over from the previous test) | 100 cards, no message | |
| 503 | 15 | none | 100 cards, no message | |
| Network failure (status 0) | 15 | none | 100 cards, no message | |

Every failure is only logged with `console.error("Graph folder search failed for <library>")` – 15 lines per search. Because one failure stops that library's loop, a failed run makes 15 calls instead of the normal 32.

Retry rule confirmed in `microsoft-graph.ts` (`graphGetWithRetry`): **only 429** is retried (max 4, honours `Retry-After`, else 1/2/4/8 s). Every other status or network error is thrown on the first attempt.

## 4. Cold-start scenarios

| # | State / action before | Faults | Requests | User saw |
|---|---|---|---|---|
| C1 | Reload, open All Files, search `duncan` | Graph search 403 (15 calls) | 32 (15×403, rest 200) | **100 cards, no message.** Search failed, but the crawl then filled from the Redis cache (`localhost:3000`) so results appeared anyway. |
| C2 | Reload, open All Files, search `duncan` | Graph search + `root/children` + (intended) cache = 502 | 62 (46×502, 16×200) | **Two messages at once:** "Could not load folders from SharePoint. Please refresh to try again." **and** "No matching folders found." 0 cards. |
| C2b | Same, Graph healthy, next search `duncan` | none | 49, all 200 | Recovered automatically: 100 cards, error gone. |
| C3 | Reload, open All Files, search `duncan` | Redis cache `/api/cache/*` = 500 (Graph healthy) | 54 (8×500) | "No matching folders found." at 20 s, **cards then trickled in: 43 at ~25 s** as the crawl paged each library (350 ms gaps, one library at a time). Cache GET and PUT both failed silently. |

C3 is the important performance finding: matches on **Assigned To** (and library) exist only after the root crawl, so when the Redis cache is cold or down, the first assignee search shows nothing for tens of seconds. Graph search only matches folder/file names.

## 5. Single-library "Files" view (`/drives/{id}/root/children`)

Triggered by: reload → All Files → filter "Files" (default library), then switching library.

| Action before | Fault | Retries | User saw |
|---|---|---|---|
| Select "Files" filter (Call for Applications) | children 403 (+cache 403) | none | **Raw error text:** `Http failure response for https://graph.microsoft.com/v1.0/drives/<driveId>/root/children?$select=id,name,webUrl,… 403 simulated` |
| Switch to Procurement | children 404 | none | same raw text, "404 simulated" |
| Switch to Illicit Connections | children 429 | 4 retries over ~15 s (1, 2, 4, 8 s) | "Loading folders…" the whole time, then the raw text, "429 simulated" |
| Switch back to Procurement, Graph healthy | none | – | Loaded normally (recovers) |

No `console.error` is written for these; the message only goes to the screen.

## 6. Findings (ranked)

1. **Raw Angular HTTP error shown to users in the Files view** (all statuses). It exposes the Graph URL and drive id, is unreadable on mobile (text overflows) and has no Retry button.
2. **Graph search failures are invisible.** If every library search fails (400/401/403/404/5xx/network) but local data exists, the user gets stale/partial results with no indication. Only the console shows it.
3. **Contradictory messages** when everything fails: the "Could not load folders…" banner and "No matching folders found." appear together.
4. **Only 429 is retried.** 502/503/network blips, which are usually transient, fail on the first attempt, and one blip drops a whole library from search results.
5. **401 is not handled** in search/listing: no token re-acquire and retry, no sign-in prompt.
6. **Slow/blank first assignee search when the Redis cache is cold or down** (C3). Search finds nothing until the crawl reaches each library. There is no "still loading more libraries" hint, only "No matching folders found."
7. **Redis cache errors are silent** (GET and PUT 500). That is good for resilience, but it hides that every user falls back to the slow crawl path.
8. **Housekeeping observation:** `DRIVE_CHILDREN_QUERY` requests `$top=10000` while comments in the code describe `$top=100` paging; worth confirming which is intended.

## 7. What worked well

- 429 handling: shared cooldown, `Retry-After` honoured, spinner kept while retrying, loader stays on.
- One failing library does not abort the others in search.
- Recovery is automatic on the next search or library switch.
- The Redis cache fallback (C1) hides Graph search problems from the user.
- No uncaught JS errors or unhandled promise rejections in any scenario.

## 8. Suggested follow-ups (not implemented)

- Replace the raw `err.message` with a friendly message plus a Retry button; log details to the console only.
- Show a small "search partly failed (n of 15 libraries)" notice instead of staying silent.
- Retry 502/503/network once or twice with a short back-off; on 401 re-acquire the token once and retry.
- Show "Still loading n more libraries…" while the crawl runs, so an empty result is not read as "no match".
- Make the banner and the "No matching folders found." message mutually exclusive.
