# Paperless Backend

Express server that sits between the Angular app and Microsoft Graph, caching
per-person task lookups in Redis. Uses the on-behalf-of (OBO) flow so Graph
enforces each signed-in user's own SharePoint permissions.

## Local Redis (no Docker)

Docker Desktop cannot run on this machine (WSL install is blocked by group
policy), so local dev uses a portable native Windows Redis build in
`../redis-windows`. Start it with:

```powershell
..\redis-windows\redis-server.exe --port 6379
```

Check it: `..\redis-windows\redis-cli.exe ping` should reply `PONG`.
In production you would use Azure Cache for Redis instead — only `REDIS_URL`
in `.env` changes, no code changes.

## One-time Azure Portal setup (required before /api/tasks works)

The health endpoint works without this, but the OBO token exchange needs two
things on the app registration (`45a7a07b-7d2b-4240-98cf-2847ae6245f9`):

1. **Client secret** — App registrations → your app → Certificates & secrets →
   New client secret. Paste the *value* into `CLIENT_SECRET` in `.env`.
2. **Expose an API** — App registrations → your app → Expose an API →
   Set the Application ID URI (default `api://45a7a07b-7d2b-4240-98cf-2847ae6245f9`)
   → Add a scope named `access_as_user`, enabled, admins+users can consent.
   The Angular app requests this scope (see `backendApiScope` in
   `my-app/src/app/sharepoint.config.ts`) and sends the resulting token here.

Without step 2, MSAL in Angular cannot acquire a token whose audience is this
backend, and the OBO exchange will fail with `invalid_grant`.

## Run

```powershell
npm install
npm run dev     # nodemon, auto-restart on changes (or: npm start)
```

- `GET /api/health` — no auth; confirms server + Redis are up.
- `GET /api/metrics/ram` — requires auth; Redis RAM by shared vs per-user cache (for sizing the VM / Azure Cache). Or on the server: `npm run ram-report`.
- `GET /api/tasks/:personEmail?list=<listNameOrId>` — requires
  `Authorization: Bearer <token for this API>`; checks Redis first, falls back
  to Graph, caches for `CACHE_TTL_SECONDS` (default 300).
- `DELETE /api/tasks/:personEmail?list=<listNameOrId>` — drops that Redis
  read-through entry so the next GET refetches from Graph.

All authenticated routes validate the JWT signature against the tenant's
signing keys and check audience + tenant (`authMiddleware.js`).

## Per-user cache API (replaces IndexedDB in the Angular app)

The Angular `CacheService` now persists through these endpoints instead of
IndexedDB. Keys are namespaced server-side by the caller's Entra object id,
so users cannot read or wipe each other's entries. Entries auto-expire after
`CACHE_ENTRY_TTL_SECONDS` (default 7 days, matching the old client-side max age).

- `GET /api/cache/entry?key=<k>` — `{ data, timestamp }` or `null`
- `PUT /api/cache/entry` — body `{ key, data }`
- `DELETE /api/cache/entry?key=<k>`
- `GET /api/cache/entries?prefix=<p>` — `[{ key, entry }]`
- `DELETE /api/cache/entries?prefix=<p>` — empty prefix deletes all of the user's entries

## Config

Copy `.env.example` to `.env`. `TENANT_ID`/`CLIENT_ID` are prefilled from the
Angular app's `sharepoint.config.ts`; set `CLIENT_SECRET` and optionally
`TASKS_LIST` (default SharePoint list for /api/tasks).

Optional: `CORS_ORIGINS=https://swiftpaperlessuat.enemalta.lan` (comma-separated).
Not needed if IIS/nginx reverse-proxies `/api` onto the same host as the SPA.

## Deploying to UAT (`https://swiftpaperlessuat.enemalta.lan`)

Local works because the SPA calls `http://localhost:3000`. On UAT that breaks:

1. **Mixed content** — HTTPS page cannot call `http://…`
2. **Wrong host** — `localhost` is the user's PC, not the app server
3. **Backend/Redis must run on the UAT server** (or a reachable host)

### Checklist

1. **Azure AD app registration** (`45a7a07b-…`)
   - SPA redirect URI: `https://swiftpaperlessuat.enemalta.lan` (and with trailing `/` if you use it)
   - Expose API scope `access_as_user` (same as local)
   - Client secret present in the **server** `.env` (`CLIENT_SECRET`)

2. **Run Node + Redis on the UAT host**
   - `REDIS_URL` pointing at Redis on that machine (or Azure Cache)
   - `npm start` (or a Windows service / PM2) listening on `127.0.0.1:3000`

3. **Reverse-proxy `/api` to Node** (recommended — keeps `backendUrl: ''`)
   - IIS: install URL Rewrite + ARR, reverse proxy
     `https://swiftpaperlessuat.enemalta.lan/api/*` → `http://127.0.0.1:3000/api/*`
   - Or nginx: `location /api/ { proxy_pass http://127.0.0.1:3000; … }`

4. **Angular config** (`sharepoint.config.ts`) for UAT:
   - `redirectUri: 'https://swiftpaperlessuat.enemalta.lan'`
   - `backendUrl: ''` (same origin via the proxy above)
   - Rebuild and redeploy the SPA (`ng build` → IIS site root)

5. **Smoke test** from a browser on the corporate network:
   - `https://swiftpaperlessuat.enemalta.lan/api/health` → `{ "server":"ok","redis":"PONG" }`
   - Sign in; DevTools Network should show `/api/cache/…` as 200, not blocked/failed

If you cannot reverse-proxy, set `backendUrl` to a separate **HTTPS** API host
(e.g. `https://swiftpaperlessuat-api.enemalta.lan`) and add that origin to
`CORS_ORIGINS`. Never use `http://localhost:3000` in a UAT build.
