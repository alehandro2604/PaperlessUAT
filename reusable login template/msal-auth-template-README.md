# Reusable Azure AD login for Angular (standalone, Angular 17+)

A drop-in `auth/` folder for any Angular app that needs "Sign in with
Microsoft" — pulled out of the Paperless project's login flow and made
config-driven so nothing about a specific tenant, client ID, or API is
hardcoded. Copy the folder, write one config object, done.

Built for standalone Angular apps (signals, functional guards/interceptors,
`provideAppInitializer`). Tested against Angular 21 / `@azure/msal-browser` 5.

## What's in here

- `auth.config.ts` — the `MsalAuthConfig` interface + injection token. This is the only file you fill in per app.
- `auth.service.ts` — `MsalAuthService`: silent sign-in on load, redirect-based interactive sign-in, signals for template binding, `acquireToken()` for any scope.
- `auth.guard.ts` — `authGuard`, a route guard that requires sign-in.
- `auth.interceptor.ts` — `authInterceptor`, attaches bearer tokens to requests you configure.
- `provide-msal-auth.ts` — `provideMsalAuth(config)`, the one-liner that wires it all into `app.config.ts`.
- `login-button.component.ts` — a tiny example component showing the signals in a template. Replace or delete.

## 1. Install

```
npm install @azure/msal-browser
```

## 2. Register the app in Azure AD (one-time, per app)

Azure Portal → App registrations → New registration:
- Platform: **Single-page application**
- Redirect URI: your dev URL (`http://localhost:4200`) — add the prod URL too, or swap it per environment
- Note the **Application (client) ID** and **Directory (tenant) ID**
- API permissions → add whatever Graph/API scopes you need (e.g. `User.Read`), grant admin consent if required
- If calling your own backend API, expose a scope there (`Expose an API` → `access_as_user`) and add it as a permission here too

## 3. Copy the folder in

Copy `auth/` into `src/app/auth/` of the new project.

## 4. Write the config

Anywhere sensible (e.g. `src/app/auth/auth.config.instance.ts`, or your existing `environment.ts`):

```ts
import { MsalAuthConfig } from './auth/auth.config';

export const msalAuthConfig: MsalAuthConfig = {
  tenantId: 'YOUR-TENANT-ID',
  clientId: 'YOUR-CLIENT-ID',
  redirectUri: 'http://localhost:4200',
  loginScopes: ['User.Read'],

  // Optional — populate `auth.currentUser()` from Graph /me:
  loadProfile: async (accessToken) => {
    const res = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    return res.json();
  },

  // Optional — which outgoing requests get a bearer token, and which scope:
  apiScopeRules: [
    {
      matches: (url) => url.startsWith('http://localhost:3000/api'),
      scopes: ['api://YOUR-CLIENT-ID/access_as_user'],
    },
  ],

  // false = sign-in is optional/on-demand instead of forced on every visitor.
  autoRedirectToSignIn: true,
};
```

## 5. Wire it into `app.config.ts`

```ts
import { ApplicationConfig } from '@angular/core';
import { provideHttpClient, withInterceptors } from '@angular/common/http';
import { provideRouter } from '@angular/router';
import { routes } from './app.routes';
import { provideMsalAuth } from './auth/provide-msal-auth';
import { authInterceptor } from './auth/auth.interceptor';
import { msalAuthConfig } from './auth/auth.config.instance';

export const appConfig: ApplicationConfig = {
  providers: [
    provideMsalAuth(msalAuthConfig),
    provideHttpClient(withInterceptors([authInterceptor])),
    provideRouter(routes),
  ],
};
```

That's it — `provideMsalAuth` registers an app initializer that resolves
sign-in (redirect handling / silent sign-in) *before* your app's first
render, so `MsalAuthService`'s signals are already correct in `AppComponent`'s
first frame.

## 6. Use it

**In a template**, inject `MsalAuthService` and read its signals directly —
see `login-button.component.ts`:

```ts
protected readonly auth = inject(MsalAuthService);
// auth.isAuthenticated(), auth.currentUser(), auth.isBusy(), auth.errorMessage()
```

**Protect a route:**

```ts
import { authGuard } from './auth/auth.guard';

export const routes: Routes = [
  { path: 'reports', canActivate: [authGuard], component: ReportsComponent },
];
```

**Call an API needing a scope not covered by `apiScopeRules`:**

```ts
const token = await this.auth.acquireToken(['Sites.Read.All']);
```

## Design notes (so future-you knows why it's shaped this way)

- **Config-driven, not copy-edited.** The old Paperless `AuthService` had
  `sharePointConfig` imports and hardcoded scopes baked in. This version takes
  all of that through `MsalAuthConfig`, so the files themselves never change
  between projects — only the config object you write does.
- **`autoRedirectToSignIn`** exists because it genuinely differs by app type:
  an internal line-of-business tool (like Paperless) wants everyone
  auto-redirected into Microsoft sign-in on load; a public site with an
  optional "client portal" section wants visitors left alone until they hit
  a guarded route or click "Sign in". Same module, one flag.
- **`apiScopeRules` generalizes the SharePoint/Graph/backend multi-token
  pattern** from the original `AuthService` (it had three near-identical
  `acquireXToken()` methods). Now it's data: a list of `{matches, scopes}`,
  and the interceptor handles all of them the same way.
- **Signals, not `@Output`/manual change detection.** The original app used
  plain fields + `ChangeDetectorRef.detectChanges()`. Signals make
  `isAuthenticated`/`currentUser`/`isBusy` reactive for free.
- **`ensureReady()` instead of polling.** The guard and anything else that
  needs to wait for startup sign-in resolution awaits one promise rather than
  polling a status flag in a loop.

## What this does *not* include

Anything Paperless-specific stayed out on purpose: SharePoint site/list
config, the "not logged in" alert-once flag, HR folder/file logic, the
`FileCrawlCacheService` reset on user switch, etc. Those belong in each
app's own services, wired up in whatever runs after `afterSignIn()` resolves
(e.g. by reacting to `auth.currentUser()` with `effect()`), not in this module.
