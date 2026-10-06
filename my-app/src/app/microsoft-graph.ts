// ============================================================
// GRAPH API HELPERS
// Stateless utility functions used by multiple services.
// Pass HttpClient as a parameter so these remain tree-shakeable
// pure functions rather than services.
// Usage:  import { graphGet, graphPatch, toGraphPath, normalizeName, withTimeout } from './microsoft-graph';
// ============================================================

import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { timeout }    from 'rxjs/operators';
import { AppConstants } from './app.constants';

// ── HTTP wrapper ──────────────────────────────────────────────
/** Returns a typed Observable for a Graph v1.0 GET request. */
export function graphGet(
  http: HttpClient,
  path: string,
  token: string,
  timeoutMs: number = AppConstants.graphDefaultTimeoutMs,
  extraHeaders: Record<string, string> = {}
) {
  return http
    .get(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}`, ...extraHeaders },
    })
    .pipe(timeout(timeoutMs));
}

// ── Throttle control ──────────────────────────────────────────
// SharePoint throttles per user/app across the whole session, so limiting each feature
// on its own is not enough — a To Do load and a folder browse compete for the same
// budget. These module-level guards cap what the app has in flight overall, and make a
// 429 pause *every* caller instead of each one discovering it separately.

/** Hard ceiling on concurrent Graph requests, app-wide. */
const MAX_CONCURRENT_GRAPH_REQUESTS = 2;

/** While set, no request goes out until this time — shared cooldown after a 429. */
let graphThrottledUntilMs = 0;
/** Last time SharePoint answered 429 — background warm-up stays off for a while after. */
let lastGraphThrottleAtMs = 0;
/** How long background (non-click) Graph work stays paused after any 429. */
const BACKGROUND_BACKOFF_AFTER_429_MS = 2 * 60_000;

/**
 * True shortly after a 429. Prefetch / warm-up / polling should skip their run while
 * this holds: the retry cooldown only delays the failed request, and background work
 * resuming right after it is what kept re-tripping the throttle.
 */
export function isGraphBackgroundBackedOff(): boolean {
  return Date.now() - lastGraphThrottleAtMs < BACKGROUND_BACKOFF_AFTER_429_MS;
}
let activeGraphRequests = 0;
const graphSlotWaiters: Array<() => void> = [];

async function acquireGraphSlot(priority = false): Promise<void> {
  const wait = graphThrottledUntilMs - Date.now();
  if (wait > 0) await new Promise(resolve => setTimeout(resolve, wait));

  if (activeGraphRequests < MAX_CONCURRENT_GRAPH_REQUESTS) {
    activeGraphRequests += 1;
    return;
  }
  // Priority callers go to the front so they are not stuck behind background warm-up.
  await new Promise<void>(resolve =>
    priority ? graphSlotWaiters.unshift(resolve) : graphSlotWaiters.push(resolve)
  );
  activeGraphRequests += 1;
}

function releaseGraphSlot(): void {
  activeGraphRequests = Math.max(0, activeGraphRequests - 1);
  graphSlotWaiters.shift()?.();
}

/**
 * Drop a leftover 429 cooldown so an interactive click (HR Files / folder open)
 * is not stuck waiting out a pause caused by earlier background To Do traffic.
 */
export function clearGraphThrottleCooldown(): void {
  graphThrottledUntilMs = 0;
}

/** Honour Retry-After (seconds) when present, else exponential backoff. */
function resolveThrottleDelayMs(err: any, attempt: number): number {
  const retryAfterRaw =
    err?.headers?.get?.('Retry-After') ?? err?.headers?.get?.('retry-after');
  const retryAfterSec = Number(retryAfterRaw);

  if (Number.isFinite(retryAfterSec) && retryAfterSec > 0) {
    // Seconds → milliseconds, capped so a large Retry-After cannot stall the UI.
    return Math.min(retryAfterSec * 1000, 30_000);
  }
  return Math.min(16_000, 1_000 * Math.pow(2, attempt));
}

/**
 * Graph GET that respects the shared concurrency cap and backs off on HTTP 429.
 *
 * Prefer this over bare `graphGet` for anything in a loop or fan-out: an un-retried
 * page failure silently drops rows, and unlimited fan-out is what triggers throttling
 * in the first place.
 */
export async function graphGetWithRetry(
  http: HttpClient,
  path: string,
  token: string,
  timeoutMs: number = AppConstants.graphDefaultTimeoutMs,
  extraHeaders: Record<string, string> = {},
  maxRetries: number = 4,
  priority: boolean = false
): Promise<any> {
  let attempt = 0;

  while (true) {
    await acquireGraphSlot(priority);

    let failure: any = null;
    try {
      return await firstValueFrom(graphGet(http, path, token, timeoutMs, extraHeaders));
    } catch (err: any) {
      failure = err;
    } finally {
      // Released before sleeping so a backing-off request does not hold a slot.
      releaseGraphSlot();
    }

    const status = failure?.status ?? failure?.error?.status;
    if (status === 429) lastGraphThrottleAtMs = Date.now();
    if (status !== 429 || attempt >= maxRetries) throw failure;

    const delayMs = resolveThrottleDelayMs(failure, attempt);
    // Push out the shared cooldown so parallel callers stop hammering too.
    graphThrottledUntilMs = Math.max(graphThrottledUntilMs, Date.now() + delayMs);
    await new Promise(resolve => setTimeout(resolve, delayMs));
    attempt += 1;
  }
}

/** Returns a typed Observable for a Graph v1.0 POST request. */
export function graphPost(
  http: HttpClient,
  path: string,
  body: unknown,
  token: string,
  timeoutMs: number = AppConstants.graphDefaultTimeoutMs
) {
  return http
    .post(`https://graph.microsoft.com/v1.0${path}`, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    })
    .pipe(timeout(timeoutMs));
}

/** Returns a typed Observable for a Graph v1.0 PUT request (e.g. file upload). */
export function graphPut(
  http: HttpClient,
  path: string,
  body: BodyInit | null,
  token: string,
  contentType: string = 'application/octet-stream',
  timeoutMs: number = AppConstants.graphFileListingTimeoutMs
) {
  return http
    .put(`https://graph.microsoft.com/v1.0${path}`, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType,
      },
    })
    .pipe(timeout(timeoutMs));
}

/** Returns a typed Observable for a Graph v1.0 PATCH request. */
export function graphPatch(
  http: HttpClient,
  path: string,
  body: unknown,
  token: string,
  timeoutMs: number = AppConstants.graphDefaultTimeoutMs,
  useBeta: boolean = false
) {
  const baseUrl = useBeta ? 'https://graph.microsoft.com/beta' : 'https://graph.microsoft.com/v1.0';
  return http
    .patch(`${baseUrl}${path}`, body, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    })
    .pipe(timeout(timeoutMs));
}

// ── Pagination ────────────────────────────────────────────────
/**
 * Converts an @odata.nextLink absolute URL to the path portion
 * expected by graphGet(), or null when there are no more pages.
 * Accepts v1.0 / beta hosts and already-relative paths so a mismatched
 * prefix does not silently drop the rest of a paged result set.
 */
export function toGraphPath(nextLink?: string | null): string | null {
  if (!nextLink) return null;
  const link = String(nextLink).trim();
  if (!link) return null;
  if (link.startsWith('/')) return link;

  const prefixMatch = link.match(/^https?:\/\/graph\.microsoft\.com\/[^/]+(\/.*)$/i);
  if (prefixMatch?.[1]) return prefixMatch[1];

  const legacyPrefix = 'https://graph.microsoft.com/v1.0';
  if (link.startsWith(legacyPrefix)) return link.slice(legacyPrefix.length);

  return null;
}

// ── String helpers ────────────────────────────────────────────
/** Strips spaces, underscores and hyphens then lowercases — used for fuzzy name matching. */
export function normalizeName(value: string): string {
  return value.replace(/[\s_-]+/g, '').toLowerCase();//this replaces spaces, underscores and hyphens with empty strings and lowercases the string
}
// ── Promise helpers ───────────────────────────────────────────
/** Rejects with a readable error message if the inner promise takes longer than `ms`. */
export function withTimeout<T>(
  promise: Promise<T>,
  ms: number,
  message: string
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const id = setTimeout(() => reject(new Error(message)), ms);
    promise
      .then(v  => { clearTimeout(id); resolve(v); })
      .catch(e => { clearTimeout(id); reject(e);  });
  });
}
