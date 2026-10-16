// Persistent cache backed by Redis via the Node backend (paperless-backend),
// replacing the previous IndexedDB implementation. The public API is
// unchanged, so callers (FileCrawlCacheService) work as before. Most keys are
// namespaced per signed-in user (e.g. My Tasks). Shared across every
// account/browser on the backend:
//   fc:files:… (Attachments, HR people list, All Files listings)
//   fc:tasks:lib:v7:… / fc:tasks:hr-folder:… / fc:tasks:all-files-comments:…
// so one user's crawl warms Redis for everyone.
import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { AuthService } from './auth.service';
import { CacheSessionService } from './cache-session.service';
import { sharePointConfig } from '../sharepoint.config';

interface CacheEntry<T> {
  data: T;
  timestamp: number;
}

/** Bodies above this are gzipped before upload (task snapshots shrink ~90%+). */
const GZIP_MIN_BYTES = 8 * 1024;

/** FNV-1a over the serialized body — cheap enough for multi-MB snapshots. */
function fingerprint(text: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${text.length}:${(hash >>> 0).toString(16)}`;
}

/** Gzip via the browser's native CompressionStream; null where unsupported. */
async function gzip(text: string): Promise<ArrayBuffer | null> {
  if (typeof CompressionStream === 'undefined') return null;
  const stream = new Blob([text]).stream().pipeThrough(new CompressionStream('gzip'));
  return new Response(stream).arrayBuffer();
}

@Injectable({ providedIn: 'root' })
export class CacheService {
  private readonly base = `${sharePointConfig.backendUrl}/api/cache`;
  /**
   * Fingerprint of the last body stored per key. Soft refreshes often re-save an
   * identical snapshot; skipping those avoids re-uploading multi-MB payloads.
   */
  private readonly lastStored = new Map<string, string>();
  /**
   * Keys the backend answered 403 for. The backend remembers a denial for 5 minutes,
   * so asking again sooner only repeats the refusal (one request per page of tasks).
   */
  private readonly deniedUntil = new Map<string, number>();
  private static readonly DENIED_RETRY_MS = 5 * 60_000;

  private isDenied(key: string): boolean {
    const until = this.deniedUntil.get(key);
    if (until === undefined) return false;
    if (Date.now() < until) return true;
    this.deniedUntil.delete(key);
    return false;
  }

  private noteDenied(err: unknown, key: string): boolean {
    if ((err as { status?: number })?.status !== 403) return false;
    this.deniedUntil.set(key, Date.now() + CacheService.DENIED_RETRY_MS);
    return true;
  }

  constructor(
    private http: HttpClient,
    private auth: AuthService,
    private session: CacheSessionService,
  ) {}

  /** Token for the backend API (not Graph) — required by the OBO/cache endpoints. */
  private async authHeaders(): Promise<{ Authorization: string }> {
    const token = await this.auth.acquireTokenSilent([sharePointConfig.backendApiScope]);
    this.session.noteToken(token);
    void this.session.ensure();
    return { Authorization: `Bearer ${token}` };
  }

  async get<T>(key: string): Promise<CacheEntry<T> | undefined> {
    if (this.isDenied(key)) return undefined;
    const headers = await this.authHeaders();
    try {
      const entry = await firstValueFrom(
        this.http.get<CacheEntry<T> | null>(`${this.base}/entry`, { headers, params: { key } })
      );
      return entry ?? undefined;
    } catch (err) {
      // Treated as a cache miss by callers; just don't ask again for a while.
      if (this.noteDenied(err, key)) return undefined;
      throw err;
    }
  }

  async set<T>(key: string, data: T): Promise<void> {
    const json = JSON.stringify({ key, data });
    const print = fingerprint(json);
    if (this.lastStored.get(key) === print) return;
    if (this.isDenied(key)) return;
    try {
      const headers: Record<string, string> = {
        ...(await this.authHeaders()),
        'Content-Type': 'application/json',
      };
      // Express inflates gzip request bodies itself (body-parser `inflate`).
      const zipped = json.length >= GZIP_MIN_BYTES ? await gzip(json) : null;
      if (zipped) headers['Content-Encoding'] = 'gzip';
      await firstValueFrom(this.http.put(`${this.base}/entry`, zipped ?? json, { headers }));
      this.lastStored.set(key, print);
    } catch (err: unknown) {
      const status = (err as { status?: number })?.status;
      if (this.noteDenied(err, key)) {
        console.warn(`[cache] "${key}" refused by the backend (403); memory-only for 5 min`);
        return;
      }
      // Oversized snapshot — keep the memory cache, but drop the Redis copy: it is now
      // older than what we failed to save (e.g. still lists a task just completed),
      // and a fresh load beats painting stale rows on the next visit.
      if (status === 413) {
        console.warn(`[cache] skipped persist for "${key}" (payload too large); clearing stale copy`);
        await this.clear(key).catch(() => undefined);
        return;
      }
      throw err;
    }
  }

  async clear(key: string): Promise<void> {
    this.lastStored.delete(key);
    const headers = await this.authHeaders();
    await firstValueFrom(this.http.delete(`${this.base}/entry`, { headers, params: { key } }));
  }

  /** All stored entries whose key starts with `prefix`, with their timestamps. */
  async entriesByPrefix<T>(prefix: string): Promise<Array<{ key: string; entry: CacheEntry<T> }>> {
    const headers = await this.authHeaders();
    return firstValueFrom(
      this.http.get<Array<{ key: string; entry: CacheEntry<T> }>>(`${this.base}/entries`, {
        headers,
        params: { prefix },
      })
    );
  }

  /** Deletes every stored entry whose key starts with `prefix`. */
  async deleteByPrefix(prefix: string): Promise<void> {
    for (const key of [...this.lastStored.keys()]) {
      if (key.startsWith(prefix)) this.lastStored.delete(key);
    }
    const headers = await this.authHeaders();
    await firstValueFrom(this.http.delete(`${this.base}/entries`, { headers, params: { prefix } }));
  }

  async clearAll(): Promise<void> {
    return this.deleteByPrefix('');
  }
}
