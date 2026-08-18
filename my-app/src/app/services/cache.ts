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
import { sharePointConfig } from '../sharepoint.config';

interface CacheEntry<T> {
  data: T;
  timestamp: number;
}

@Injectable({ providedIn: 'root' })
export class CacheService {
  private readonly base = `${sharePointConfig.backendUrl}/api/cache`;

  constructor(private http: HttpClient, private auth: AuthService) {}

  /** Token for the backend API (not Graph) — required by the OBO/cache endpoints. */
  private async authHeaders(): Promise<{ Authorization: string }> {
    const token = await this.auth.acquireTokenSilent([sharePointConfig.backendApiScope]);
    return { Authorization: `Bearer ${token}` };
  }

  async get<T>(key: string): Promise<CacheEntry<T> | undefined> {
    const headers = await this.authHeaders();
    const entry = await firstValueFrom(
      this.http.get<CacheEntry<T> | null>(`${this.base}/entry`, { headers, params: { key } })
    );
    return entry ?? undefined;
  }

  async set<T>(key: string, data: T): Promise<void> {
    try {
      const headers = await this.authHeaders();
      await firstValueFrom(this.http.put(`${this.base}/entry`, { key, data }, { headers }));
    } catch (err: unknown) {
      const status = (err as { status?: number })?.status;
      // Oversized All Files Comments snapshots etc. — keep memory cache, skip Redis.
      if (status === 413) {
        console.warn(`[cache] skipped persist for "${key}" (payload too large)`);
        return;
      }
      throw err;
    }
  }

  async clear(key: string): Promise<void> {
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
    const headers = await this.authHeaders();
    await firstValueFrom(this.http.delete(`${this.base}/entries`, { headers, params: { prefix } }));
  }

  async clearAll(): Promise<void> {
    return this.deleteByPrefix('');
  }
}
