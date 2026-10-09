// To Do cache, one Redis entry per task (backend: /api/cache/todo and /api/cache/tasks
// in cacheRoutes.js). Replaces the single `tasks:hr-user:v5` snapshot, which outgrew the
// upload limit and was re-sent whole after every change. Same read/write shape as the
// FileCrawlCacheService calls it replaces, so AppComponent only swaps call sites:
//   getStale(HR_USER_TASKS_CACHE_KEY)   → getAll()
//   get(HR_USER_TASKS_CACHE_KEY)        → getFresh()
//   set(HR_USER_TASKS_CACHE_KEY, x)     → replaceAll(x)
//   invalidate(HR_USER_TASKS_CACHE_KEY) → invalidate()
import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { AuthService } from './auth.service';
import { sharePointConfig } from '../sharepoint.config';

interface TodoResponse {
  updatedAt: number;
  tasks: { key: string; data: any }[];
}

/** Same freshness window as FileCrawlCacheService.CACHE_TTL. */
const FRESH_FOR_MS = 60 * 60 * 1000;
const SAVE_DEBOUNCE_MS = 1500;
const UPSERTS_PER_REQUEST = 200;

/** FNV-1a, as in cache.ts — detects which tasks changed since they were last saved. */
function fingerprint(text: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${text.length}:${(hash >>> 0).toString(16)}`;
}

@Injectable({ providedIn: 'root' })
export class TodoTaskStoreService {
  private readonly base = `${sharePointConfig.backendUrl}/api/cache`;
  private items: any[] = [];
  private updatedAt = 0;
  private loaded = false;
  private loading: Promise<void> | null = null;

  /** Fingerprint of each task as Redis has it, so only changed tasks are uploaded. */
  private readonly saved = new Map<string, string>();
  private readonly pendingUpserts = new Map<string, { data: any; print: string }>();
  private readonly pendingRemoves = new Set<string>();
  private saveTimer: ReturnType<typeof setTimeout> | null = null;
  /** Saves run one after another so a delete can never overtake an earlier upload. */
  private saving: Promise<void> = Promise.resolve();

  constructor(private readonly http: HttpClient, private readonly auth: AuthService) {}

  static keyOf(item: any): string {
    return `${item?.listName ?? 'unknown'}:${item?.id}`;
  }

  /** Pulls this user's To Do from Redis once per session. Safe to call repeatedly. */
  load(): Promise<void> {
    if (this.loaded) return Promise.resolve();
    this.loading ??= this.fetchAll().finally(() => { this.loading = null; });
    return this.loading;
  }

  /** Every cached task regardless of age (was getStale). */
  getAll(): any[] | null {
    return this.items.length ? this.items : null;
  }

  /** Cached tasks only if saved within the last hour (was get). */
  getFresh(): any[] | null {
    return Date.now() - this.updatedAt <= FRESH_FOR_MS ? this.getAll() : null;
  }

  /** Keep this list in memory and upload only the tasks that changed or left (was set). */
  replaceAll(items: any[]): void {
    this.items = items;
    this.updatedAt = Date.now();
    this.loaded = true;

    const present = new Set<string>();
    for (const item of items) {
      const key = TodoTaskStoreService.keyOf(item);
      present.add(key);
      const print = fingerprint(JSON.stringify(item));
      if (this.saved.get(key) !== print) {
        this.pendingUpserts.set(key, { data: item, print });
        this.pendingRemoves.delete(key);
      } else {
        this.pendingUpserts.delete(key);
      }
    }
    for (const key of this.saved.keys()) {
      if (!present.has(key)) {
        this.pendingRemoves.add(key);
        this.pendingUpserts.delete(key);
      }
    }
    this.scheduleSave();
  }

  /** Forget the whole To Do cache, in memory and in Redis (was invalidate). */
  invalidate(): void {
    this.items = [];
    this.updatedAt = 0;
    this.saved.clear();
    this.pendingUpserts.clear();
    this.pendingRemoves.clear();
    if (this.saveTimer) {
      clearTimeout(this.saveTimer);
      this.saveTimer = null;
    }
    this.saving = this.saving.then(async () => {
      try {
        await firstValueFrom(this.http.delete(`${this.base}/todo`, { headers: await this.headers() }));
      } catch (err) {
        console.warn('[todo-store] clear failed', err);
      }
    });
  }

  private async fetchAll(): Promise<void> {
    try {
      const res = await firstValueFrom(
        this.http.get<TodoResponse>(`${this.base}/todo`, { headers: await this.headers() }),
      );
      for (const { key, data } of res.tasks) {
        this.saved.set(key, fingerprint(JSON.stringify(data)));
      }
      if (this.loaded) {
        // SharePoint answered first: keep its list, just work out what Redis needs.
        this.replaceAll(this.items);
      } else {
        this.items = res.tasks.map((t) => t.data);
        this.updatedAt = res.updatedAt;
      }
    } catch (err) {
      console.warn('[todo-store] load failed; To Do will load from SharePoint', err);
    }
    this.loaded = true;
  }

  private scheduleSave(): void {
    if (this.saveTimer) clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => {
      this.saveTimer = null;
      this.saving = this.saving.then(() => this.flush());
    }, SAVE_DEBOUNCE_MS);
  }

  private async flush(): Promise<void> {
    const upserts = [...this.pendingUpserts.entries()];
    const removes = [...this.pendingRemoves];
    this.pendingUpserts.clear();
    this.pendingRemoves.clear();
    if (upserts.length === 0 && removes.length === 0) return;

    try {
      const headers = await this.headers();
      // One request when there are only removals; otherwise chunks of 200 tasks.
      for (let i = 0; i < Math.max(upserts.length, 1); i += UPSERTS_PER_REQUEST) {
        const chunk = upserts.slice(i, i + UPSERTS_PER_REQUEST);
        await firstValueFrom(this.http.put(`${this.base}/tasks`, {
          upsert: chunk.map(([key, { data }]) => ({ key, data })),
          remove: i === 0 ? removes : [],
        }, { headers }));
        for (const [key, { print }] of chunk) this.saved.set(key, print);
      }
      for (const key of removes) this.saved.delete(key);
    } catch (err) {
      // Put them back so the next change retries them; newer pending edits win.
      for (const [key, value] of upserts) {
        if (!this.pendingUpserts.has(key)) this.pendingUpserts.set(key, value);
      }
      for (const key of removes) this.pendingRemoves.add(key);
      console.warn('[todo-store] save failed; will retry with the next change', err);
    }
  }

  /** Token for the backend API (not Graph), as CacheService uses. */
  private async headers(): Promise<{ Authorization: string }> {
    const token = await this.auth.acquireTokenSilent([sharePointConfig.backendApiScope]);
    return { Authorization: `Bearer ${token}` };
  }
}
