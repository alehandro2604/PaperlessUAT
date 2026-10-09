//this is the file crawl cache service
//it is used to cache the file crawl results
//it is used to cache the file crawl results for a given library

import { Injectable } from '@angular/core';
import { CacheService } from './cache';

interface CacheEntry {
  data: unknown;
  timestamp: number;
}

/** Redis / persist key prefix for crawl-cache entries. */
const PERSIST_PREFIX = 'fc:';
/** Persist key holding which user the personal entries belong to. */
const OWNER_KEY = 'fc-meta:owner';

/**
 * Shared Redis families (must match paperless-backend/cacheRoutes.js) pulled into
 * memory in the background after sign-in, for Comments search.
 * Loaded separately from personal `fc:` hydrate so logout never scans/wipes them.
 *
 * `fc:files:` is deliberately NOT listed: the personal `fc:` scan in bindToUser() already
 * returned those entries, so listing it again downloaded the same ~3 MB twice. Shared
 * file listings (`fc:files:folder:…`, `fc:files:lib:…`) load one key at a time via
 * hydrateFromPersistent() when a screen needs them.
 *
 * `fc:tasks:hr-folder:` is shared too but deliberately not bulk-loaded: it is ~200 MB
 * across every HR person, and blocking sign-in on it took 10+ s and ~500 MB of heap.
 * HR snapshots are pulled one person at a time via hydrateFromPersistent().
 * all-files-comments is never persisted — the full search index is too large (413).
 */

// Shared prefixes bulk-loaded in the background after sign-in.
// `fc:files:` is deliberately NOT listed: the personal `fc:` scan in bindToUser()
// already returns those entries, so listing it again downloaded the same ~3 MB twice.
// Shared folder listings (`fc:files:folder:…`) are loaded one key at a time
// by hydrateFromPersistent() instead.
const SHARED_HYDRATE_PREFIXES = [
  'fc:tasks:lib:v7:',
] as const;

@Injectable({ providedIn: 'root' })
export class FileCrawlCacheService {
  private cache = new Map<string, CacheEntry>();
  /** Keep All Files / folder listings warm across tab switches and library changes. */
  private readonly CACHE_TTL = 60 * 60 * 1000; // 60 minutes

  /** Persisted snapshots older than this are dropped at sign-in instead of hydrated. */
  private readonly PERSIST_MAX_AGE = 7 * 24 * 60 * 60 * 1000; // 7 days

  private boundUser: string | null = null;
  private readyResolve!: () => void;
  private readonly ready = new Promise<void>(resolve => (this.readyResolve = resolve));
  /**
   * Persist writes are chained behind `ready` so they always land after
   * sign-in hydration, keeping Redis and memory consistent.
   */
  private persistQueue: Promise<void>;

  constructor(private readonly persistent: CacheService) {
    this.persistQueue = this.ready;
  }

  /**
   * Hydrates personal Redis entries for the signed-in user, then — without
   * blocking `ready` — warms the shared All Files snapshots in the background.
   * Call once after sign-in, before the first cache read.
   */
  async bindToUser(userKey: string): Promise<void> {
    const normalized = (userKey || '').trim().toLowerCase();
    if (this.boundUser === normalized) return;
    if (this.boundUser !== null) {
      // Mid-session account swap — the old user's in-memory data must go too.
      this.cache.clear();
    }
    this.boundUser = normalized;

    try {
      const owner = await this.persistent.get<string>(OWNER_KEY);
      if (owner?.data !== normalized) {
        // Personal keys only — shared folder caches must survive account swap.
        await this.persistent.deleteByPrefix(PERSIST_PREFIX);
        await this.persistent.set(OWNER_KEY, normalized);
      } else {
        await this.ingestPrefixEntries(PERSIST_PREFIX);
      }
    } catch {
      // Redis / backend unavailable — memory-only mode.
    } finally {
      this.readyResolve();
    }
    // Shared keys are not included in the broad `fc:` scan (logout-safe). Loaded after
    // `ready` so the first paint never waits on them; entries fetched meanwhile win.
    void this.hydrateSharedPrefixes();
  }

  /**
   * Pull shared All Files + HR Files Redis snapshots into memory without
   * touching personal keys. Safe to call after bindToUser or on demand.
   */
  async hydrateSharedPrefixes(): Promise<void> {
    await this.ready;
    for (const prefix of SHARED_HYDRATE_PREFIXES) {
      try {
        await this.ingestPrefixEntries(prefix);
      } catch {
        // Best-effort — click path still hydrates one key at a time.
      }
    }
  }

  /** Load Redis entries for `prefix` into the in-memory map (skip expired). */
  private async ingestPrefixEntries(prefix: string): Promise<void> {
    const entries = await this.persistent.entriesByPrefix<unknown>(prefix);
    const now = Date.now();
    for (const { key, entry } of entries) {
      if (!entry || now - entry.timestamp > this.PERSIST_MAX_AGE) {
        void this.persistent.clear(key).catch(() => {});
        continue;
      }
      const memoryKey = key.startsWith(PERSIST_PREFIX)
        ? key.slice(PERSIST_PREFIX.length)
        : key;
      // Anything already fetched this session wins over the stored snapshot.
      if (!this.cache.has(memoryKey)) {
        this.cache.set(memoryKey, { data: entry.data, timestamp: entry.timestamp });
      }
    }
  }

  /**
   * Pulls one key from Redis into memory when the in-memory map misses.
   * Used for shared folder tasks + file listings (All Files + HR Files) so
   * another user's Redis write can paint without waiting on SharePoint.
   * Returns { data, fresh } on hit, or null on miss / error / expired persist.
   */
  async hydrateFromPersistent<T = any[]>(libraryKey: string): Promise<{ data: T; fresh: boolean } | null> {
    try {
      await this.ready;
      const entry = await this.persistent.get<T>(PERSIST_PREFIX + libraryKey);
      if (!entry) return null;
      if (Date.now() - entry.timestamp > this.PERSIST_MAX_AGE) {
        void this.persistent.clear(PERSIST_PREFIX + libraryKey).catch(() => {});
        return null;
      }
      if (!this.cache.has(libraryKey)) {
        this.cache.set(libraryKey, { data: entry.data, timestamp: entry.timestamp });
      }
      const fresh = Date.now() - entry.timestamp <= this.CACHE_TTL;
      return { data: entry.data as T, fresh };
    } catch {
      return null;
    }
  }

  /**
   * Returns cached data for a key if it exists and hasn't expired.
   * Returns null if there's no cache or it's stale, so the caller knows to re-crawl.
   */
  get<T = any[]>(libraryKey: string): T | null {
    const entry = this.cache.get(libraryKey);
    if (!entry) {
      return null;
    }
    // Expired entries are kept (not deleted) so getStale() can still paint them.
    if (Date.now() - entry.timestamp > this.CACHE_TTL) {
      return null;
    }
    return entry.data as T;
  }

  /**
   * Cached data regardless of age — including snapshots persisted from a previous
   * session. Use for an instant first paint, then refresh from SharePoint.
   */
  getStale<T = any[]>(libraryKey: string): T | null {
    return (this.cache.get(libraryKey)?.data as T) ?? null;
  }

  /**
   * Stores crawl / listing results for a given key (memory + Redis).
   */
  set(libraryKey: string, data: unknown): void {
    this.cache.set(libraryKey, {
      data,
      timestamp: Date.now(),
    });
    this.enqueuePersist(() => this.persistent.set(PERSIST_PREFIX + libraryKey, data));
  }

  /**
   * Memory-only write — use for oversized payloads (e.g. All Files Comments index)
   * that exceed the Redis/API body limit (HTTP 413).
   */
  setMemoryOnly(libraryKey: string, data: unknown): void {
    this.cache.set(libraryKey, {
      data,
      timestamp: Date.now(),
    });
  }

  /**
   * Clears the cache for one key — call after upload/delete/rename
   * so the next load re-crawls fresh instead of showing stale data.
   */
  invalidate(libraryKey: string): void {
    this.cache.delete(libraryKey);
    this.enqueuePersist(() => this.persistent.clear(PERSIST_PREFIX + libraryKey));
  }

  /**
   * Clears every key that starts with the given prefix
   * (e.g. `files:folder:` after a manual All Files refresh).
   */
  invalidatePrefix(prefix: string): void {
    for (const key of [...this.cache.keys()]) {
      if (key.startsWith(prefix)) {
        this.cache.delete(key);
      }
    }
    this.enqueuePersist(() => this.persistent.deleteByPrefix(PERSIST_PREFIX + prefix));
  }

  /**
   * Clears everything, memory and disk — used on logout so a shared
   * machine does not keep the previous user's data around.
   */
  clearAll(): void {
    this.cache.clear();
    this.enqueuePersist(() => this.persistent.deleteByPrefix(PERSIST_PREFIX));
  }

  /** Non-expired cache entries whose keys start with `prefix`. */
  getEntriesByPrefix(prefix: string): Array<{ key: string; data: unknown }> {
    const results: Array<{ key: string; data: unknown }> = [];
    for (const [key, entry] of this.cache.entries()) {
      if (!key.startsWith(prefix)) continue;
      if (Date.now() - entry.timestamp > this.CACHE_TTL) continue;
      results.push({ key, data: entry.data });
    }
    return results;
  }

  /** Cache entries for `prefix` regardless of age — used to patch Redis snapshots after SharePoint writes. */
  getStaleEntriesByPrefix(prefix: string): Array<{ key: string; data: unknown }> {
    const results: Array<{ key: string; data: unknown }> = [];
    for (const [key, entry] of this.cache.entries()) {
      if (!key.startsWith(prefix)) continue;
      results.push({ key, data: entry.data });
    }
    return results;
  }

  private enqueuePersist(op: () => Promise<void>): void {
    this.persistQueue = this.persistQueue.then(op).catch(() => {});
  }
}