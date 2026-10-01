import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { AppConstants } from '../app.constants';
import { graphGetWithRetry, toGraphPath } from '../microsoft-graph';
import { sharePointConfig } from '../sharepoint.config';
import { AuthService } from './auth.service';
import { FileCrawlCacheService } from './file-crawl.service';
import { SiteMetadataService } from './site-metadata.service';

export interface DomainUserInfo {
  fullName: string;
  pinNo: string;
}

export interface DomainUserLookup {
  /** Keyed by lower-cased email (column ADmail). */
  byEmail: Map<string, DomainUserInfo>;
  /** Keyed by PinNo without leading zeros. PINs shared by several rows are left out. */
  byPin: Map<string, DomainUserInfo>;
}

/** One AllDomainUsers row as stored in the cache: [email, fullName, pinNo]. */
type DomainUserRow = [string, string, string];

interface DomainUserColumns {
  email: string;
  fullName: string;
  pinNo: string;
  /** True when all three internal names were read from the list's column definitions. */
  resolved: boolean;
}

/** Display names tried for each value, most likely first. */
const COLUMN_CANDIDATES = {
  email: ['ADmail', 'Email', 'EMail'],
  fullName: ['FullName', 'DisplayName'],
  pinNo: ['PinNo', 'PIN'],
};

const ROWS_CACHE_KEY = 'domain-users:v2';
/** Safety stop for the paging loop only - far above the real size of the list. */
const MAX_PAGES = 500;

const normalizeKey = (value: string): string => value.replace(/[^a-z0-9]/gi, '').toLowerCase();

function buildLookup(rows: DomainUserRow[]): DomainUserLookup {
  const lookup: DomainUserLookup = { byEmail: new Map(), byPin: new Map() };
  const sharedPins = new Set<string>();
  for (const [email, fullName, pinNo] of rows) {
    const user: DomainUserInfo = { fullName, pinNo };
    if (email) lookup.byEmail.set(email, user);

    const pinKey = pinNo.replace(/^0+/, '');
    if (!pinKey || sharedPins.has(pinKey)) continue;
    if (lookup.byPin.has(pinKey)) {
      lookup.byPin.delete(pinKey);
      sharedPins.add(pinKey);
    } else {
      lookup.byPin.set(pinKey, user);
    }
  }
  return lookup;
}

/**
 * Reads the AllDomainUsers SharePoint list once and exposes { FullName, PinNo } lookups
 * keyed by email and by PIN. The rows are kept in the crawl cache so names paint
 * straight away on the next visit instead of waiting on SharePoint.
 */
@Injectable({ providedIn: 'root' })
export class DomainUserService {
  private cache: Promise<DomainUserLookup> | null = null;
  private loaded: DomainUserLookup | null = null;

  constructor(
    private readonly http: HttpClient,
    private readonly authService: AuthService,
    private readonly siteMetadata: SiteMetadataService,
    private readonly fileCrawlCache: FileCrawlCacheService,
  ) { }

  /** Whatever is available right now without a network call (possibly stale), or null. */
  peekUsers(): DomainUserLookup | null {
    if (this.loaded) return this.loaded;
    const rows = this.fileCrawlCache.getStale<DomainUserRow[]>(ROWS_CACHE_KEY);
    return Array.isArray(rows) && rows.length ? buildLookup(rows) : null;
  }

  getUsers(): Promise<DomainUserLookup> {
    if (!this.cache) {
      const pending = this.load()
        .catch((err) => {
          console.warn('[DomainUserService] Could not load AllDomainUsers', err);
          return buildLookup([]);
        })
        .then((lookup) => {
          if (lookup.byEmail.size || lookup.byPin.size) {
            this.loaded = lookup;
          } else if (this.cache === pending) {
            this.cache = null; // an empty result is not cached, so the next caller tries again
          }
          return lookup;
        });
      this.cache = pending;
    }
    return this.cache;
  }

  private async load(): Promise<DomainUserLookup> {
    const fresh = this.fileCrawlCache.get<DomainUserRow[]>(ROWS_CACHE_KEY);
    if (Array.isArray(fresh) && fresh.length) return buildLookup(fresh);

    const listName = sharePointConfig.allDomainUsersListDisplayName;
    const token = await this.authService.acquireSharePointToken();

    let list = await this.siteMetadata.findList(token, listName);
    if (!list) {
      const direct = await this.siteMetadata.lookupListDirect(token, listName);
      list = direct.list;
      if (!list) {
        console.warn(`[DomainUserService] List "${listName}" not available (${direct.status})`);
        return buildLookup([]);
      }
    }

    const { siteId } = await this.siteMetadata.resolve(token);
    const columns = await this.resolveColumns(siteId, list.id, token);
    // Only the three columns we need - the full rows carry ~35 columns each. Without the
    // real internal names every field is fetched, because selecting an unknown one fails.
    const fields = columns.resolved
      ? `fields($select=${[columns.email, columns.fullName, columns.pinNo].map(encodeURIComponent).join(',')})`
      : 'fields';
    let nextPath: string | null =
      `/sites/${siteId}/lists/${list.id}/items?$select=id&$expand=${fields}&$top=999`;
    let pages = 0;
    const rows: DomainUserRow[] = [];
    while (nextPath && pages < MAX_PAGES) {
      const page: any = await this.graphGet(nextPath, token, AppConstants.graphFileListingTimeoutMs);
      for (const row of page?.value ?? []) {
        const f = row?.fields ?? {};
        const pin = f[columns.pinNo]; // Number column in SharePoint
        const email = String(f[columns.email] ?? '').trim().toLowerCase();
        const fullName = String(f[columns.fullName] ?? '').trim();
        const pinNo = pin === null || pin === undefined || pin === '' ? '' : String(pin).trim().replace(/\.0+$/, '');
        if (email || pinNo) rows.push([email, fullName, pinNo]);
      }
      nextPath = toGraphPath(page?.['@odata.nextLink']);
      pages += 1;
    }

    const lookup = buildLookup(rows);
    if (rows.length) this.fileCrawlCache.set(ROWS_CACHE_KEY, rows);
    console.info(
      `[DomainUserService] Loaded ${rows.length} rows in ${pages} pages (${lookup.byEmail.size} emails / ${lookup.byPin.size} PINs) from ${listName}`,
    );
    return lookup;
  }

  /** Front of the Graph queue: the HR Files names should not wait behind background warm-up. */
  private graphGet(path: string, token: string, timeoutMs: number): Promise<any> {
    return graphGetWithRetry(this.http, path, token, timeoutMs, {}, 4, true);
  }

  /** Internal names can differ from the display names (e.g. "field_3"), so read them from the list. */
  private async resolveColumns(siteId: string, listId: string, token: string): Promise<DomainUserColumns> {
    let defined: Array<{ name?: string; displayName?: string }> = [];
    try {
      const resp: any = await this.graphGet(
        `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`,
        token,
        AppConstants.graphDefaultTimeoutMs,
      );
      defined = resp?.value ?? [];
    } catch (err) {
      console.warn('[DomainUserService] Could not read AllDomainUsers columns', err);
    }

    const find = (candidates: string[]): string | undefined => {
      for (const candidate of candidates) {
        const target = normalizeKey(candidate);
        const column = defined.find(
          (c) => normalizeKey(c.name ?? '') === target || normalizeKey(c.displayName ?? '') === target,
        );
        if (column?.name) return column.name;
      }
      return undefined;
    };

    const email = find(COLUMN_CANDIDATES.email);
    const fullName = find(COLUMN_CANDIDATES.fullName);
    const pinNo = find(COLUMN_CANDIDATES.pinNo);
    const resolved = !!email && !!fullName && !!pinNo;
    if (!resolved) {
      console.warn('[DomainUserService] AllDomainUsers columns not all found', { email, fullName, pinNo });
    }
    return {
      email: email ?? COLUMN_CANDIDATES.email[0],
      fullName: fullName ?? COLUMN_CANDIDATES.fullName[0],
      pinNo: pinNo ?? COLUMN_CANDIDATES.pinNo[0],
      resolved,
    };
  }
}
