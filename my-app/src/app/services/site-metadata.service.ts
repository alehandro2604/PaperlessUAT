import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { graphGetWithRetry, toGraphPath } from '../microsoft-graph';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';

export interface SiteMetadata {
  siteId: string;
  siteWebUrl: string;
  lists: any[];
}

/**
 * One resolution of the SharePoint site id and its list metadata, shared by everything
 * that needs it.
 *
 * Each feature used to resolve these for itself, so a single page load issued the same
 * site lookup and /lists listing several times over. Both answers are fixed for the
 * session, so they are fetched once and handed out; concurrent callers await the same
 * in-flight promise rather than starting their own.
 *
 * Uses graphGetWithRetry so these bootstrap calls share the app-wide Graph slot and
 * back off on HTTP 429 instead of bypassing the throttle with bare graphGet.
 */
@Injectable({ providedIn: 'root' })
export class SiteMetadataService {
  private cached: SiteMetadata | null = null;
  private inFlight: Promise<SiteMetadata> | null = null;

  constructor(private readonly http: HttpClient) {}

  async resolve(token: string): Promise<SiteMetadata> {
    if (this.cached) return this.cached;
    if (this.inFlight) return this.inFlight;

    this.inFlight = this.fetch(token);
    try {
      this.cached = await this.inFlight;
      return this.cached;
    } finally {
      this.inFlight = null;
    }
  }

  /** List metadata by internal name or display name. */
  async findList(token: string, listName: string): Promise<any | undefined> {
    const target = String(listName ?? '').trim().toLowerCase();
    if (!target) return undefined;

    const { lists } = await this.resolve(token);
    return lists.find(
      (list: any) =>
        String(list?.name ?? '').toLowerCase() === target ||
        String(list?.displayName ?? '').toLowerCase() === target,
    );
  }

  reset(): void {
    this.cached = null;
    this.inFlight = null;
  }

  private async fetch(token: string): Promise<SiteMetadata> {
    const site: any = await graphGetWithRetry(
      this.http,
      `/sites/${sharePointConfig.siteHostName}:/${sharePointConfig.sitePath}:?$select=id,webUrl`,
      token,
      AppConstants.graphDefaultTimeoutMs,
    );
    if (!site?.id) throw new Error('Could not find SharePoint site');

    // Superset of the columns callers ask for (incl. system lists for User Information List).
    // Follow @odata.nextLink — a single page can omit lists on large sites.
    const lists: any[] = [];
    let nextPath: string | null =
      `/sites/${site.id}/lists?$select=id,name,displayName,webUrl,list,parentReference,system`;
    while (nextPath) {
      const page: any = await graphGetWithRetry(
        this.http,
        nextPath,
        token,
        AppConstants.graphFileListingTimeoutMs,
      );
      lists.push(...(page?.value ?? []));
      nextPath = toGraphPath(page?.['@odata.nextLink']);
    }

    return {
      siteId: site.id,
      siteWebUrl: site.webUrl ?? '',
      lists,
    };
  }

  /**
   * Direct lookup for a list missing from the /lists enumeration. Graph accepts a
   * list title in place of the id. Returns the list, or why it could not be read
   * (403 = this user has no access; 404 = no such list on the site).
   */
  async lookupListDirect(
    token: string,
    listName: string,
  ): Promise<{ list: any | null; status: 'ok' | 'forbidden' | 'not-found' | 'error' }> {
    const { siteId } = await this.resolve(token);
    try {
      const list: any = await graphGetWithRetry(
        this.http,
        `/sites/${siteId}/lists/${encodeURIComponent(listName)}?$select=id,name,displayName,webUrl,list,parentReference,system`,
        token,
        AppConstants.graphDefaultTimeoutMs,
      );
      if (!list?.id) return { list: null, status: 'not-found' };
      // Remember it so later lookups in this session find it without another call.
      if (this.cached && !this.cached.lists.some((l: any) => l?.id === list.id)) {
        this.cached.lists.push(list);
      }
      return { list, status: 'ok' };
    } catch (err: any) {
      const status = Number(err?.status ?? err?.statusCode);
      if (status === 403 || status === 401) return { list: null, status: 'forbidden' };
      if (status === 404) return { list: null, status: 'not-found' };
      return { list: null, status: 'error' };
    }
  }
}
