import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { graphGetWithRetry } from '../microsoft-graph';
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
    const listsResp: any = await graphGetWithRetry(
      this.http,
      `/sites/${site.id}/lists?$select=id,name,displayName,webUrl,list,parentReference,system`,
      token,
      AppConstants.graphFileListingTimeoutMs,
    );

    return {
      siteId: site.id,
      siteWebUrl: site.webUrl ?? '',
      lists: listsResp?.value ?? [],
    };
  }
}
