import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { sharePointConfig } from '../sharepoint.config';
import { graphGetWithRetry, toGraphPath } from '../microsoft-graph';
import { AuthService } from './auth.service';
import { AppConstants } from '../app.constants';
import { SiteMetadataService } from './site-metadata.service';


@Injectable({
    providedIn: 'root'
})
export class SharePointListService {
    constructor(
        private readonly http: HttpClient,
        private readonly authService: AuthService,
        private readonly siteMetadata: SiteMetadataService,
    ) { }

    /** Reuses the shared site-metadata cache — avoids a second /lists Graph hit on login. */
    async getAllLists(): Promise<any[]> {
        const token = await this.authService.acquireSharePointToken();
        const { lists } = await this.siteMetadata.resolve(token);
        return lists;
    }

    async getListByName(name: string) {
        const lists = await this.getAllLists();
        return lists.find(
            (l: any) =>
                (l.displayName ?? '').toLowerCase() === name.toLowerCase() ||
                (l.name ?? '').toLowerCase() === name.toLowerCase()
        );
    }

    /** Fetch list items with a hard page cap so callers cannot drain entire large lists. */
    async getListItems(listId: string, siteId = sharePointConfig.siteId): Promise<any[]> {
        const token = await this.authService.acquireSharePointToken();
        const byId = new Map<string, any>();
        let nextPath: string | null =
            `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;
        const maxPages = 10;
        let pages = 0;

        while (nextPath && pages < maxPages) {
            const page = await graphGetWithRetry(
                this.http,
                nextPath,
                token,
                AppConstants.graphFileListingTimeoutMs,
            ) as { value?: any[]; '@odata.nextLink'?: string };

            for (const item of page?.value ?? []) {
                if (item?.id != null) byId.set(String(item.id), item);
            }
            nextPath = toGraphPath(page?.['@odata.nextLink']);
            pages += 1;
        }

        return [...byId.values()];
    }
}
