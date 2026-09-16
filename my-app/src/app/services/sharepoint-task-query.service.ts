// ============================================================
// SHAREPOINT TASK QUERY
// Every Graph query used to pull HR / To Do task rows, plus the
// per-list caches that keep those queries cheap: which columns
// accept $filter, which lists already rejected AssignedToLookupId,
// and the User Information List lookup used to resolve people.
//
// Extracted from app.ts. Site context is pushed in via
// setSiteContext() by the component's ensureSiteMetadata, which
// still owns the site id / list mirrors the rest of the UI reads.
// ============================================================
import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { AppConstants } from '../app.constants';
import { sharePointConfig } from '../sharepoint.config';
import { graphGet, graphGetWithRetry, toGraphPath } from '../microsoft-graph';
import { buildPersonLookupFilterExpr } from '../utils/hr-task-matching';
import { getEmailFromHrFolderName } from '../utils/person-name-matching';
import { SubordinaryTaskService } from './subordinaryTask.service';

@Injectable({ providedIn: 'root' })
export class SharePointTaskQueryService {
  constructor(
    private readonly http: HttpClient,
    private readonly subordinaryTaskService: SubordinaryTaskService,
  ) { }


  async fetchSharePointListItemsByTitleEmail(
    siteId: string,
    listId: string,
    token: string,
    email: string,
    options?: {
      maxPages?: number;
      onPage?: (items: any[]) => void;
      /** HRPersonal folder pin, e.g. "5852" from "5852 Aidan Harrington". */
      folderPin?: string;
    },
  ): Promise<any[] | null> {
    const prefer = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
    const filters: Array<{ field: string; expr: string }> = [];

    const value = email.trim().toLowerCase();
    if (value.includes('@')) {
      const escaped = value.replace(/'/g, "''");
      filters.push(
        { field: 'TitleContains', expr: `contains(fields/Title,'${escaped}')` },
        { field: 'TitleEq', expr: `fields/Title eq '${escaped}'` },
      );
      const local = escaped.split('@')[0] ?? '';
      const pin = String(options?.folderPin ?? '').trim();
      if (pin && local) {
        filters.push({
          field: 'TitlePinEmail',
          expr: `contains(fields/Title,'${pin.replace(/'/g, "''")} ${local.replace(/'/g, "''")}')`,
        });
      }
    }

    const pinOnly = String(options?.folderPin ?? '').trim();
    if (pinOnly) {
      filters.push({
        field: 'TitlePinPrefix',
        expr: `startswith(fields/Title,'${pinOnly.replace(/'/g, "''")} ')`,
      });
    }

    if (filters.length === 0) return null;

    const byId = new Map<string, any>();
    let anyOk = false;
    const maxPages = options?.maxPages ?? AppConstants.hrFilesPersonLookupMaxPages;
    const pageSize = AppConstants.hrFilesPersonLookupPageSize;

    for (const { field, expr } of filters) {
      if (this.getListFilterFieldStatus(listId, field) === false) continue;
      let nextPath: string | null =
        `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${pageSize}` +
        `&$filter=${encodeURIComponent(expr)}`;
      let pages = 0;
      try {
        while (nextPath && pages < maxPages) {
          const page: any = await graphGetWithRetry(
            this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs, prefer,
          );
          this.markListFilterField(listId, field, true);
          anyOk = true;
          const pageItems: any[] = [];
          for (const item of page?.value ?? []) {
            if (item?.id == null) continue;
            byId.set(String(item.id), item);
            pageItems.push(item);
          }
          pages += 1;
          if (pageItems.length) options?.onPage?.(pageItems);
          nextPath = toGraphPath(page?.['@odata.nextLink']);
        }
      } catch (err: any) {
        if ((err?.status ?? err?.error?.status) === 400) {
          this.markListFilterField(listId, field, false);
        }
      }
    }
  
    return anyOk ? [...byId.values()] : null;
  }

  private siteId: string | null = null;
  private siteLists: any[] = [];
  private cachedUserInformationListId: string | null = null;
  /** SharePoint User Information List Id keyed by lowercase email. */
  private readonly sharePointUserLookupIdByEmail = new Map<string, string>();
  /**
   * Per-list Graph $filter field viability: true = filterable, false = previously 400'd.
   * Avoids re-probing known-bad columns on every folder/task load (console flood).
   * Also persisted to localStorage so a refresh does not re-spam DevTools with 400s.
   */
  private readonly listFilterFieldStatus = new Map<string, Map<string, boolean>>();
  private listFilterStatusHydrated = false;
  /** Person/group columns discovered via Graph columns API (avoids filter probes that 400). */
  private readonly listPersonColumnsCache = new Map<string, Array<{ name: string; allowMultiple: boolean }>>();
  /** Lists where AssignedToLookupId filter already returned 400 - skip on future To Do loads. */
  private readonly assigneeLookupBlockedListIds = new Set<string>();
  private readonly sharePointListOrderBySupported = new Map<string, boolean>();

  /** Lists whose AssignedToLookupId filter already 400'd, read by To Do scope decisions. */
  get blockedAssigneeLookupListIds(): ReadonlySet<string> {
    return this.assigneeLookupBlockedListIds;
  }

  /** Pushed in by the component once the SharePoint site metadata resolves. */
  setSiteContext(siteId: string | null, lists: any[]): void {
    this.siteId = siteId;
    this.siteLists = lists ?? [];
  }

  /** Drops every per-user cache on sign-out. */
  reset(): void {
    this.siteId = null;
    this.siteLists = [];
    this.cachedUserInformationListId = null;
    this.sharePointUserLookupIdByEmail.clear();
    this.listFilterFieldStatus.clear();
    this.assigneeLookupBlockedListIds.clear();
    this.listPersonColumnsCache.clear();
    this.listFilterStatusHydrated = false;
  }
  getListFilterFieldStatus(listId: string, field: string): boolean | undefined {
    this.ensureListFilterStatusHydrated();
    return this.listFilterFieldStatus.get(listId)?.get(field);
  }

  markListFilterField(listId: string, field: string, works: boolean, persist = true): void {
    this.ensureListFilterStatusHydrated();
    let byField = this.listFilterFieldStatus.get(listId);
    if (!byField) {
      byField = new Map();
      this.listFilterFieldStatus.set(listId, byField);
    }
    byField.set(field, works);
    if (field === 'AssignedToLookupId' && works === false) {
      this.assigneeLookupBlockedListIds.add(listId);
    }
    if (persist) this.persistListFilterStatus();
  }

  private listFilterStatusLsKey(): string {
    return `${AppConstants.listFilterFieldLsPrefix}${sharePointConfig.siteId || 'default'}`;
  }

  private ensureListFilterStatusHydrated(): void {
    if (this.listFilterStatusHydrated) return;
    this.listFilterStatusHydrated = true;
    try {
      const raw = localStorage.getItem(this.listFilterStatusLsKey());
      if (!raw) return;
      const parsed = JSON.parse(raw) as { fields?: Record<string, Record<string, boolean>> };
      for (const [listId, byField] of Object.entries(parsed.fields ?? {})) {
        for (const [field, ok] of Object.entries(byField)) {
          this.markListFilterField(listId, field, ok === true, false);
        }
      }
    } catch {
      // Corrupt cache - ignore and rebuild from live probes / column metadata.
    }
  }

  private persistListFilterStatus(): void {
    try {
      const fields: Record<string, Record<string, boolean>> = {};
      for (const [listId, byField] of this.listFilterFieldStatus.entries()) {
        fields[listId] = Object.fromEntries(byField.entries());
      }
      localStorage.setItem(this.listFilterStatusLsKey(), JSON.stringify({ fields }));
    } catch {
      // Quota / private mode - in-memory map still prevents repeat probes this session.
    }
  }

  /** Known-good fields first, then untried; skip fields that already returned 400 for this list. */
  private orderFilterableFields(listId: string, candidates: string[]): string[] {
    const good: string[] = [];
    const unknown: string[] = [];
    for (const field of candidates) {
      const status = this.getListFilterFieldStatus(listId, field);
      if (status === false) continue;
      if (status === true) good.push(field);
      else unknown.push(field);
    }
    return [...good, ...unknown];
  }

  private async getListPersonColumns(
    siteId: string,
    listId: string,
    token: string,
  ): Promise<Array<{ name: string; allowMultiple: boolean }>> {
    const cached = this.listPersonColumnsCache.get(listId);
    if (cached) return cached;

    try {
      const resp: any = await firstValueFrom(
        graphGet(
          this.http,
          `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,personOrGroup`,
          token,
        ),
      );
      const cols = (resp?.value ?? [])
        .filter((col: any) => !!col?.personOrGroup && !!col?.name)
        .map((col: any) => ({
          name: String(col.name),
          allowMultiple: col.personOrGroup.allowMultipleSelection === true,
        }));
      this.listPersonColumnsCache.set(listId, cols);
      return cols;
    } catch {
      this.listPersonColumnsCache.set(listId, []);
      return [];
    }
  }


  /**
   * Resolve AssignedTo (or Assigned / field_7) filter expression from column metadata.
   * Returns null when the list has no assignee person column - avoids a Graph 400 probe.
   */
  private async resolveAssigneeLookupFilterExpr(
    siteId: string,
    listId: string,
    token: string,
    lookupId: string,
  ): Promise<string | null> {
    if (this.getListFilterFieldStatus(listId, 'AssignedToLookupId') === false) {
      return null;
    }

    const personCols = await this.getListPersonColumns(siteId, listId, token);
    const candidates = ['AssignedTo', 'Assigned', 'field_7'];
    const match = candidates
      .map(name => personCols.find(col => col.name.toLowerCase() === name.toLowerCase()))
      .find((col): col is { name: string; allowMultiple: boolean } => !!col);

    if (!match) {
      this.markListFilterField(listId, 'AssignedToLookupId', false);
      return null;
    }

    return buildPersonLookupFilterExpr(`${match.name}LookupId`, lookupId, match.allowMultiple);
  }

  /** Fetch workflow siblings that share an eForm id (manager finalised steps, etc.). */
  async fetchSharePointListItemsForEFormKeys(
    siteId: string,
    listId: string,
    token: string,
    eFormKeys: Set<string>,
  ): Promise<any[]> {
    const prefer = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
    const byId = new Map<string, any>();
    const fieldNames = this.orderFilterableFields(listId, ['eFormListId', 'ListId', 'field_11']);
    if (fieldNames.length === 0) return [];

    for (const key of eFormKeys) {
      const safeKey = key.replace(/'/g, "''");
      for (const field of fieldNames) {
        if (this.getListFilterFieldStatus(listId, field) === false) continue;
        const filter = encodeURIComponent(`fields/${field} eq '${safeKey}'`);
        let nextPath: string | null =
          `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=100&$filter=${filter}`;
        let pages = 0;
        try {
          while (nextPath && pages < 20) {
            const page: any = await graphGetWithRetry(
              this.http,
              nextPath,
              token,
              AppConstants.graphFileListingTimeoutMs,
              prefer,
            );
            this.markListFilterField(listId, field, true);
            for (const item of page?.value ?? []) {
              if (item?.id != null) byId.set(String(item.id), item);
            }
            nextPath = toGraphPath(page?.['@odata.nextLink']);
            pages += 1;
          }
        } catch (err: any) {
          const status = err?.status ?? err?.error?.status;
          // 400 = not filterable; timeout/other = don't keep retrying this field forever.
          if (status === 400 || status == null) {
            this.markListFilterField(listId, field, false);
          }
          // Field not filterable on this list - try next field/key.
        }
      }
    }

    return [...byId.values()];
  }

  /** True when a raw SharePoint item references one of the known eForm IDs (field or comment text). */


  /**
   * Email for HR Files LookupId filtering: prefer the address embedded in the
   * folder name, otherwise resolve PinNo / display name via AllDomainUsers.
   */
  async resolveHrFolderPersonEmail(folderName: string): Promise<string> {
    const fromFolder = getEmailFromHrFolderName(folderName);
    if (fromFolder) return fromFolder;

    try {
      const user = await this.subordinaryTaskService.findUserForHrFolder(folderName);
      const email = String(user?.ADmail ?? '').trim().toLowerCase();
      if (email.includes('@')) return email;
    } catch {
      // Best-effort - caller falls back to page scanning.
    }
    return '';
  }

  /** Resolve SharePoint site-local user Id (LookupId) from email via User Information List. */
  async resolveSharePointUserLookupId(email: string, token: string): Promise<string | null> {
    const emailLower = email.trim().toLowerCase();
    if (!emailLower || !this.siteId) return null;
    const cached = this.sharePointUserLookupIdByEmail.get(emailLower);
    if (cached) return cached;

    try {
      if (!this.cachedUserInformationListId) {
        const fromCache = this.siteLists.find((list: any) => {
          const name = String(list.name ?? '').toLowerCase();
          const display = String(list.displayName ?? '').toLowerCase();
          return name === 'users' || display === 'user information list' || name === 'user information list';
        });
        if (fromCache?.id) {
          this.cachedUserInformationListId = fromCache.id;
        } else {
          const listsResp: any = await graphGetWithRetry(
            this.http,
            `/sites/${this.siteId}/lists?$select=id,name,displayName,system`,
            token,
            AppConstants.graphFileListingTimeoutMs,
          );
          const userInfoList = (listsResp?.value ?? []).find((list: any) => {
            const name = String(list.name ?? '').toLowerCase();
            const display = String(list.displayName ?? '').toLowerCase();
            return name === 'users' || display === 'user information list' || name === 'user information list';
          });
          this.cachedUserInformationListId = userInfoList?.id ?? null;
        }
      }
      if (!this.cachedUserInformationListId) return null;

      const userInfoListId = this.cachedUserInformationListId;
      const emailFilterOk = this.getListFilterFieldStatus(userInfoListId, 'EMail');
      if (emailFilterOk !== false) {
        const prefer = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
        const filter = encodeURIComponent(`fields/EMail eq '${emailLower.replace(/'/g, "''")}'`);
        try {
          const filtered: any = await graphGetWithRetry(
            this.http,
            `/sites/${this.siteId}/lists/${userInfoListId}/items?$expand=fields($select=Id,EMail,Title)&$filter=${filter}&$top=5`,
            token,
            AppConstants.graphFileListingTimeoutMs,
            prefer,
          );
          this.markListFilterField(userInfoListId, 'EMail', true);
          const hit = filtered?.value?.[0];
          // Graph item id is the site-local LookupId used by person columns.
          const id = String(hit?.id ?? hit?.fields?.Id ?? '').trim();
          if (id) {
            this.sharePointUserLookupIdByEmail.set(emailLower, id);
            return id;
          }
        } catch (err: any) {
          const status = err?.status ?? err?.error?.status;
          if (status === 400) {
            this.markListFilterField(userInfoListId, 'EMail', false);
          }
          // Fall through to a short scan of the User Information List.
        }
      }

      let nextPath: string | null =
        `/sites/${this.siteId}/lists/${this.cachedUserInformationListId}/items?$expand=fields($select=Id,EMail,Title)&$top=200`;
      let pages = 0;
      while (nextPath && pages < 20) {
        const page: any = await graphGetWithRetry(
          this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs,
        );
        const match = (page?.value ?? []).find((item: any) => {
          const itemEmail = String(item.fields?.EMail ?? item.fields?.Email ?? '').toLowerCase();
          return itemEmail === emailLower;
        });
        const id = String(match?.id ?? match?.fields?.Id ?? '').trim();
        if (id) {
          this.sharePointUserLookupIdByEmail.set(emailLower, id);
          return id;
        }
        nextPath = toGraphPath(page?.['@odata.nextLink']);
        pages += 1;
      }
    } catch {
      return null;
    }
    return null;
  }

  /**
   * Load list items involving a person (submitted by / created by).
   * Returns null when no LookupId filter worked so the caller can fall back to page scanning.
   * Pages newest-first and reports each page so Comments can paint August-dated rows
   * before a 800-task drain finishes.
   */
  async fetchSharePointListItemsForPersonLookup(
    siteId: string,
    listId: string,
    token: string,
    lookupId: string,
    options?: {
      maxPages?: number;
      pageSize?: number;
      onPage?: (items: any[], pageIndex: number) => void;
    },
  ): Promise<any[] | null> {
    const prefer = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
    const preferredBases = ['Requestor', 'Author', 'SubmittedBy', 'CreatedBy', 'Employee', 'EmployeeName'];
    const skipBases = new Set([
      'assignedto', 'assigned', 'field_7', 'superior', 'manager', 'approver', 'editor',
    ]);
    const personCols = await this.getListPersonColumns(siteId, listId, token);
    const colByName = new Map(personCols.map(col => [col.name.toLowerCase(), col]));
    const candidateBases = [
      ...preferredBases,
      ...personCols
        .map(col => col.name)
        .filter(name => !skipBases.has(name.toLowerCase())),
    ].filter((name, index, all) =>
      all.findIndex(other => other.toLowerCase() === name.toLowerCase()) === index
    );

    // Mark missing preferred columns as non-filterable without a Graph $filter.
    for (const base of preferredBases) {
      const lookupField = `${base}LookupId`;
      if (this.getListFilterFieldStatus(listId, lookupField) !== undefined) continue;
      if (!colByName.has(base.toLowerCase())) {
        this.markListFilterField(listId, lookupField, false);
      }
    }

    const lookupFields = this.orderFilterableFields(
      listId,
      candidateBases.map(base => `${base}LookupId`),
    );
    if (lookupFields.length === 0) return null;

    const byId = new Map<string, any>();
    let anyFilterSucceeded = false;
    let firstFieldHit = false;
    const pageSize = options?.pageSize ?? AppConstants.hrFilesPersonLookupPageSize;
    const maxPages = options?.maxPages ?? AppConstants.hrFilesPersonLookupMaxPages;

    const pageField = async (
      field: string,
      filter: string,
      useOrderBy: boolean,
      fieldMaxPages: number,
    ): Promise<'ok' | 'bad-field' | 'retry-unordered'> => {
      const orderQuery = useOrderBy ? '&$orderby=lastModifiedDateTime desc' : '';
      let nextPath: string | null =
        `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${pageSize}${orderQuery}&$filter=${filter}`;
      let pages = 0;
      try {
        while (nextPath && pages < fieldMaxPages) {
          const page: any = await graphGetWithRetry(
            this.http,
            nextPath,
            token,
            AppConstants.graphFileListingTimeoutMs,
            prefer,
          );
          this.markListFilterField(listId, field, true);
          anyFilterSucceeded = true;
          const pageItems: any[] = [];
          for (const item of page?.value ?? []) {
            if (item?.id == null) continue;
            byId.set(String(item.id), item);
            pageItems.push(item);
          }
          pages += 1;
          if (pageItems.length > 0) {
            firstFieldHit = true;
            options?.onPage?.(pageItems, pages);
          }
          nextPath = toGraphPath(page?.['@odata.nextLink']);
        }
        return 'ok';
      } catch (err: any) {
        const status = err?.status ?? err?.error?.status;
        // Ordered+filtered queries often 400/422 on large lists. Retry without $orderby
        // before giving up on the person column itself.
        if (useOrderBy && (status === 400 || status === 422)) {
          this.sharePointListOrderBySupported.set(listId, false);
          return 'retry-unordered';
        }
        if (status === 400) {
          this.markListFilterField(listId, field, false);
        }
        return 'bad-field';
      }
    };

    for (const field of lookupFields) {
      if (this.getListFilterFieldStatus(listId, field) === false) continue;
      const baseName = field.replace(/LookupId$/i, '');
      const col = colByName.get(baseName.toLowerCase());
      const filterExpr = buildPersonLookupFilterExpr(
        field,
        lookupId,
        col?.allowMultiple === true,
      );
      const filter = encodeURIComponent(filterExpr);
      const fieldMaxPages = firstFieldHit
        ? Math.min(maxPages, AppConstants.hrFilesLookupExtraFieldPages)
        : maxPages;
      const orderOk = this.sharePointListOrderBySupported.get(listId) !== false;
      let result = await pageField(field, filter, orderOk, fieldMaxPages);
      if (result === 'retry-unordered') {
        result = await pageField(field, filter, false, fieldMaxPages);
      }
      if (result === 'bad-field') continue;
      // Keep querying other person columns. An empty-but-successful
      // RequestorLookupId must not abort Author/SubmittedBy.
    }

    return anyFilterSucceeded ? [...byId.values()] : null;
  }

  /**
   * Load list items currently assigned to the signed-in user.
   * Returns null when no AssignedTo LookupId filter works on this list.
   */
  async fetchSharePointListItemsForAssigneeLookup(
    siteId: string,
    listId: string,
    token: string,
    lookupId: string,
    options?: {
      maxPages?: number;
      onPage?: (items: any[]) => void;
    },
  ): Promise<any[] | null> {
    if (this.assigneeLookupBlockedListIds.has(listId)) return null;
    if (this.getListFilterFieldStatus(listId, 'AssignedToLookupId') === false) {
      this.assigneeLookupBlockedListIds.add(listId);
      return null;
    }

    const filterExpr = await this.resolveAssigneeLookupFilterExpr(siteId, listId, token, lookupId);
    if (!filterExpr) {
      this.assigneeLookupBlockedListIds.add(listId);
      return null;
    }

    const prefer = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
    const maxPages = options?.maxPages ?? AppConstants.hrTasksTodoLookupMaxPages;
    const pageSize = AppConstants.hrFilesPersonLookupPageSize;
    const orderOk = this.sharePointListOrderBySupported.get(listId) !== false;

    const fetchPages = async (useOrderBy: boolean): Promise<any[] | null> => {
      const byId = new Map<string, any>();
      const order = useOrderBy ? '&$orderby=lastModifiedDateTime desc' : '';
      let nextPath: string | null =
        `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${pageSize}`
        + `&$filter=${encodeURIComponent(filterExpr)}${order}`;
      let pages = 0;
      try {
        while (nextPath && pages < maxPages) {
          const page: any = await graphGetWithRetry(
            this.http,
            nextPath,
            token,
            AppConstants.graphFileListingTimeoutMs,
            prefer,
          );
          const pageItems: any[] = [];
          for (const item of page?.value ?? []) {
            if (item?.id == null) continue;
            byId.set(String(item.id), item);
            pageItems.push(item);
          }
          pages += 1;
          if (pageItems.length > 0) {
            options?.onPage?.(pageItems);
          }
          nextPath = toGraphPath(page?.['@odata.nextLink']);
        }
        return [...byId.values()];
      } catch (err: any) {
        const status = err?.status ?? err?.error?.status;
        if (useOrderBy && (status === 400 || status === 422)) {
          this.sharePointListOrderBySupported.set(listId, false);
          return null;
        }
        throw err;
      }
    };

    try {
      let items = orderOk ? await fetchPages(true) : null;
      if (items == null) {
        items = await fetchPages(false);
      }
      if (items == null) return null;
      this.markListFilterField(listId, 'AssignedToLookupId', true);
      return items;
    } catch (err: any) {
      const status = err?.status ?? err?.error?.status;
      if (status === 400) {
        this.markListFilterField(listId, 'AssignedToLookupId', false);
        this.assigneeLookupBlockedListIds.add(listId);
        return null;
      }
      throw err;
    }
  }

  /**
   * Page through Graph list items and hand back where it stopped.
   *
   * `maxPages <= 0` means follow every `@odata.nextLink` until the list ends.
   * Prefer a positive cap for To Do scans so large lists cannot flood Graph;
   * a positive cap still returns a resume cursor when more pages remain.
   */
  async fetchGraphItemPagesWithCursor(
    startPath: string,
    token: string,
    maxPages: number,
    prefer?: Record<string, string>,
  ): Promise<{ items: any[]; nextLink: string | null }> {
    const byId = new Map<string, any>();
    let nextPath: string | null = startPath;
    let pages = 0;
    const unlimited = !(maxPages > 0);

    while (nextPath && (unlimited || pages < maxPages)) {
      // Always the retrying path: these loops fan out across ~30 lists, which is
      // exactly what trips SharePoint throttling, and an un-retried 429 here would
      // silently drop a whole list's rows.
      const page: any = await graphGetWithRetry(
        this.http,
        nextPath,
        token,
        AppConstants.graphFileListingTimeoutMs,
        prefer ?? {},
      );

      for (const item of page?.value ?? []) {
        if (item?.id != null) byId.set(String(item.id), item);
      }
      nextPath = toGraphPath(page?.['@odata.nextLink']);
      pages += 1;
    }

    return { items: [...byId.values()], nextLink: nextPath };
  }

  private async fetchSharePointListItemsWithFilter(
    siteId: string,
    listId: string,
    token: string,
    filterExpr: string,
    prefer: Record<string, string>,
    maxPages = 0,
    orderByExpr?: string,
  ): Promise<any[]> {
    const byId = new Map<string, any>();
    const filter = encodeURIComponent(filterExpr);
    const order = orderByExpr ? `&$orderby=${encodeURIComponent(orderByExpr)}` : '';
    let nextPath: string | null =
      `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${AppConstants.hrTasksFastLoadPageSize}&$filter=${filter}${order}`;
    let pages = 0;
    const unlimited = !(maxPages > 0);

    while (nextPath && (unlimited || pages < maxPages)) {
      const page: any = await graphGetWithRetry(
        this.http,
        nextPath,
        token,
        AppConstants.graphFileListingTimeoutMs,
        prefer,
      );
      for (const item of page?.value ?? []) {
        if (item?.id != null) byId.set(String(item.id), item);
      }
      nextPath = toGraphPath(page?.['@odata.nextLink']);
      pages += 1;
    }

    return [...byId.values()];
  }

  /**
   * To Do items for a procurement list, newest first.
   *
   * These lists are far past SharePoint's 5k list-view threshold, and Graph rejects
   * the two obvious narrowing options on them: an assignee filter
   * (`AssignedToLookupId` - 400) and a created-date sort (`createdDateTime` - 422,
   * threshold). What Live *does* accept is `Progress eq 'Pending'` combined with
   * `lastModifiedDateTime desc`, so we narrow on those instead.
   *
   * `Progress` only ever holds `Pending` or `Complete` on these lists, and To Do
   * hides completed work anyway, so the filter costs no visible rows. Assignee
   * matching still happens client-side afterwards - this only decides which rows
   * are worth fetching.
   *
   * Returns null when the list rejects the query, so the caller can fall back.
   */

  async fetchSharePointListPages(
    siteId: string,
    listId: string,
    token: string,
    top: number,
    maxPages: number,
    newestFirst = false
  ): Promise<any[]> {
    const collected: any[] = [];
    // Live rejects $orderby=createdDateTime with 422 (list-view threshold), but accepts
    // lastModifiedDateTime - verified against ProcTasks/ProcTasksArchive/ECTasks.
    // Track this per list - one archive list must not disable newest-first for HRTask*.
    const useOrderBy = newestFirst && this.sharePointListOrderBySupported.get(listId) !== false;
    const orderQuery = useOrderBy ? '&$orderby=lastModifiedDateTime desc' : '';
    let nextPath: string | null =
      `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${top}${orderQuery}`;
    let pagesLoaded = 0;
    const unlimited = !(maxPages > 0);

    // maxPages <= 0: follow every nextLink (avoid for To Do - use positive caps).
    // Positive: early stop; caller may resume via todoListCursors.
    while (nextPath) {
      try {
        const page: any = await graphGetWithRetry(
          this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs,
        );
        collected.push(...((page?.value ?? []) as any[]));
        pagesLoaded += 1;
        if (!unlimited && pagesLoaded >= maxPages) break;
        nextPath = toGraphPath(page?.['@odata.nextLink']);
      } catch {
        if (useOrderBy && pagesLoaded === 0 && orderQuery) {
          this.sharePointListOrderBySupported.set(listId, false);
          nextPath = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${top}`;
          continue;
        }
        break;
      }
    }
    return collected;
  }

  /**
   * Populate the component's site/list caches from the shared resolver, so every feature
   * in the app resolves the site and its list metadata exactly once per session.
   */
}
