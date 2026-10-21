import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { AppConstants } from '../app.constants';
import { isTodoProcurementTaskList } from '../hr-task-lists.config';
import { graphGetWithRetry, toGraphPath } from '../microsoft-graph';
import { HrTaskMapperService } from './hr-task-mapper.service';
import { SharePointTaskQueryService } from './sharepoint-task-query.service';
import { TaskLoaderService } from './task-loader.service';

/**
 * Pages through SharePoint task lists for To Do and remembers where each list stopped.
 *
 * The paging state (cursors, watermarks, which lists were scanned) lives here because
 * only these fetchers write it; app.ts reads/clears it when a To Do load starts or resets.
 * (The task cache is TodoTaskStoreService; task actions are TodoService.)
 */
@Injectable({ providedIn: 'root' })
export class TodoFetchService {
  private readonly http = inject(HttpClient);
  private readonly hrTaskMapper = inject(HrTaskMapperService);
  private readonly taskQuery = inject(SharePointTaskQueryService);
  private readonly taskLoader = inject(TaskLoaderService);

  /**
   * Where each To Do list's paging stopped, keyed by list name. Normally empty after a
   * full load (every nextLink drained). A non-null cursor means a rare partial remaining
   * for the "Load remaining tasks" retry.
   */
  readonly todoListCursors = new Map<string, string | null>();

  /**
   * Newest `lastModifiedDateTime` already ingested per list. The manual refresh pages
   * newest-first and stops at the first row at or below this mark, so it downloads only
   * what changed since the last look instead of re-collecting and re-filtering the lot.
   */
  readonly todoListWatermarks = new Map<string, string>();

  /** Proc* lists whose first scan stopped after pages with nothing assigned to this user. */
  readonly todoProcScanStoppedEarly = new Set<string>();

  /** Assignee-blocked lists already page-scanned by this To Do load's first pass. */
  readonly todoBlockedListsScanned = new Set<string>();

  async fetchItemsForList(
    listName: string,
    siteListsArr: any[],
    siteId: string,
    siteWebUrl: string,
    token: string,
    userEmail: string,
    userUpn: string,
    range: { createdSinceIso?: string; createdBeforeIso?: string } | null = null,
    subordinateTasksOnly = false,
    fastLoadPageLimit?: number,
    assigneeLookupId: string | null = null,
    todoAssigneeOnly = false,
    todoScan = false,
    onPartial?: (items: any[]) => void,
  ): Promise<any[]> {
    const list = this.taskLoader.findSiteList(siteListsArr, listName);
    if (!list?.id) return [];

    // A list already known to reject the assignee filter (remembered from earlier loads)
    // is page-scanned in this first pass. It is ranked ahead of the slow Proc* lists, so
    // its cards show early instead of after the whole pass and a second fallback pass.
    // Procurement lists have their own Progress=Pending path below.
    const scanBlockedListNow =
      todoAssigneeOnly &&
      !!assigneeLookupId &&
      !isTodoProcurementTaskList(listName) &&
      this.taskQuery.blockedAssigneeLookupListIds.has(list.id);

    const isOlderBackfill = !!range?.createdBeforeIso;
    const top = isOlderBackfill ? 999 : AppConstants.hrTasksFastLoadPageSize;
    const maxPages = isOlderBackfill
      ? Number.POSITIVE_INFINITY
      : (fastLoadPageLimit ?? AppConstants.hrTasksFastLoadPageLimit);

    let rawItems: any[];
    let skipUserFilter = false;
    try {
      if (scanBlockedListNow) {
        this.todoBlockedListsScanned.add(listName);
        rawItems = await this.fetchTodoScanPages(siteId, list.id, token, listName);
      } else if (todoAssigneeOnly && isTodoProcurementTaskList(listName)) {
        // Graph cannot filter these lists by assignee at all, so narrow on pending +
        // newest-modified instead and match the assignee client-side. Follow every
        // Progress=Pending page (Live parity); unordered fallback stays page-capped.
        // Read in small slices and hand each one to onPartial, so rows appear steadily
        // instead of after the whole 30-page scan finishes (which looked like one burst).
        const pendingItems = await this.fetchProcurementTodoPendingSliced(
          siteId,
          list.id,
          token,
          listName,
          (sliceRaw) => {
            const shapedSlice: any[] = [];
            for (const raw of sliceRaw) {
              const shaped = this.hrTaskMapper.mapSharePointItemToHrTask(
                raw, listName, list, siteWebUrl, userEmail, userUpn, false, subordinateTasksOnly,
              );
              if (shaped) shapedSlice.push(shaped);
            }
            if (onPartial && shapedSlice.length > 0) onPartial(shapedSlice);
            // Rows To Do actually shows; the early stop counts these, not every kept row.
            return shapedSlice.filter(task => task.isAssignedToCurrentUser === true).length;
          },
        );

        rawItems = pendingItems ?? await this.taskQuery.fetchSharePointListPages(
          siteId,
          list.id,
          token,
          top,
          AppConstants.hrTasksTodoProcPageLimit,
          true,
        );
        skipUserFilter = false;
      } else if (assigneeLookupId && !subordinateTasksOnly) {
        if (!this.taskQuery.blockedAssigneeLookupListIds.has(list.id)) {
          const lookupItems = await this.taskQuery.fetchSharePointListItemsForAssigneeLookup(
            siteId,
            list.id,
            token,
            assigneeLookupId,
          );
          if (lookupItems != null) {
            if (lookupItems.length > 0) {
              rawItems = lookupItems;
              skipUserFilter = true;
            } else {
              rawItems = lookupItems;
              skipUserFilter = true;
            }
          } else if (todoAssigneeOnly) {
            return [];
          } else {
            rawItems = todoScan
              ? await this.fetchTodoScanPages(siteId, list.id, token, listName)
              : await this.taskQuery.fetchSharePointListPages(
                siteId,
                list.id,
                token,
                top,
                maxPages,
                false,
              );
          }
        } else if (todoAssigneeOnly) {
          return [];
        } else {
          // The blocked-lists To Do pass lands here - the bulk of To Do's requests.
          rawItems = todoScan
            ? await this.fetchTodoScanPages(siteId, list.id, token, listName)
            : await this.taskQuery.fetchSharePointListPages(
              siteId,
              list.id,
              token,
              top,
              maxPages,
              false,
            );
        }
      } else {
        rawItems = todoScan
          ? await this.fetchTodoScanPages(siteId, list.id, token, listName)
          : await this.taskQuery.fetchSharePointListPages(
            siteId,
            list.id,
            token,
            top,
            maxPages,
            false,
          );
      }
    } catch {
      return [];
    }

    const items: any[] = [];
    for (const item of rawItems) {
      const shaped = this.hrTaskMapper.mapSharePointItemToHrTask(
        item,
        listName,
        list,
        siteWebUrl,
        userEmail,
        userUpn,
        skipUserFilter,
        subordinateTasksOnly
      );
      if (shaped) items.push(shaped);
    }
    return items;
  }

  /**
   * To Do scan for a list that cannot be filtered by assignee: newest-modified first.
   * Hard-capped via `todoInitialPagesPerList`; leftover pages stay on `todoListCursors`
   * for manual "Load older" instead of draining the whole list on login.
   */
  async fetchTodoScanPages(
    siteId: string,
    listId: string,
    token: string,
    listName: string,
    resumePath?: string,
    maxPages: number = AppConstants.todoInitialPagesPerList,
  ): Promise<any[]> {
    const startPath = resumePath ?? this.buildTodoListStartPath(siteId, listId, listName);

    try {
      const { items, nextLink } = await this.taskQuery.fetchGraphItemPagesWithCursor(
        startPath,
        token,
        maxPages,
      );
      if (nextLink) {
        this.todoListCursors.set(listName, nextLink);
      } else {
        this.todoListCursors.delete(listName);
      }
      this.noteTodoWatermark(listName, items);
      return items;
    } catch {
      // List rejected the sort - fall back to an unordered full scan (no cursor).
      this.todoListCursors.delete(listName);
      return this.taskQuery.fetchSharePointListPages(
        siteId,
        listId,
        token,
        AppConstants.hrTasksFastLoadPageSize,
        AppConstants.hrTasksTodoLoadPageLimit,
        false,
      );
    }
  }

  async fetchProcurementTodoPendingItems(
    siteId: string,
    listId: string,
    token: string,
    listName: string,
    resumePath?: string,
    maxPages: number = AppConstants.todoProcInitialPages,
  ): Promise<any[] | null> {
    if (this.taskQuery.getListFilterFieldStatus(listId, 'ProgressPendingOrdered') === false) {
      return null;
    }

    const startPath = resumePath ?? this.buildTodoListStartPath(siteId, listId, listName);

    try {
      const { items, nextLink } = await this.taskQuery.fetchGraphItemPagesWithCursor(
        startPath,
        token,
        maxPages,
        { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
      );
      this.taskQuery.markListFilterField(listId, 'ProgressPendingOrdered', true);
      if (nextLink) {
        this.todoListCursors.set(listName, nextLink);
      } else {
        this.todoListCursors.delete(listName);
      }
      this.noteTodoWatermark(listName, items);
      return items;
    } catch (err: any) {
      const status = err?.status ?? err?.error?.status;
      if (status === 400 || status === 422) {
        this.taskQuery.markListFilterField(listId, 'ProgressPendingOrdered', false);
        this.todoListCursors.delete(listName);
        return null;
      }
      throw err;
    }
  }

  /**
   * First slice of a procurement list (todoProcInitialPages pages) read a few pages at a
   * time. `onSlice` gets each slice's raw rows as soon as it lands; the full set is still
   * returned (null = Progress filter unavailable, same as fetchProcurementTodoPendingItems).
   * Any remaining cursor stays in todoListCursors for the background drain.
   */
  async fetchProcurementTodoPendingSliced(
    siteId: string,
    listId: string,
    token: string,
    listName: string,
    onSlice: (rawItems: any[]) => number,
  ): Promise<any[] | null> {
    const slicePages = AppConstants.todoProcSlicePages;
    let pagesLeft = AppConstants.todoProcInitialPages;
    let resumePath: string | undefined;
    let emptySlices = 0;
    const all: any[] = [];

    while (pagesLeft > 0) {
      const pages = Math.min(slicePages, pagesLeft);
      const slice = await this.fetchProcurementTodoPendingItems(
        siteId, listId, token, listName, resumePath, pages,
      );
      if (slice === null) return all.length > 0 ? all : null;
      all.push(...slice);
      const assigned = slice.length > 0 ? onSlice(slice) : 0;
      emptySlices = assigned > 0 ? 0 : emptySlices + 1;
      pagesLeft -= pages;
      resumePath = this.todoListCursors.get(listName) ?? undefined;
      if (!resumePath) break;
      if (emptySlices >= AppConstants.todoProcEmptySlicesBeforeStop) {
        // Keep the cursor for "Load older", but leave this list out of the background drain.
        this.todoProcScanStoppedEarly.add(listName);
        console.info(`[todo] ${listName}: stopped after ${all.length} pending rows (none assigned in the last pages)`);
        break;
      }
    }
    return all;
  }

  /**
   * The one place that decides how a To Do list is queried, so the first load,
   * "Load older tasks" and the manual refresh can never drift apart. Always
   * newest-modified first; procurement lists additionally narrow to pending rows.
   */
  buildTodoListStartPath(siteId: string, listId: string, listName: string): string {
    const order = `&$orderby=${encodeURIComponent('lastModifiedDateTime desc')}`;
    const top = AppConstants.hrTasksFastLoadPageSize;

    if (
      isTodoProcurementTaskList(listName) &&
      this.taskQuery.getListFilterFieldStatus(listId, 'ProgressPendingOrdered') !== false
    ) {
      return `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${top}`
        + `&$filter=${encodeURIComponent("fields/Progress eq 'Pending'")}${order}`;
    }

    return `/sites/${siteId}/lists/${listId}/items?$expand=fields`
      + `&$top=${top}${order}`;
  }

  todoPreferHeader(listName: string): Record<string, string> | undefined {
    return isTodoProcurementTaskList(listName)
      ? { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }
      : undefined;
  }

  /** Advance a list's high-water mark to the newest row seen. */
  noteTodoWatermark(listName: string, items: any[]): void {
    let newest = this.todoListWatermarks.get(listName) ?? '';
    for (const item of items) {
      const modified = String(item?.lastModifiedDateTime ?? '');
      if (modified && modified > newest) newest = modified;
    }
    if (newest) this.todoListWatermarks.set(listName, newest);
  }

  /**
   * Rows added or changed since the last time this list was read. Because results are
   * ordered newest-first, the first already-known row means everything after it is known
   * too - so this stops there rather than paging the whole list again.
   */
  async fetchTodoRowsNewerThanWatermark(
    siteId: string,
    listId: string,
    token: string,
    listName: string,
    watermarkIso: string,
  ): Promise<any[]> {
    const prefer = this.todoPreferHeader(listName);
    // Procurement lists normally load Pending rows only. A top-up must also see rows
    // that turned Complete since the last look, or their old Pending copy stays in the
    // To Do snapshot forever; so ask for every recently changed row (newest first).
    const allChangedPath = isTodoProcurementTaskList(listName)
      ? `/sites/${siteId}/lists/${listId}/items?$expand=fields`
      + `&$top=${AppConstants.hrTasksFastLoadPageSize}`
      + `&$orderby=${encodeURIComponent('lastModifiedDateTime desc')}`
      : null;
    let nextPath: string | null = allChangedPath ?? this.buildTodoListStartPath(siteId, listId, listName);
    const fresh: any[] = [];
    let pages = 0;

    while (nextPath && pages < AppConstants.todoRefreshMaxPagesPerList) {
      let page: any;
      try {
        page = await graphGetWithRetry(
          this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs, prefer ?? {},
        );
      } catch (err) {
        // List refused the unfiltered query (e.g. view threshold): fall back to Pending only.
        if (pages > 0 || nextPath !== allChangedPath) throw err;
        nextPath = this.buildTodoListStartPath(siteId, listId, listName);
        continue;
      }

      let reachedKnownRows = false;
      for (const item of page?.value ?? []) {
        const modified = String(item?.lastModifiedDateTime ?? '');
        if (modified && modified <= watermarkIso) {
          reachedKnownRows = true;
          break;
        }
        fresh.push(item);
      }

      if (reachedKnownRows) break;
      nextPath = toGraphPath(page?.['@odata.nextLink']);
      pages += 1;
    }

    return fresh;
  }
}
