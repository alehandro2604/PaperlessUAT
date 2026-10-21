import { Injectable, inject } from '@angular/core';
import { AppConstants } from '../app.constants';
import {
  HR_SOURCE_EFORM_LIST_NAMES, HR_TASK_LIST_NAMES, TODO_EXTRA_TASK_LIST_NAMES,
  isHrSourceEFormList, isHrTitleMatchedTaskList,
} from '../hr-task-lists.config';
import { AppUser } from '../models/app-user.model';
import { FileCrawlCacheService } from './file-crawl.service';
import { FormConfigurationService } from './form-configuration.service';
import { HrTaskFolderMatchService } from './hr-task-folder-match.service';
import { HrTaskMapperService } from './hr-task-mapper.service';
import { SharePointTaskQueryService } from './sharepoint-task-query.service';
import { SiteMetadataService } from './site-metadata.service';
import { hrPersonIdsCacheKey } from '../utils/cache-keys';
import { isCommentTypeItem, isSharePointCommentEntry } from '../utils/comment-utils';
import {
  collectEFormKeysFromMappedTasks, collectEFormKeyToTaskNames, doesMappedItemMatchSeedEFormIdAndTaskName,
  doesRawItemMatchSeedEFormIdAndTaskName, flattenHrMappedByList, mergeHrTasks,
} from '../utils/hr-task-matching';
import { normalizeTaskMatchText, stringifyTaskFieldValue } from '../utils/person-name-matching';
import { sleep } from '../utils/timing';

/**
 * Fetches tasks from SharePoint lists and turns them into task objects.
 *
 * It never touches the screen: callers (app.ts) decide what to show, and pass
 * small callbacks when they need progress (onPartial) or want to cancel (isStale).
 * Low-level single-list Graph queries live in SharePointTaskQueryService.
 */
@Injectable({ providedIn: 'root' })
export class TaskLoaderService {
  private readonly site = inject(SiteMetadataService);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly formConfigService = inject(FormConfigurationService);
  private readonly folderMatch = inject(HrTaskFolderMatchService);
  private readonly hrTaskMapper = inject(HrTaskMapperService);
  private readonly taskQuery = inject(SharePointTaskQueryService);

  // Site info is owned by SiteMetadataService; these read it.
  private get cachedSiteId(): string | null { return this.site.current?.siteId ?? null; }
  private get cachedSiteWebUrl(): string { return this.site.current?.siteWebUrl ?? ''; }
  private get cachedSiteLists(): any[] { return this.site.current?.lists ?? []; }

  /**
   * Site list for a configured list name. An internal-name match wins over a display-name
   * match: a list whose display name equals another list's internal name (e.g. ProcTasks1
   * shown as "ProcTasks") must not stand in for that list.
   */
  findSiteList(siteLists: any[], listName: string): any | undefined {
    const key = String(listName ?? '').toLowerCase();
    return siteLists.find((l: any) => (l.name ?? '').toLowerCase() === key)
      ?? siteLists.find((l: any) => (l.displayName ?? '').toLowerCase() === key);
  }

  /** Drop names that resolve to a list already in the set, so no list is paged twice. */
  dedupeTaskListsBySiteList(listNames: string[], siteLists: any[]): string[] {
    const seen = new Map<string, string>();
    const unique: string[] = [];
    for (const name of listNames) {
      const id = this.findSiteList(siteLists, name)?.id;
      if (!id) continue;
      const first = seen.get(id);
      if (first !== undefined) {
        console.info(`[todo] "${name}" is the same SharePoint list as "${first}"; querying it once`);
        continue;
      }
      seen.set(id, name);
      unique.push(name);
    }
    return unique;
  }

  /**
    this is for the proctasks in archive to load the tasks from the ProcTasksArchive list
   */
  getArchiveCompanionTaskLists(listName: string): string[] {
    const key = String(listName ?? '').trim().toLowerCase();
    if (key === 'proctasks' || key === 'proctasks1') {
      return ['ProcTasksArchive'];
    }
    return [];
  }

  /** Add archive companion lists (e.g. ProcTasksArchive for procurement). */
  expandTaskListsWithArchiveCompanions(listNames: string[]): string[] {
    const expanded = new Set(listNames);
    for (const name of listNames) {
      for (const archive of this.getArchiveCompanionTaskLists(name)) {
        expanded.add(archive);
      }
    }
    return [...expanded];
  }

  /** Task lists from eForms config, hard-coded names, and HRTask* auto-detect. */
  getTaskListsToQuery(
    siteLists: Array<{ name?: string }>,
    options?: { includeAllWorkflows?: boolean; includeArchiveCompanions?: boolean },
  ): string[] {
    const autoDetected = siteLists
      .filter((l) => (l.name ?? '').toLowerCase().startsWith('hrtask'))
      .map((l) => l.name as string);

    // HR Files: TaskListName + source eForm lists (URL/Prefix) so Sick Certificate
    // submissions appear in person Comments. To Do: all workflows' task + source lists.
    const fromEForms = options?.includeAllWorkflows
      ? this.formConfigService.getAllQueryListNames()
      : this.formConfigService.getHrQueryListNames();

    let names = [...new Set([
      ...HR_TASK_LIST_NAMES,
      ...HR_SOURCE_EFORM_LIST_NAMES,
      ...autoDetected,
      ...fromEForms,
    ])];

    if (options?.includeAllWorkflows) {
      names = [...new Set([
        ...names,
        ...TODO_EXTRA_TASK_LIST_NAMES,
        ...this.formConfigService.getAllQueryListNames(),
      ])];
    }

    if (options?.includeArchiveCompanions) {
      names = this.expandTaskListsWithArchiveCompanions(names);
    }

    return names;
  }

  findCachedSiteListByName(listName: string): { id: string; name?: string; displayName?: string; webUrl?: string } | null {
    const target = String(listName ?? '').trim().toLowerCase();
    if (!target) return null;
    return this.cachedSiteLists.find((l: any) =>
      (l.name ?? '').toLowerCase() === target ||
      (l.displayName ?? '').toLowerCase() === target
    ) ?? null;
  }

  mapHrFolderRawItems(
    rawItems: any[],
    listName: string,
    listObj: any,
    folderName: string,
    personLookupId: string | null,
    folderMatchHints: string[],
    userEmail: string,
    userUpn: string,
  ): any[] {
    return rawItems
      .filter((item: any) => !isSharePointCommentEntry(item))
      .filter((item: any) =>
        this.folderMatch.doesHrTaskBelongToFolderPerson(
          item,
          listName,
          folderName,
          personLookupId,
          folderMatchHints,
        )
      )
      .map((item: any) => this.hrTaskMapper.mapSharePointItemToHrTask(
        item, listName, listObj, this.cachedSiteWebUrl, userEmail, userUpn, true,
      ))
      .filter((item): item is NonNullable<typeof item> => item !== null)
      .filter((item: any) => !isCommentTypeItem(item))
      .filter((item: any) =>
        this.folderMatch.doesMappedHrTaskBelongToFolderPerson(
          item,
          folderName,
          personLookupId,
          folderMatchHints,
          listName,
        )
      );
  }

  async collectHrPersonalTasksFromLists(
    listsToQuery: string[],
    folderName: string,
    token: string,
    user: AppUser | null,
    isStale: () => boolean,
    maxPages: number,
    onPartial?: (mapped: any[]) => void,
    options?: {
      /** Called once after all lists' seed (submitter) tasks are collected - use for fast first paint. */
      onSeedComplete?: (mapped: any[]) => void;
      /** Newest LookupId pages only — merge into the existing 800+ cache instead of re-reading it. */
      softRefresh?: boolean;
      /** User click (not prefetch): parallel worker pool with no batch gaps. */
      interactive?: boolean;
    },

  ): Promise<any[]> {
    const crawlStartedAt = Date.now();
    const userEmail = (user?.email ?? '').toLowerCase();
    const userUpn = (user?.userPrincipalName ?? '').toLowerCase();
    const concurrency = AppConstants.hrFilesTaskListConcurrency;
    const softRefresh = options?.softRefresh === true;
    const folderPin = String(folderName ?? '').match(/^\d+/)?.[0] ?? '';
    // Email + LookupId never change for a folder — shared Redis skips two Graph
    // round-trips on every cold open after anyone has resolved this person once.
    const idsKey = hrPersonIdsCacheKey(folderName);
    const cachedIds =
      this.fileCrawlCache.getStale<{ email: string; lookupId: string | null }>(idsKey) ??
      (await this.fileCrawlCache.hydrateFromPersistent<{ email: string; lookupId: string | null }>(idsKey))?.data;
    const folderEmail = cachedIds?.email || await this.taskQuery.resolveHrFolderPersonEmail(folderName);
    let personLookupId: string | null = cachedIds?.email ? cachedIds.lookupId ?? null : null;
    if (personLookupId) {
      this.taskQuery.primeSharePointUserLookupId(folderEmail, personLookupId);
    } else if (folderEmail) {
      // Re-resolve a missing LookupId so one transient failure is not cached forever.
      personLookupId = await this.taskQuery.resolveSharePointUserLookupId(folderEmail, token);
    }
    if (folderEmail && (cachedIds?.email !== folderEmail || cachedIds?.lookupId !== personLookupId)) {
      this.fileCrawlCache.set(idsKey, { email: folderEmail, lookupId: personLookupId });
    }
    const folderMatchHints = folderEmail
      ? [folderEmail, folderEmail.split('@')[0] ?? '']
      : [];

    // listName -> mapped tasks belonging to / involving this person
    const mappedByList = new Map<string, any[]>();
    // listName -> list metadata for later related-eForm queries
    const listObjByName = new Map<string, any>();
    // Lists whose first-pass person lookup left rows unread — only these are re-read in the deep crawl.
    const lookupIncompleteLists = new Set<string>();
    // listName -> newest unfiltered pages already scanned in the first pass (not scanned twice).
    const scannedListPages = new Map<string, number>();
    // The deep crawl only runs with a LookupId; it then reads each list's newest pages itself.
    const deepCrawlWillRun = !!personLookupId;

    // Phase timings for clicks / refreshes (not background prefetch) — shows where a slow load spends its time.
    const crawlMode = softRefresh ? 'refresh' : options?.interactive ? 'open' : '';
    const logCrawlPhase = (phase: string, detail = ''): void => {
      if (!crawlMode) return;
      const count = flattenHrMappedByList(mappedByList).length;
      const seconds = ((Date.now() - crawlStartedAt) / 1000).toFixed(1);
      console.info(`[HR Files] ${crawlMode} "${folderName}" ${phase}: ${count} tasks at ${seconds}s${detail}`);
    };

    const publishListHits = (listName: string, mapped: any[]): void => {
      if (mapped.length === 0) return;
      const existing = mappedByList.get(listName) ?? [];
      mappedByList.set(listName, mergeHrTasks(existing, mapped));
      onPartial?.(flattenHrMappedByList(mappedByList));
    };

    const processList = async (listName: string): Promise<void> => {
      const listObj = this.cachedSiteLists.find((l: any) =>
        (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
        (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
      );
      if (!listObj?.id) return;
      listObjByName.set(listName, listObj);

      const mapRaw = (raw: any[]) => this.mapHrFolderRawItems(
        raw, listName, listObj, folderName, personLookupId, folderMatchHints, userEmail, userUpn,
      );

      let rawItems: any[] | null = null;
      const isSourceEForm = isHrSourceEFormList(listName);
      // Soft refresh: LookupId newest pages already catch new + updated rows.
      // Skip the unfiltered company-wide scan so idle polls do not double Graph traffic.
      const supplementPages = softRefresh
        ? AppConstants.hrFilesSoftRefreshSupplementPages
        : (isSourceEForm
          ? AppConstants.hrFilesSourceEFormLookupSupplementPages
          : AppConstants.hrFilesLookupSupplementPages);
      // Cold first pass: newest-first pages only so Comments paints recent tasks fast.
      // Deeper LookupId + assignee continue after onSeedComplete.
      const lookupMaxPages = softRefresh
        ? AppConstants.hrFilesSoftRefreshLookupPages
        : AppConstants.hrFilesFirstPaintLookupPages;

      // Cold click: LookupId (+ Title) only so first Comments paint isn't fighting
      // assignee + company-wide supplement for Graph slots.
      const deferHeavyFanOut = !softRefresh;

      const lookupPromise = personLookupId
        ? this.taskQuery.fetchSharePointListItemsForPersonLookup(
          this.cachedSiteId!,
          listObj.id,
          token,
          personLookupId,
          {
            maxPages: lookupMaxPages,
            onPage: (pageItems) => {
              if (isStale()) return;
              publishListHits(listName, mapRaw(pageItems));
            },
            onIncomplete: () => lookupIncompleteLists.add(listName),
          },
        )
        : Promise.resolve(null);

      // Soft refresh keeps assignee off (LookupId + Title enough). Cold defers it
      // to the post-seed top-up so first paint gets Graph capacity sooner.
      const assigneePromise =
        !deferHeavyFanOut && !softRefresh && personLookupId
          ? this.taskQuery.fetchSharePointListItemsForAssigneeLookup(
            this.cachedSiteId!,
            listObj.id,
            token,
            personLookupId,
            {
              maxPages: lookupMaxPages,
              onPage: (pageItems) => {
                if (isStale()) return;
                publishListHits(listName, mapRaw(pageItems));
              },
            },
          )
          : Promise.resolve(null);

      const titlePromise = isHrTitleMatchedTaskList(listName) && (folderEmail || folderPin)
        ? this.taskQuery.fetchSharePointListItemsByTitleEmail(
          this.cachedSiteId!,
          listObj.id,
          token,
          folderEmail,
          {
            maxPages: lookupMaxPages,
            folderPin,
            onPage: (pageItems) => {
              if (isStale()) return;
              publishListHits(listName, mapRaw(pageItems));
            },
          },
        )
        : Promise.resolve(null);

      // Skipped when the deep crawl will run: it reads these same newest pages for every list.
      const supplementPromise =
        deferHeavyFanOut || supplementPages <= 0 || deepCrawlWillRun
          ? Promise.resolve([] as any[])
          : this.taskQuery.fetchSharePointListPages(
            this.cachedSiteId!,
            listObj.id,
            token,
            AppConstants.hrTasksFastLoadPageSize,
            supplementPages,
            true,
          );

      const [lookupItems, assigneeItems, titleItems, supplement] = await Promise.all([
        lookupPromise,
        assigneePromise,
        titlePromise,
        supplementPromise,
      ]);
      if (isStale()) return;

      const mergeRaw = (base: any[] | null, extra: any[] | null): any[] | null => {
        if (!extra?.length) return base;
        if (!base?.length) return extra;
        const byId = new Map<string, any>();
        for (const item of base) byId.set(String(item.id), item);
        for (const item of extra) byId.set(String(item.id), item);
        return [...byId.values()];
      };

      rawItems = mergeRaw(mergeRaw(lookupItems, assigneeItems), titleItems);

      if (supplement.length > 0) {
        publishListHits(listName, mapRaw(supplement));
        rawItems = mergeRaw(rawItems, supplement);
      }

      if (rawItems == null || rawItems.length === 0) {
        // LookupId unavailable or empty — capped newest-first scan + name match.
        const scanPages = isSourceEForm
          ? AppConstants.hrFilesSourceEFormScanPages
          : maxPages;
        const deepSupplementPages = isSourceEForm
          ? AppConstants.hrFilesSourceEFormLookupSupplementPages
          : AppConstants.hrFilesLookupSupplementPages;
        // The deep crawl reads the same newest pages for this list, so scanning them here
        // too would download them twice. Only scan now when it reaches further than that.
        const deepCrawlCoversScan = deepCrawlWillRun && scanPages <= deepSupplementPages;
        if (scanPages <= 0 || deepCrawlCoversScan) {
          rawItems = [];
        } else {
          rawItems = await this.taskQuery.fetchSharePointListPages(
            this.cachedSiteId!,
            listObj.id,
            token,
            AppConstants.hrTasksFastLoadPageSize,
            scanPages,
            true,
          );
          scannedListPages.set(listName, scanPages);
        }
      }

      const mapped = mapRaw(rawItems);
      if (mapped.length > 0) {
        publishListHits(listName, mapped);
      }
    };

    // Runs `run` for each item. Clicks and refreshes use a worker pool, so a slow list
    // never leaves other workers idle waiting for its batch. Background prefetch keeps
    // small paced batches to stay gentle on Graph.
    const forEachList = async <T>(items: T[], run: (item: T) => Promise<void>): Promise<void> => {
      const publish = (): void => {
        if (onPartial && mappedByList.size > 0) onPartial(flattenHrMappedByList(mappedByList));
      };
      if (crawlMode) {
        let next = 0;
        const workers = Array.from(
          { length: Math.min(AppConstants.hrFilesInteractiveListConcurrency, items.length) },
          async () => {
            while (!isStale()) {
              const index = next++;
              if (index >= items.length) break;
              await run(items[index]);
              publish();
            }
          },
        );
        await Promise.all(workers);
        return;
      }
      for (let i = 0; i < items.length; i += concurrency) {
        if (isStale()) break;
        await Promise.all(items.slice(i, i + concurrency).map(run));
        publish();
        if (i + concurrency < items.length) {
          await sleep(AppConstants.hrTaskListBatchGapMs);
        }
      }
    };

    if (options?.interactive) {
      // Click path: worker pool with no inter-batch sleep so a slow list never
      // stalls the rest and first paint is not delayed by idle gaps.
      let next = 0;
      const workers = Array.from(
        { length: Math.min(AppConstants.hrFilesInteractiveListConcurrency, listsToQuery.length) },
        async () => {
          while (!isStale()) {
            const index = next++;
            if (index >= listsToQuery.length) break;
            await processList(listsToQuery[index]);
            if (onPartial && mappedByList.size > 0) {
              onPartial(flattenHrMappedByList(mappedByList));
            }
          }
        },
      );
      await Promise.all(workers);
    } else {
      for (let i = 0; i < listsToQuery.length; i += concurrency) {
        if (isStale()) break;
        const batch = listsToQuery.slice(i, i + concurrency);
        await Promise.all(batch.map(listName => processList(listName)));

        if (onPartial && mappedByList.size > 0) {
          onPartial(flattenHrMappedByList(mappedByList));
        }
        // Pacing gaps are for background prefetch only; the shared Graph limiter already
        // caps what a click or refresh has in flight.
        if (!crawlMode && i + concurrency < listsToQuery.length) {
          await sleep(AppConstants.hrTaskListBatchGapMs);
        }
      }
    }

    // Interactive seed: newest LookupId/Title pages only — paints recent tasks first.
    const seedItems = flattenHrMappedByList(mappedByList);
    options?.onSeedComplete?.(seedItems);
    logCrawlPhase('first pass', ` (${listsToQuery.length} lists)`);

    // Deep crawl (cold and soft refresh): full assignee rows and the newest unfiltered pages
    // for every list, plus a LookupId re-read only where the first pass left rows unread.
    if (personLookupId && !isStale()) {
      const listsForTopUp = [...listObjByName.entries()];
      const deepLookupPages = AppConstants.hrFilesPersonLookupMaxPages;
      await forEachList(listsForTopUp, async ([listName, listObj]) => {
        {
          if (!listObj?.id) return;
          const mapRaw = (raw: any[]) =>
            this.mapHrFolderRawItems(
              raw,
              listName,
              listObj,
              folderName,
              personLookupId,
              folderMatchHints,
              userEmail,
              userUpn,
            );
          const isSourceEForm = isHrSourceEFormList(listName);
          const topUpSupplementPages = isSourceEForm
            ? AppConstants.hrFilesSourceEFormLookupSupplementPages
            : AppConstants.hrFilesLookupSupplementPages;

          // Pages the first pass already scanned for this list are not downloaded again.
          const supplementPagesNeeded =
            topUpSupplementPages > (scannedListPages.get(listName) ?? 0) ? topUpSupplementPages : 0;

          const [deepLookupItems, assigneeItems, supplement] = await Promise.all([
            deepLookupPages > AppConstants.hrFilesFirstPaintLookupPages &&
              lookupIncompleteLists.has(listName)
              ? this.taskQuery.fetchSharePointListItemsForPersonLookup(
                this.cachedSiteId!,
                listObj.id,
                token,
                personLookupId,
                {
                  maxPages: deepLookupPages,
                  onPage: (pageItems) => {
                    if (isStale()) return;
                    publishListHits(listName, mapRaw(pageItems));
                  },
                },
              )
              : Promise.resolve(null),
            this.taskQuery.fetchSharePointListItemsForAssigneeLookup(
              this.cachedSiteId!,
              listObj.id,
              token,
              personLookupId,
              {
                maxPages: deepLookupPages,
                onPage: (pageItems) => {
                  if (isStale()) return;
                  publishListHits(listName, mapRaw(pageItems));
                },
              },
            ).catch((err) => {
              console.warn(`[HR Files] assignee lookup failed for ${listName}`, err);
              return null;
            }),
            supplementPagesNeeded <= 0
              ? Promise.resolve([] as any[])
              : this.taskQuery.fetchSharePointListPages(
                this.cachedSiteId!,
                listObj.id,
                token,
                AppConstants.hrTasksFastLoadPageSize,
                supplementPagesNeeded,
                true,
              ),
          ]);

          if (isStale()) return;
          if (deepLookupItems?.length) {
            publishListHits(listName, mapRaw(deepLookupItems));
          }
          if (assigneeItems?.length) {
            publishListHits(listName, mapRaw(assigneeItems));
          }
          if (supplement.length) {
            publishListHits(listName, mapRaw(supplement));
          }
        }
      });
    }

    logCrawlPhase('deep crawl', ` (${lookupIncompleteLists.size} lists re-read)`);

    // Second pass: workflow siblings that share the same eForm ID AND the same task name.
    const seedTasksForSiblings = this.preferCurrentUserSeedTasks(
      flattenHrMappedByList(mappedByList),
      user,
    );
    // Every eForm id a person ever had costs one filter chunk per list. Workflow steps matter
    // for recent requests, so look up siblings for the newest ones only.
    const newestFirstSeeds = [...seedTasksForSiblings].sort(
      (a: any, b: any) =>
        new Date(b?.submittedDate || 0).getTime() - new Date(a?.submittedDate || 0).getTime(),
    );
    const eFormKeyToTaskNames = new Map(
      [...collectEFormKeyToTaskNames(newestFirstSeeds)].slice(0, AppConstants.hrFilesSiblingLookupMaxEForms),
    );
    if (eFormKeyToTaskNames.size > 0 && !isStale()) {
      const listsWithHits = [...mappedByList.keys()];
      const backfillPages = AppConstants.hrFilesTaskListBackfillPageLimit;
      await forEachList(listsWithHits, async listName => {
        {
          const listObj = listObjByName.get(listName);
          if (!listObj?.id) return;

          const eFormKeys = new Set(eFormKeyToTaskNames.keys());
          const relatedRaw = await this.taskQuery.fetchSharePointListItemsForEFormKeys(
            this.cachedSiteId!,
            listObj.id,
            token,
            eFormKeys,
          );
          // Optional fallback scan - off by default for speed (see hrFilesTaskListBackfillPageLimit).
          const scannedRaw = backfillPages > 0
            ? await this.taskQuery.fetchSharePointListPages(
              this.cachedSiteId!,
              listObj.id,
              token,
              AppConstants.hrTasksFastLoadPageSize,
              backfillPages,
              false,
            )
            : [];
          const relatedCombined = [...relatedRaw, ...scannedRaw].filter(item =>
            doesRawItemMatchSeedEFormIdAndTaskName(item, eFormKeyToTaskNames)
          );
          if (relatedCombined.length === 0) return;

          const relatedMapped = relatedCombined
            .filter((item: any) => !isSharePointCommentEntry(item))
            .map((item: any) => this.hrTaskMapper.mapSharePointItemToHrTask(item, listName, listObj, this.cachedSiteWebUrl, userEmail, userUpn, true))
            .filter((item): item is NonNullable<typeof item> => item !== null)
            .filter((item: any) => !isCommentTypeItem(item))
            .filter((item: any) =>
              doesMappedItemMatchSeedEFormIdAndTaskName(item, eFormKeyToTaskNames)
            );

          if (relatedMapped.length === 0) return;
          const existing = mappedByList.get(listName) ?? [];
          mappedByList.set(listName, mergeHrTasks(existing, relatedMapped));
        }
      });
    }

    logCrawlPhase(
      isStale() ? 'cancelled' : 'finished',
      ` (${eFormKeyToTaskNames.size} ids x ${mappedByList.size} lists for related steps)`,
    );
    return flattenHrMappedByList(mappedByList);
  }

  /**
   Prefer the logged-in user's own submissions as Pass-2 seeds.
   If none are present (e.g. browsing another person's folder), fall back to all seed tasks.
   */
  preferCurrentUserSeedTasks(seedTasks: any[], user: AppUser | null): any[] {
    const own = seedTasks.filter(task => this.isMappedTaskSubmittedByCurrentUser(task, user));
    return own.length > 0 ? own : seedTasks;
  }

  isMappedTaskSubmittedByCurrentUser(item: any, user: AppUser | null): boolean {
    const email = (user?.email ?? '').toLowerCase();
    const upn = (user?.userPrincipalName ?? '').toLowerCase();
    const display = (user?.username ?? '').toLowerCase();
    const tokens = [email, upn, display].map(t => normalizeTaskMatchText(t)).filter(t => t.length >= 3);
    if (tokens.length === 0) return false;

    const submitterCandidates = [
      item?.submittedBy,
      item?.eFormDetails?.submitter,
      item?.eFormDetails?.submittedBy,
      item?.eFormDetails?.commentSubmittedBy,
    ];
    return submitterCandidates.some(value => {
      const normalized = normalizeTaskMatchText(stringifyTaskFieldValue(value));
      return !!normalized && tokens.some(token => normalized.includes(token) || token.includes(normalized));
    });
  }
}
