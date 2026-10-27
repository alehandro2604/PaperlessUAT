import { Injectable, Injector, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { graphGetWithRetry, toGraphPath, normalizeName } from '../microsoft-graph';
import { FileCrawlCacheService } from '../services/file-crawl.service';
import { SiteMetadataService } from '../services/site-metadata.service';
import { HrTaskMapperService } from '../services/hr-task-mapper.service';
import { HrTaskFolderMatchService } from '../services/hr-task-folder-match.service';
import { SharePointTaskQueryService } from '../services/sharepoint-task-query.service';
import { TaskLoaderService } from '../services/task-loader.service';
import { sortCommentItemsByDateDesc, applyDueDatePatch } from '../utils/comment-utils';
import { extractTrailingFolderId, collectEFormKeysFromMappedTasks, prioritizeHrPersonTaskLists, mergeHrTasks } from '../utils/hr-task-matching';
import { hrFolderTaskCacheKey, libFolderTaskCacheKey } from '../utils/cache-keys';
import { ViewRefreshService } from './view-refresh.service';
import { UiStore } from './ui.store';
import { SessionStore } from './session.store';
import { CommentsStore } from './comments.store';
import { TodoStore } from './todo.store';
import { AttachmentsStore } from './attachments.store';
import { HrFilesStore } from './hr-files.store';
import { AllFilesStore } from './all-files.store';
import { getTaskListName, isSameTask } from '../utils/task-identity';

/**
 * Loading the tasks that belong to a folder (All Files library folder or HR Files person), with shared Redis caching, prefetch and cache patching after edits.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class FolderTasksStore {
  private readonly injector = inject(Injector);
  private readonly http = inject(HttpClient);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly siteMetadataService = inject(SiteMetadataService);
  private readonly hrTaskMapper = inject(HrTaskMapperService);
  private readonly folderMatch = inject(HrTaskFolderMatchService);
  private readonly taskQuery = inject(SharePointTaskQueryService);
  private readonly taskLoader = inject(TaskLoaderService);

  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get comments(): CommentsStore { return this.injector.get(CommentsStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get attachments(): AttachmentsStore { return this.injector.get(AttachmentsStore); }
  private get hr(): HrFilesStore { return this.injector.get(HrFilesStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  invalidateFolderTaskCache(libraryName: string, folderName: string): void {
    const folder = String(folderName || '').trim();
    if (!folder) return;
    if (normalizeName(libraryName) === normalizeName(this.hr.targetLibraryName)) {
      const key = hrFolderTaskCacheKey(folder);
      this.fileCrawlCache.invalidate(key);
      this.hrFolderCompleteAt.delete(key);
      this.hrFolderPartialKeys.delete(key);
      return;
    }
    const mappedTaskList = this.allFiles.getTaskListForLibrary(libraryName);
    if (mappedTaskList) {
      this.fileCrawlCache.invalidate(libFolderTaskCacheKey(mappedTaskList, folder));
    }
  }

  applyCachedFolderTaskItems(cached: any[], folderName: string, pendingRefresh = false): void {
    if (pendingRefresh) {
      this.comments.isLoadingMoreComments = true;
      this.comments.commentsMessage = '';
    } else {
      this.comments.isLoadingMoreComments = false;
    }
    this.publishFolderTaskProgress(
      sortCommentItemsByDateDesc(cached),
      {
        done: !pendingRefresh,
        emptyMessage: 'No tasks found.',
      },
    );
  }

  /** Count of in-flight HR / All Files folder Comments Graph crawls. */
  folderTaskGraphInFlight = 0;

  /** In-flight silent HR person task prefetches keyed by `tasks:hr-folder:{name}`. */
  readonly hrTaskPrefetchInFlight = new Set<string>();

  hrTaskPrefetchTimer: ReturnType<typeof setTimeout> | null = null;

  /** Debounce timer for hover/pointerdown single-person Comments warm. */
  hrPersonHoverPrefetchTimer: ReturnType<typeof setTimeout> | null = null;

  /** True while the one-at-a-time "warm every HR person" loop is running. */
  hrWarmAllRunning = false;

  /**
   * When each HR folder last finished a FULL crawl this session (not a partial seed or a
   * Redis snapshot). Reopening within `hrFilesSkipSoftRefreshMs` skips the Graph refresh.
   */
  readonly hrFolderCompleteAt = new Map<string, number>();

  /**
   * HR folder keys whose in-memory snapshot came from a crawl that has not finished
   * (first paints, prefetch seeds, cancelled loads). These stay out of shared Redis
   * and are resumed, never presented as the final list.
   */
  readonly hrFolderPartialKeys = new Set<string>();

  /** Prevent stale All Files task loads when folders are clicked rapidly. */
  allFilesFolderTaskLoadSeq = 0;

  /**
   * Publish folder/HR task results: always date-sorted (latest first).
   * Prefer painting when a load is complete so newer rows do not keep jumping to the top.
   * "Load more" pages the already-loaded set.
   * `background: true` = silent merge after the interactive spinner is already gone.
   */
  publishFolderTaskProgress(
    items: any[],
    options: {
      done?: boolean;
      emptyMessage?: string;
      background?: boolean;
      /** While true, show "Loading more tasks…" during silent background Graph top-up. */
      keepLoadingMore?: boolean;
    } = {},
  ): void {
    const sorted = sortCommentItemsByDateDesc(
      this.filterTasksForSelectedAllFilesFolder(items)
    );
    this.comments.commentItems = sorted;
    void this.allFiles.publishSelectedFolderChildSubjects(sorted, !!options.done && !options.background);

    // Silent top-up after first paint — keep Load more position.
    // `keepLoadingMore` stays true while deeper Graph pages are still coming in.
    if (options.background) {
      this.comments.isLoadingComments = false;
      this.comments.isLoadingMoreComments = !!options.keepLoadingMore;
      if (sorted.length > 0) {
        this.comments.commentsMessage = '';
      }
      this.comments.invalidateCommentFilters(false);
      this.view.refreshView();
      this.comments.focusPendingCommentIfNeeded();
      return;
    }

    if (options.done) {
      const wasLoading = this.comments.isLoadingComments;
      this.comments.isLoadingComments = false;
      this.comments.isLoadingMoreComments = false;
      this.comments.commentsMessage = sorted.length === 0 ? (options.emptyMessage || 'No tasks found.') : '';
      // Reset page window only on the first completed paint - keep Load more position
      // when a background related-steps refresh replaces the list.
      if (wasLoading || this.comments.visibleItemCount < this.comments.commentsInitialPageSize) {
        this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
        this.comments.invalidateCommentFilters(true);
      } else {
        this.comments.invalidateCommentFilters(false);
      }
      this.view.refreshView();
      this.comments.focusPendingCommentIfNeeded();
      return;
    }

    // Incomplete batches: paint first hits, then keep "Loading more tasks..." until
    // done:true — otherwise assignee/supplement top-up looks stuck mid-count.
    if (this.comments.isLoadingComments) {
      if (sorted.length === 0) {
        this.comments.commentsMessage = 'Loading tasks...';
        this.view.refreshView();
        return;
      }
      this.comments.isLoadingComments = false;
      this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
      this.comments.invalidateCommentFilters(true);
      this.comments.commentsMessage = '';
    } else {
      this.comments.invalidateCommentFilters(false);
      if (sorted.length > 0) {
        this.comments.commentsMessage = '';
      }
    }
    this.comments.isLoadingMoreComments = true;
    this.view.refreshView();
    this.comments.focusPendingCommentIfNeeded();
  }

  hrOlderHistoryLoadInProgress = false;

  selectedHrPersonalTaskFolder = '';

  /** Survives tab switches so HR Files Comments can restore from cache instantly. */
  lastHrFilesCommentsFolder = '';

  /** Rehydrate Comments for the last opened HR Files person without a Graph crawl. */
  restoreHrFilesCommentsFromCache(): void {
    const folder = String(this.lastHrFilesCommentsFolder || '').trim();
    if (!folder) return;
    const cached =
      this.fileCrawlCache.get(hrFolderTaskCacheKey(folder)) ??
      this.fileCrawlCache.getStale(hrFolderTaskCacheKey(folder));
    if (!Array.isArray(cached)) return;

    // Leaving the view cancels any crawl in flight. Unless a crawl for this person finished
    // recently, start it again instead of presenting the cached list as final.
    const cacheKey = hrFolderTaskCacheKey(folder);
    const completeAt = this.hrFolderCompleteAt.get(cacheKey);
    const finishedRecently =
      completeAt !== undefined && Date.now() - completeAt <= AppConstants.hrFilesSkipSoftRefreshMs;
    // A snapshot under an hour old from a finished crawl is final too (it may have come from
    // Redis or an earlier visit); only unfinished or older ones are crawled again.
    const freshAndComplete =
      this.fileCrawlCache.get(cacheKey) != null && !this.hrFolderPartialKeys.has(cacheKey);
    if (this.hrFolderPartialKeys.has(cacheKey) || (!finishedRecently && !freshAndComplete)) {
      this.ui.hideComments = false;
      this.loadAllFilesTasksForLibrary(this.hr.targetLibraryName, folder);
      return;
    }

    this.allFiles.selectedFolderName = folder;
    this.selectedHrPersonalTaskFolder = folder;
    this.ui.hideComments = false;
    this.publishFolderTaskProgress(cached, {
      done: true,
      emptyMessage: 'No tasks found.',
    });
  }

  /** Show the Comments panel and load tasks for the clicked folder. */
  loadAllFilesTasksForLibrary(libraryName: string, folderLabel: string): void {
    this.cancelPendingFolderTaskLoads();
    const loadSeq = ++this.allFilesFolderTaskLoadSeq;
    const loadFolder = String(folderLabel || '').trim();

    // Switching folders: drop the previous folder's rows so a slow cold load shows a
    // spinner instead of the old folder's tasks under the new folder's badge.
    // Same folder (search-hit seed, refresh) keeps what is already painted.
    if (String(this.allFiles.selectedFolderName ?? '').trim() !== loadFolder) {
      this.comments.commentItems = [];
      this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
      this.comments.invalidateCommentFilters(true);
    }

    this.todo.selectedSubmitter = null;
    this.todo.selectedEFormListId = null;
    this.todo.selectedEFormTitle = null;
    this.ui.hideComments = false;
    this.allFiles.selectedFolderName = loadFolder;
    if (normalizeName(libraryName) === normalizeName(this.hr.targetLibraryName)) {
      this.selectedHrPersonalTaskFolder = loadFolder;
    } else {
      this.selectedHrPersonalTaskFolder = '';
    }
    // Do not force Comments tab here - folder clicks open Attachments first so
    // files appear immediately while tasks load in the background.

    // Serve from cache when the same folder was crawled recently.
    if (normalizeName(libraryName) === normalizeName(this.hr.targetLibraryName)) {
      this.lastHrFilesCommentsFolder = loadFolder;
      const cacheKey = hrFolderTaskCacheKey(loadFolder);
      const fresh = this.fileCrawlCache.get(cacheKey);
      const cached = fresh ?? this.fileCrawlCache.getStale(cacheKey);
      // Paint cache instantly, then soft-refresh from Graph so new tasks appear —
      // unless a full crawl for this person just finished (rapid re-clicks).
      if (Array.isArray(cached) && cached.length > 0) {
        const completeAt = this.hrFolderCompleteAt.get(cacheKey);
        // A snapshot under an hour old that came from a finished crawl (only those reach
        // Redis) needs no Graph visit; the open-folder poll and the refresh button top it up.
        const freshAndComplete = fresh != null && !this.hrFolderPartialKeys.has(cacheKey);
        const needsRefresh =
          !freshAndComplete && (
            completeAt === undefined ||
            Date.now() - completeAt > AppConstants.hrFilesSkipSoftRefreshMs ||
            this.hrFolderPartialKeys.has(cacheKey)
          );
        this.applyCachedFolderTaskItems(cached, loadFolder, needsRefresh);
        if (needsRefresh) {
          void this.loadHrPersonalTasksForFolder(loadFolder, loadSeq, { softRefresh: true });
        }
        return;
      }
      if (Array.isArray(cached) && cached.length === 0) {
        this.applyCachedFolderTaskItems(cached, loadFolder, true);
        void this.loadHrPersonalTasksForFolder(loadFolder, loadSeq, { softRefresh: true });
        return;
      }

      // Memory miss — shared Redis first (no spinner yet), then Graph.
      void this.loadHrFolderTasksWithSharedCache(loadFolder, loadSeq);
      return;
    } else {
      const mappedTaskList = this.allFiles.getTaskListForLibrary(libraryName);
      if (mappedTaskList) {
        const cacheKey = libFolderTaskCacheKey(mappedTaskList, loadFolder);
        const fresh = this.fileCrawlCache.get(cacheKey);
        const cached = fresh ?? this.fileCrawlCache.getStale(cacheKey);
        if (Array.isArray(cached)) {
          this.applyCachedFolderTaskItems(cached, loadFolder);
          return;
        }

        // Child folders: reuse parent task cache when Title still points at the parent.
        const fromParent = this.tryReuseParentLibFolderTaskCache(mappedTaskList, loadFolder);
        if (fromParent) {
          this.applyCachedFolderTaskItems(fromParent, loadFolder);
          return;
        }

        // Memory miss — shared Redis first (no spinner yet), then Graph.
        void this.loadLibFolderTasksWithSharedCache(mappedTaskList, loadFolder, loadSeq);
        return;
      }
    }

    this.beginFolderTaskLoadingUi();

    const mappedTaskList = this.allFiles.getTaskListForLibrary(libraryName);
    if (mappedTaskList) {
      void this.loadDocumentLibraryTasksForFolder(mappedTaskList, loadFolder, loadSeq);
      return;
    }

    if (loadSeq !== this.allFilesFolderTaskLoadSeq) return;
    this.comments.isLoadingComments = false;
    this.comments.commentsMessage = `No task list configured for ${libraryName}.`;
    if (this.allFiles.selectedAllFilesFolderId) {
      this.ui.allFilesComponent?.setChildSubjectFolderNames(this.allFiles.selectedAllFilesFolderId, []);
    }
    this.view.refreshView();
  }

  beginFolderTaskLoadingUi(): void {
    // Keep already-painted Comments visible (seeded search hit or prior cache) —
    // only full-screen spinner when the panel is empty.
    if (this.comments.commentItems.length === 0) {
      this.comments.commentsMessage = 'Loading tasks...';
      this.comments.isLoadingComments = true;
      this.comments.isLoadingMoreComments = false;
    } else {
      this.comments.isLoadingComments = false;
      this.comments.isLoadingMoreComments = true;
    }
    this.view.refreshView();
  }

  /**
   * Doc-lib folder tasks: try shared Redis (written by any signed-in user) before Graph.
   * Spinner only after Redis miss so cached folders feel instant.
   */
  async loadLibFolderTasksWithSharedCache(
    listName: string,
    folderName: string,
    loadSeq: number,
  ): Promise<void> {
    const cacheKey = libFolderTaskCacheKey(listName, folderName);
    const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
    if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

    if (hydrated && Array.isArray(hydrated.data)) {
      this.applyCachedFolderTaskItems(hydrated.data, folderName);
      // Skip auto soft-refresh — Graph fan-out belongs on explicit refresh / cold miss.
      return;
    }

    this.beginFolderTaskLoadingUi();
    await this.loadDocumentLibraryTasksForFolder(listName, folderName, loadSeq);
  }

  /**
   * HR Files person-folder tasks: try shared Redis before Graph.
   * Spinner only after Redis miss so cached people feel instant (same as All Files).
   */
  async loadHrFolderTasksWithSharedCache(
    folderName: string,
    loadSeq: number,
  ): Promise<void> {
    const cacheKey = hrFolderTaskCacheKey(folderName);
    const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
    if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

    // Array hit from shared Redis. A snapshot under an hour old is shown as is (like
    // document-library folders); an older one is painted, then soft-refreshed.
    if (hydrated && Array.isArray(hydrated.data)) {
      this.applyCachedFolderTaskItems(hydrated.data, folderName, !hydrated.fresh);
      if (!hydrated.fresh) {
        void this.loadHrPersonalTasksForFolder(folderName, loadSeq, { softRefresh: true });
      }
      return;
    }

    this.beginFolderTaskLoadingUi();
    await this.loadHrPersonalTasksForFolder(folderName, loadSeq);
  }

  /**
   * When opening a ChildSubject folder, parent Comments may already be cached.
   * Filter those rows by the child's trailing eForm id instead of waiting on a full list scan.
   */
  tryReuseParentLibFolderTaskCache(listName: string, childFolderName: string): any[] | null {
    const parentName = String(this.allFiles.selectedAllFilesParentFolderName ?? '').trim();
    if (!parentName || normalizeName(parentName) === normalizeName(childFolderName)) {
      return null;
    }
    const childEFormId = extractTrailingFolderId(childFolderName);
    if (!childEFormId) return null;

    const parentCached = this.fileCrawlCache.getStale(libFolderTaskCacheKey(listName, parentName));
    if (!Array.isArray(parentCached) || parentCached.length === 0) return null;

    const matched = parentCached.filter((item) => this.folderMatch.doesMappedHrTaskMatchFolder(item, childFolderName));
    if (matched.length === 0) return null;

    const sorted = sortCommentItemsByDateDesc(matched);
    this.fileCrawlCache.set(libFolderTaskCacheKey(listName, childFolderName), sorted);
    return sorted;
  }

  isStaleAllFilesFolderTaskLoad(loadSeq: number): boolean {
    return loadSeq !== this.allFilesFolderTaskLoadSeq;
  }

  /**
   * Idle-prefetch Comments for visible HR Files people so the first click is often cached
   * (same idea as All Files folder content prefetch).
   * Capped batch — full-list prefetch flooded Graph and caused 429s.
   */
  scheduleHrFolderTasksPrefetch(
    folders: Array<{ name?: string }>,
  ): void {
    const names = folders
      .map(f => String(f?.name ?? '').trim())
      .filter(Boolean)
      .filter(name => !this.fileCrawlCache.get(hrFolderTaskCacheKey(name)))
      .filter(name => !this.hrTaskPrefetchInFlight.has(hrFolderTaskCacheKey(name)))
      .slice(0, AppConstants.hrFilesTaskPrefetchMaxBatch);
    if (names.length === 0) return;

    if (this.hrTaskPrefetchTimer) {
      clearTimeout(this.hrTaskPrefetchTimer);
    }
    this.hrTaskPrefetchTimer = setTimeout(() => {
      this.hrTaskPrefetchTimer = null;
      if (this.attachments.folderPrefetchPaused) return;
      if (this.comments.isLoadingComments || this.comments.isLoadingMoreComments) return;
      void this.prefetchHrFolderTasks(names);
    }, AppConstants.hrFilesTaskPrefetchDelayMs);
  }

  /** Hover / pointerdown: start warming this person's Comments before the click. */
  prefetchHrPersonComments(item: { name?: string } | null | undefined): void {
    const name = String(item?.name ?? '').trim();
    if (!name) return;
    const key = hrFolderTaskCacheKey(name);
    if (this.fileCrawlCache.get(key) || this.hrTaskPrefetchInFlight.has(key)) return;
    if (this.attachments.folderPrefetchPaused && this.todo.isLoadingTodoTasks && !this.todo.todoFirstPassComplete) {
      return;
    }
    // Debounce scroll/hover spam; pointerdown still lands within one frame of click.
    if (this.hrPersonHoverPrefetchTimer) {
      clearTimeout(this.hrPersonHoverPrefetchTimer);
    }
    this.hrPersonHoverPrefetchTimer = setTimeout(() => {
      this.hrPersonHoverPrefetchTimer = null;
      void this.prefetchHrFolderTasks([name]);
    }, 120);
  }

  /** Silently warm `tasks:hr-folder:*` without touching the Comments UI. */
  async prefetchHrFolderTasks(folderNames: string[]): Promise<void> {
    if (!AppConstants.hrFilesBackgroundPrefetchEnabled) return;
    if (this.attachments.folderPrefetchPaused || folderNames.length === 0) return;
    if (!this.session.backgroundGraphAllowed()) return;
    if (this.hr.isLoadingHrFilesList) return;
    // Allow during To Do background drain — only skip while the first To Do pass
    // still owns the spinner (pauseFolderPrefetch covers that window).
    if (this.todo.isLoadingTodoTasks && !this.todo.todoFirstPassComplete) return;

    const pending = folderNames.filter(name => {
      const key = hrFolderTaskCacheKey(name);
      if (this.fileCrawlCache.get(key)) return false;
      if (this.hrTaskPrefetchInFlight.has(key)) return false;
      return true;
    });
    if (pending.length === 0) return;

    let token: string;
    try {
      token = await this.session.getSharePointToken();
    } catch {
      return;
    }

    try {
      await this.session.ensureFormConfigLoaded();
      await this.session.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );
      if (!this.session.cachedSiteId) return;
    } catch {
      return;
    }

    const listsToQuery = this.taskLoader.getTaskListsToQuery(this.session.cachedSiteLists);
    if (listsToQuery.length === 0) return;

    const concurrency = AppConstants.hrFilesTaskPrefetchConcurrency;
    for (let i = 0; i < pending.length; i += concurrency) {
      if (this.attachments.folderPrefetchPaused) return;
      if (this.comments.isLoadingComments || this.comments.isLoadingMoreComments) return;
      if (!this.session.backgroundGraphAllowed()) return;

      const batch = pending.slice(i, i + concurrency);
      await Promise.all(
        batch.map(async folderName => {
          const key = hrFolderTaskCacheKey(folderName);
          if (this.fileCrawlCache.get(key) || this.hrTaskPrefetchInFlight.has(key)) return;
          // HR snapshots are not bulk-loaded at sign-in — skip the Graph crawl when
          // another user's recent crawl is already in shared Redis.
          const shared = await this.fileCrawlCache.hydrateFromPersistent<any[]>(key);
          if (shared?.fresh || this.hrTaskPrefetchInFlight.has(key)) return;

          // Snapshot active load seq - abort if the user starts a real folder task load.
          const loadSeq = this.allFilesFolderTaskLoadSeq;
          this.hrTaskPrefetchInFlight.add(key);
          let seedSnapshot: any[] | null = null;
          try {
            const mapped = await this.taskLoader.collectHrPersonalTasksFromLists(
              listsToQuery,
              folderName,
              token,
              this.session.currentUser,
              () => this.isStaleAllFilesFolderTaskLoad(loadSeq),
              AppConstants.hrFilesTaskListPageLimit,
              undefined,
              {
                onSeedComplete: (seedItems) => {
                  if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
                  if (this.attachments.folderPrefetchPaused) return;
                  // Cache seed early so a click during prefetch is already instant — memory
                  // only, and never over an existing snapshot (it may be a complete one).
                  if (this.fileCrawlCache.getStale(key) == null) {
                    seedSnapshot = sortCommentItemsByDateDesc(seedItems);
                    this.fileCrawlCache.setMemoryOnly(key, seedSnapshot);
                    this.hrFolderPartialKeys.add(key);
                  }
                },
              },
            );
            if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
            if (this.attachments.folderPrefetchPaused) return;
            // Skip write if a click already cached this folder while we were fetching —
            // but do replace our own partial seed with the full result.
            const current = this.fileCrawlCache.get(key);
            if (current && current !== seedSnapshot) return;
            // Never shrink an older snapshot: a silently failed list must not drop its rows.
            const prior = this.fileCrawlCache.getStale<any[]>(key);
            const full = Array.isArray(prior) && prior !== seedSnapshot
              ? mergeHrTasks(prior, mapped)
              : mapped;
            this.fileCrawlCache.set(key, sortCommentItemsByDateDesc(full));
            this.hrFolderPartialKeys.delete(key);
            this.hrFolderCompleteAt.set(key, Date.now());
          } catch {
            // Prefetch failures are silent - click path will retry.
          } finally {
            this.hrTaskPrefetchInFlight.delete(key);
          }
        }),
      );

      if (i + concurrency < pending.length) {
        await new Promise(resolve =>
          setTimeout(resolve, AppConstants.hrFilesTaskPrefetchBatchGapMs),
        );
      }
    }
  }

  /** Soft-refresh the open folder Comments cache from Graph without wiping the painted UI. */
  softRefreshOpenFolderTaskCache(options: { onlyIfStale?: boolean } = {}): void {
    const folder = String(this.allFiles.selectedFolderName || this.selectedHrPersonalTaskFolder || '').trim();
    if (!folder) return;

    // The idle poll leaves a finished snapshot under an hour old alone: re-crawling it only
    // re-sent the same requests and flashed "Loading more tasks…" on a list that was done.
    if (options.onlyIfStale) {
      const isHr =
        normalizeName(this.attachments.currentLibraryName || this.hr.targetLibraryName) ===
        normalizeName(this.hr.targetLibraryName);
      const mappedList = isHr ? null : this.allFiles.getTaskListForLibrary(this.attachments.currentLibraryName ?? '');
      const key = isHr
        ? hrFolderTaskCacheKey(folder)
        : mappedList ? libFolderTaskCacheKey(mappedList, folder) : '';
      if (key && this.fileCrawlCache.get(key) != null && !this.hrFolderPartialKeys.has(key)) return;
    }

    const loadSeq = ++this.allFilesFolderTaskLoadSeq;

    if (this.ui.showAllFilesSection && this.attachments.currentLibraryName) {
      if (normalizeName(this.attachments.currentLibraryName) === normalizeName(this.hr.targetLibraryName)) {
        void this.loadHrPersonalTasksForFolder(folder, loadSeq, { softRefresh: true });
        return;
      }
      const mappedTaskList = this.allFiles.getTaskListForLibrary(this.attachments.currentLibraryName);
      if (mappedTaskList) {
        void this.loadDocumentLibraryTasksForFolder(mappedTaskList, folder, loadSeq, {
          softRefresh: true,
        });
      }
      return;
    }

    if (this.isHrPersonalFilesContext()) {
      void this.loadHrPersonalTasksForFolder(folder, loadSeq, { softRefresh: true });
    }
  }

  isHrPersonalFilesContext(): boolean {
    return !!String(this.selectedHrPersonalTaskFolder ?? '').trim();
  }

  /** True when Comments is scoped to a folder (HR Files person or All Files folder). */
  isFolderTaskCommentsViewActive(): boolean {
    if (this.ui.showAllFilesSection && !!String(this.allFiles.selectedFolderName ?? '').trim()) {
      return true;
    }
    return this.isHrPersonalFilesContext();
  }

  /** Cancel in-flight folder/global task fetches so a new folder click wins. */
  cancelPendingFolderTaskLoads(): void {
    this.allFilesFolderTaskLoadSeq++;
    this.todo.hrCommentsLoadSeq++;
  }

  /** Remove a task from every in-memory folder-task snapshot (HR + All Files). */
  removeTaskFromFolderCaches(taskId: string, listName?: string): void {
    this.mapTaskInFolderCaches(taskId, () => null, listName);
  }

  /** Patch a task inside every in-memory folder-task snapshot that contains it. */
  patchTaskInFolderCaches(taskId: string, patch: (item: any) => any, listName?: string): void {
    this.mapTaskInFolderCaches(taskId, patch, listName);
  }

  /**
   * Walk HR / lib folder-task caches. Return `null` from `mapFn` to remove the item.
   * Pass `listName` so a task with the same id in another list is left alone.
   */
  mapTaskInFolderCaches(taskId: string, mapFn: (item: any) => any | null, listName?: string): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;

    for (const prefix of ['tasks:hr-folder:', 'tasks:lib:v7:'] as const) {
      for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix(prefix)) {
        if (!Array.isArray(data) || data.length === 0) continue;

        let changed = false;
        const next: any[] = [];
        for (const item of data) {
          if (!isSameTask(item, id, listName)) {
            next.push(item);
            continue;
          }
          changed = true;
          const mapped = mapFn(item);
          if (mapped != null) next.push(mapped);
        }
        if (changed) {
          // Unfinished HR snapshots stay memory-only (see hrFolderPartialKeys).
          if (this.hrFolderPartialKeys.has(key)) {
            this.fileCrawlCache.setMemoryOnly(key, next);
          } else {
            this.fileCrawlCache.set(key, next);
          }
        }
      }
    }
  }

  applyAssigneePatch(item: any, assignedToName: string, assignedToEmail = ''): any {
    const label = String(assignedToName || assignedToEmail || '').trim();
    const details = { ...(item?.eFormDetails ?? {}) };
    if (label) details.assignedTo = label;
    const next: any = {
      ...item,
      assignedTo: label || item?.assignedTo,
      eFormDetails: details,
    };
    if (typeof item?.isAssignedToCurrentUser === 'boolean') {
      next.isAssignedToCurrentUser = this.isAssigneeCurrentUser(assignedToName, assignedToEmail);
    }
    return next;
  }

  isAssigneeCurrentUser(assignedToName: string, assignedToEmail = ''): boolean {
    const email = (this.session.currentUser?.email ?? '').toLowerCase();
    const upn = (this.session.currentUser?.userPrincipalName ?? '').toLowerCase();
    const name = (this.session.currentUser?.username ?? '').toLowerCase();
    const candEmail = String(assignedToEmail ?? '').trim().toLowerCase();
    const candName = String(assignedToName ?? '').trim().toLowerCase();
    if (candEmail && (candEmail === email || candEmail === upn)) return true;
    if (candName && name && candName === name) return true;
    return false;
  }

  /** Best-effort folder key invalidation when list/folder can be inferred from the task. */
  invalidateRelatedFolderTaskCache(task: any): void {
    const folder =
      String(this.allFiles.selectedFolderName ?? '').trim() ||
      String(this.selectedHrPersonalTaskFolder ?? '').trim() ||
      String(task?.submittedBy ?? task?.eFormDetails?.submittedBy ?? '').trim();
    if (!folder) return;

    const listName = String(
      task?.listName ?? task?.eFormDetails?.listName ?? this.attachments.currentLibraryName ?? '',
    ).trim();
    if (listName) {
      this.invalidateFolderTaskCache(listName, folder);
    }
    // HR person folders are keyed under HRPersonal regardless of the task list name.
    this.invalidateFolderTaskCache(this.hr.targetLibraryName, folder);
  }

  /**
   * Keep Redis (and in-memory) task snapshots aligned after an app ? SharePoint write.
   * Also drops the backend `/api/tasks` read-through key when list + email are known.
   */
  syncTaskCachesAfterMutation(
    task: any,
    kind: 'complete' | 'dueDate' | 'delegate' | 'comment',
    extras?: { dueDate?: string; assignedToName?: string; assignedToEmail?: string },
  ): void {
    const taskId = String(task?.id ?? task?.Id ?? '').trim();
    if (!taskId) return;
    // Ids repeat across lists; the list name keeps other lists' tasks untouched.
    const listName = getTaskListName(task);

    if (kind === 'complete') {
      this.todo.removeTaskFromTodoCache(taskId, listName);
      this.removeTaskFromFolderCaches(taskId, listName);
      this.invalidateRelatedFolderTaskCache(task);
    } else if (kind === 'dueDate' && extras?.dueDate) {
      const dueDate = extras.dueDate;
      this.todo.patchTaskInTodoCache(taskId, (item) => applyDueDatePatch(item, dueDate), listName);
      this.patchTaskInFolderCaches(taskId, (item) => applyDueDatePatch(item, dueDate), listName);
    } else if (kind === 'delegate') {
      const name = String(extras?.assignedToName ?? '').trim();
      const email = String(extras?.assignedToEmail ?? '').trim();
      if (!this.isAssigneeCurrentUser(name, email)) {
        this.todo.removeTaskFromTodoCache(taskId, listName);
      } else {
        this.todo.patchTaskInTodoCache(taskId, (item) => this.applyAssigneePatch(item, name, email), listName);
      }
      this.patchTaskInFolderCaches(taskId, (item) => this.applyAssigneePatch(item, name, email), listName);
      this.invalidateRelatedFolderTaskCache(task);
    } else if (kind === 'comment') {
      // Action comments may change assignee (handled via delegate); still refresh folder snapshots.
      this.invalidateRelatedFolderTaskCache(task);
    }

    this.session.invalidateBackendTasksCache(task);
  }

  filterTasksForSelectedAllFilesFolder(items: any[]): any[] {
    if (!this.ui.showAllFilesSection) return items;
    const folderName = String(this.allFiles.selectedFolderName ?? '').trim();
    if (!folderName || !extractTrailingFolderId(folderName)) return items;
    return items.filter(item => this.folderMatch.doesMappedHrTaskMatchFolder(item, folderName));
  }

  /** Loads HR tasks from SharePoint lists that belong to the person represented by folderName. */
  async loadHrPersonalTasksForFolder(
    folderName: string,
    loadSeq: number,
    options: { softRefresh?: boolean } = {},
  ): Promise<void> {
    const softRefresh = !!options.softRefresh;
    const folderCacheKey = hrFolderTaskCacheKey(folderName);
    const mergeWithCachedFolderTasks = (incoming: any[]): any[] => {
      if (!softRefresh) return incoming;
      const cached = this.fileCrawlCache.getStale<any[]>(folderCacheKey);
      if (!Array.isArray(cached) || cached.length === 0) return incoming;
      return mergeHrTasks(cached, incoming);
    };
    try {
      this.folderTaskGraphInFlight += 1;
      // Do not wait for To Do — people with 800+ tasks were blocked for minutes
      // before Comments even started. Folder click already cancelled To Do Graph.
      const tokenPromise = this.session.getSharePointToken();
      const formPromise = this.session.ensureFormConfigLoaded();
      const token = await tokenPromise;
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
      await formPromise;
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

      await this.session.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );

      const listsToQuery = prioritizeHrPersonTaskLists(
        this.taskLoader.getTaskListsToQuery(this.session.cachedSiteLists).filter(listName =>
          this.session.cachedSiteLists.some((l: any) =>
            (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
            (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
          )
        ),
      );
      if (!softRefresh) {
        this.comments.commentsMessage = 'Loading tasks...';
        this.view.refreshView();
        // Some people have no hits on the first lists — don't leave the blank
        // full-screen spinner up for the whole fan-out. Switch to inline status.
        window.setTimeout(() => {
          if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
          if (!this.comments.isLoadingComments || this.comments.commentItems.length > 0) return;
          this.comments.isLoadingComments = false;
          this.comments.isLoadingMoreComments = true;
          this.comments.commentsMessage = 'No tasks found.';
          this.view.refreshView();
        }, 900);
      }

      // Newest-first progressive paint: clear full-screen spinner on first hits, then
      // keep "Loading more tasks…" until the deep background crawl finishes.
      let interactivePaintDone = softRefresh;
      const publishMapped = (
        mapped: any[],
        opts: { background?: boolean; complete?: boolean } = {},
      ) => {
        if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
        if (softRefresh && mapped.length === 0 && !opts.complete) return;
        const items = mergeWithCachedFolderTasks(mapped);
        const sorted = sortCommentItemsByDateDesc(items);

        if (!interactivePaintDone && sorted.length === 0) {
          this.publishFolderTaskProgress(sorted, {
            done: false,
            emptyMessage: 'No tasks found.',
          });
          return;
        }

        //this only finds the tasks that are in the folder,this is the first pass
        // Only a finished crawl goes to shared Redis. In-progress paints stay in memory,
        // so a cancelled load can never leave a partial list behind for everyone.
        if (opts.complete) {
          this.fileCrawlCache.set(folderCacheKey, sorted);
          this.hrFolderPartialKeys.delete(folderCacheKey);
        } else if (sorted.length > 0 || softRefresh) {
          this.fileCrawlCache.setMemoryOnly(folderCacheKey, sorted);
          if (!softRefresh) this.hrFolderPartialKeys.add(folderCacheKey);
        }

        if (!opts.background && !interactivePaintDone && sorted.length > 0) {
          interactivePaintDone = true;
          this.lastHrFilesCommentsFolder = folderName;
          this.publishFolderTaskProgress(sorted, {
            done: true,
            keepLoadingMore: !opts.complete,
            emptyMessage: 'No tasks found.',
          });
          // First paint uses done:true which clears loading-more — put it back until deep crawl ends.
          if (!opts.complete) {
            this.comments.isLoadingMoreComments = true;
            this.view.refreshView();
          }
          return;
        }

        this.publishFolderTaskProgress(sorted, {
          done: true,
          background: interactivePaintDone || !!opts.background,
          keepLoadingMore: interactivePaintDone && !opts.complete,
          emptyMessage: 'No tasks found.',
        });
      };

      const fastMapped = await this.taskLoader.collectHrPersonalTasksFromLists(
        listsToQuery,
        folderName,
        token,
        this.session.currentUser,
        () => this.isStaleAllFilesFolderTaskLoad(loadSeq),
        AppConstants.hrFilesTaskListPageLimit,
        (partial) => publishMapped(partial, { background: interactivePaintDone }),
        {
          softRefresh,
          interactive: !softRefresh,
          onSeedComplete: (seedItems) => {
            if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
            if (softRefresh && seedItems.length === 0) return;
            publishMapped(seedItems, { background: interactivePaintDone });
            this.lastHrFilesCommentsFolder = folderName;
          },
        },
      );
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

      if (softRefresh && fastMapped.length === 0) {
        // Still finish UI loading state even when Graph returned nothing new.
        this.hrFolderCompleteAt.set(folderCacheKey, Date.now());
        this.comments.isLoadingMoreComments = false;
        this.view.refreshView();
        return;
      }

      publishMapped(fastMapped, { background: true, complete: true });
      this.hrFolderCompleteAt.set(folderCacheKey, Date.now());
      this.lastHrFilesCommentsFolder = folderName;
    } catch (err: any) {
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
      console.error(`[HR Files] Failed loading tasks for "${folderName}"`, err);
      if (softRefresh) return;
      this.comments.isLoadingMoreComments = false;
      this.comments.isLoadingComments = false;
      // Keep whatever is already painted; only show a message when the panel is empty.
      if (this.comments.commentItems.length === 0) {
        this.comments.invalidateCommentFilters();
        this.comments.commentsMessage = 'Could not load tasks. Press refresh to try again.';
      }
      this.view.refreshView();
    } finally {
      this.folderTaskGraphInFlight = Math.max(0, this.folderTaskGraphInFlight - 1);
      if (!this.isStaleAllFilesFolderTaskLoad(loadSeq) && (this.comments.isLoadingComments || this.comments.isLoadingMoreComments)) {
        this.comments.isLoadingComments = false;
        this.comments.isLoadingMoreComments = false;
        this.view.refreshView();
      }
    }
  }

  /** OData-safe Title filter for document-library folder tasks (All Files only). */
  buildDocumentLibraryFolderTitleFilter(folderName: string): string | null {
    const title = String(folderName ?? '').trim();
    if (!title) return null;
    const escaped = title.replace(/'/g, "''");
    return `fields/Title eq '${escaped}'`;
  }

  /**
   * Load tasks for the clicked folder from a document-library-associated task list.
   * Does not touch HRPersonal person-folder loading.
   *
   * Fast path: eFormListId / trailing folder id first (child folders), then Title eq
   * folder name. Full list scan only when both filters fail or find nothing.
   */
  async loadDocumentLibraryTasksForFolder(
    listName: string,
    folderName: string,
    loadSeq: number,
    options: { softRefresh?: boolean; silent?: boolean } = {},
  ): Promise<void> {
    const softRefresh = !!options.softRefresh;
    const silent = !!options.silent;
    const isStale = () => (silent ? false : this.isStaleAllFilesFolderTaskLoad(loadSeq));
    try {
      this.folderTaskGraphInFlight += 1;
      const token = await this.session.getSharePointToken();
      if (isStale()) return;

      await this.session.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );

      const siteId = this.session.cachedSiteId!;
      let listObj = this.session.cachedSiteLists.find((l: any) =>
        (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
        (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
      );

      // Not in the enumerated lists — ask Graph for it by name, which also tells us
      // whether this user lacks access (403) or the list really does not exist.
      let lookupStatus: 'ok' | 'forbidden' | 'not-found' | 'error' = 'ok';
      if (!listObj?.id) {
        const direct = await this.siteMetadataService.lookupListDirect(token, listName);
        if (isStale()) return;
        lookupStatus = direct.status;
        if (direct.list?.id) {
          // lookupListDirect already adds it to SiteMetadataService's list cache.
          listObj = direct.list;
        }
      }

      if (!listObj?.id) {
        if (isStale()) return;
        if (softRefresh || silent) return;
        console.warn(`[All Files] Task list "${listName}" unavailable for this user (${lookupStatus}).`);
        this.comments.commentsMessage =
          lookupStatus === 'forbidden'
            ? `You don't have access to the task list "${listName}". Ask a SharePoint admin to grant you read access.`
            : `Could not find task list "${listName}".`;
        this.comments.commentItems = [];
        this.comments.isLoadingComments = false;
        this.view.refreshView();
        return;
      }

      const userEmail = (this.session.currentUser?.email ?? '').toLowerCase();
      const userUpn = (this.session.currentUser?.userPrincipalName ?? '').toLowerCase();
      const collectedMapped: any[] = [];
      const seenIds = new Set<string>();
      const pageSize = AppConstants.docLibraryFolderTaskPageSize;
      const maxPages = AppConstants.docLibraryFolderTaskMaxPages;

      const publishAndCache = (done: boolean): void => {
        if (isStale()) return;
        if (softRefresh && collectedMapped.length === 0) return;
        const items = collectedMapped;
        this.fileCrawlCache.set(
          libFolderTaskCacheKey(listName, folderName),
          sortCommentItemsByDateDesc(items),
        );
        if (silent) {
          if (done) {
            this.ui.allFilesComponent?.refreshCommentSearchFromCaches();
          }
          return;
        }
        this.publishFolderTaskProgress(items, {
          done,
          emptyMessage: 'No tasks found.',
        });
      };

      const mergeMappedItems = (pageMapped: any[]): void => {
        for (const item of pageMapped) {
          const id = String(item?.id ?? '').trim();
          if (id) {
            if (seenIds.has(id)) continue;
            seenIds.add(id);
          }
          collectedMapped.push(item);
        }
      };

      const mapFolderPageItems = (
        rawItems: any[],
        sourceListName: string,
        sourceListObj: { id: string; name?: string; displayName?: string; webUrl?: string },
      ): any[] =>
        rawItems
          .filter((item: any) => this.folderMatch.doesSharePointTaskMatchFolder(item, folderName))
          .map((item: any) =>
            this.hrTaskMapper.mapSharePointItemToHrTask(
              item,
              sourceListName,
              sourceListObj,
              this.session.cachedSiteWebUrl,
              userEmail,
              userUpn,
              true,
            )
          )
          .filter((item: any): item is NonNullable<typeof item> => item !== null);

      const folderEFormId = extractTrailingFolderId(folderName);
      const titleFilter = this.buildDocumentLibraryFolderTitleFilter(folderName);

      /** Cheap eFormListId + Title filters against one SharePoint list. */
      const queryListWithFilters = async (
        sourceListName: string,
        sourceListObj: { id: string; name?: string; displayName?: string; webUrl?: string },
      ): Promise<void> => {
        if (folderEFormId) {
          try {
            const byIdRaw = await this.taskQuery.fetchSharePointListItemsForEFormKeys(
              siteId,
              sourceListObj.id,
              token,
              new Set([folderEFormId]),
            );
            if (isStale()) return;
            mergeMappedItems(mapFolderPageItems(byIdRaw, sourceListName, sourceListObj));
            if (collectedMapped.length > 0) publishAndCache(false);
          } catch (idFilterErr) {
            console.warn(
              `[All Files] eFormListId filter unavailable on "${sourceListName}"; trying Title.`,
              idFilterErr,
            );
          }
        }

        if (!titleFilter) return;
        try {
          const prefer = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
          let filteredPath: string | null =
            `/sites/${siteId}/lists/${sourceListObj.id}/items?$expand=fields&$top=${pageSize}` +
            `&$filter=${encodeURIComponent(titleFilter)}`;
          let filteredPages = 0;
          while (filteredPath && filteredPages < maxPages) {
            const page: any = await graphGetWithRetry(
              this.http,
              filteredPath,
              token,
              AppConstants.graphFileListingTimeoutMs,
              prefer,
            );
            filteredPages += 1;
            if (isStale()) return;

            mergeMappedItems(mapFolderPageItems(page?.value ?? [], sourceListName, sourceListObj));
            filteredPath = toGraphPath(page?.['@odata.nextLink']);
            if (collectedMapped.length > 0) publishAndCache(false);
            if (filteredPath) await new Promise(r => setTimeout(r, 50));
          }
        } catch (filterErr) {
          console.warn(
            `[All Files] Title filter unavailable on "${sourceListName}".`,
            filterErr,
          );
        }
      };

      const fullScanList = async (
        sourceListName: string,
        sourceListObj: { id: string; name?: string; displayName?: string; webUrl?: string },
      ): Promise<void> => {
        console.warn(`[All Files] Falling back to full scan of "${sourceListName}"`);
        let nextPath: string | null =
          `/sites/${siteId}/lists/${sourceListObj.id}/items?$expand=fields&$top=${pageSize}`;
        let pagesLoaded = 0;

        while (nextPath && pagesLoaded < maxPages) {
          const page: any = await graphGetWithRetry(
            this.http,
            nextPath,
            token,
            AppConstants.graphFileListingTimeoutMs,
          );
          pagesLoaded += 1;
          if (isStale()) return;

          if (!silent && this.comments.isLoadingComments) {
            this.comments.commentsMessage =
              pagesLoaded === 1
                ? `Scanning ${sourceListName}...`
                : `Scanning ${sourceListName} (page ${pagesLoaded})...`;
            this.view.refreshView();
          }

          const before = collectedMapped.length;
          mergeMappedItems(mapFolderPageItems(page?.value ?? [], sourceListName, sourceListObj));
          if (collectedMapped.length > before) publishAndCache(false);

          nextPath = toGraphPath(page?.['@odata.nextLink']);
          if (nextPath) await new Promise(r => setTimeout(r, 50));
        }
      };

      // 1) Live / primary list (eForms TaskListName, e.g. ProcTasks)
      await queryListWithFilters(listName, listObj);
      if (isStale()) return;

      // 2) Archive companion (ProcTasksArchive) - historical rows moved off the live list
      const archiveCompanions = this.taskLoader.getArchiveCompanionTaskLists(listName);
      if (collectedMapped.length === 0 && archiveCompanions.length > 0) {
        for (const archiveName of archiveCompanions) {
          const archiveList = this.taskLoader.findCachedSiteListByName(archiveName);
          if (!archiveList?.id) {
            console.warn(`[All Files] Archive list "${archiveName}" not found on site`);
            continue;
          }
          if (!silent && this.comments.isLoadingComments) {
            this.comments.commentsMessage = `Checking ${archiveName}...`;
            this.view.refreshView();
          }
          await queryListWithFilters(archiveName, archiveList);
          if (isStale()) return;
          if (collectedMapped.length > 0) break;
        }
      }

      // 3) Full scan only when still empty. Prefer scanning the archive companion
      //    (where old procurement tasks live) - never grind through live ProcTasks first.
      if (collectedMapped.length === 0) {
        const scanTargets =
          archiveCompanions.length > 0
            ? archiveCompanions
            : [listName];
        for (const scanName of scanTargets) {
          const scanList =
            scanName.toLowerCase() === listName.toLowerCase()
              ? listObj
              : this.taskLoader.findCachedSiteListByName(scanName);
          if (!scanList?.id) continue;
          await fullScanList(scanName, scanList);
          if (isStale()) return;
          if (collectedMapped.length > 0) break;
        }
      }

      if (isStale()) return;

      // Related Superior / sibling steps via eForm id - check primary + archive lists.
      const eFormKeys = collectEFormKeysFromMappedTasks(collectedMapped, folderName);
      if (eFormKeys.size > 0) {
        const relatedListTargets = [
          { name: listName, list: listObj },
          ...archiveCompanions
            .map((name) => ({ name, list: this.taskLoader.findCachedSiteListByName(name) }))
            .filter((entry): entry is { name: string; list: NonNullable<typeof entry.list> } => !!entry.list?.id),
        ];
        for (const target of relatedListTargets) {
          const relatedRaw = await this.taskQuery.fetchSharePointListItemsForEFormKeys(
            siteId,
            target.list.id,
            token,
            eFormKeys,
          );
          if (isStale()) return;

          const relatedMapped = mapFolderPageItems(relatedRaw, target.name, target.list);
          mergeMappedItems(relatedMapped);
        }
      }

      if (isStale()) return;
      if (softRefresh && collectedMapped.length === 0) return;
      publishAndCache(true);
    } catch (err: any) {
      if (isStale()) return;
      if (softRefresh || silent) return;
      this.comments.isLoadingMoreComments = false;
      this.comments.commentsMessage = `Failed to load tasks: ${err?.message || 'Unknown error'}`;
      this.comments.isLoadingComments = false;
      this.view.refreshView();
    } finally {
      this.folderTaskGraphInFlight = Math.max(0, this.folderTaskGraphInFlight - 1);
      if (!silent && !isStale() && (this.comments.isLoadingComments || this.comments.isLoadingMoreComments)) {
        this.comments.isLoadingComments = false;
        this.comments.isLoadingMoreComments = false;
        this.view.refreshView();
      }
    }
  }
}
