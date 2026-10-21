import { Injectable, Injector, inject } from '@angular/core';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { isTodoProcurementTaskList } from '../hr-task-lists.config';
import { getCommentsTabMobileState, getTodoTabMobileState } from '../components/mobile navigation/mobile-navigation';
import { SubmitterSelection } from '../components/todo-list/todo-list.component';
import { TodoService } from '../services/todo.service';
import { UserService } from '../services/user.service';
import { SubordinaryTaskService } from '../services/subordinaryTask.service';
import { FileCrawlCacheService } from '../services/file-crawl.service';
import { TodoTaskStoreService } from '../services/todo-task-store.service';
import { HrTaskMapperService } from '../services/hr-task-mapper.service';
import { SharePointTaskQueryService } from '../services/sharepoint-task-query.service';
import { TaskLoaderService } from '../services/task-loader.service';
import { TodoFetchService } from '../services/todo-fetch.service';
import { isSyntheticCommentCard, sortCommentItemsByDateDesc } from '../utils/comment-utils';
import { getHrTaskKey, mergeHrTasks } from '../utils/hr-task-matching';
import { sleep } from '../utils/timing';
import { ViewRefreshService } from './view-refresh.service';
import { UiStore } from './ui.store';
import { SessionStore } from './session.store';
import { CommentsStore } from './comments.store';
import { FolderTasksStore } from './folder-tasks.store';
import { AttachmentsStore } from './attachments.store';
import { HrFilesStore } from './hr-files.store';
import { AllFilesStore } from './all-files.store';

/**
 * The To Do list: first load, paging/"load more", background procurement drain, incremental refresh, subordinate mode and the To Do cache.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class TodoStore {
  private readonly injector = inject(Injector);
  private readonly todoService = inject(TodoService);
  private readonly userService = inject(UserService);
  private readonly subordinaryTaskService = inject(SubordinaryTaskService);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly todoStore = inject(TodoTaskStoreService);
  private readonly hrTaskMapper = inject(HrTaskMapperService);
  private readonly taskQuery = inject(SharePointTaskQueryService);
  private readonly taskLoader = inject(TaskLoaderService);
  private readonly todoFetch = inject(TodoFetchService);

  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get comments(): CommentsStore { return this.injector.get(CommentsStore); }
  private get folderTasks(): FolderTasksStore { return this.injector.get(FolderTasksStore); }
  private get attachments(): AttachmentsStore { return this.injector.get(AttachmentsStore); }
  private get hr(): HrFilesStore { return this.injector.get(HrFilesStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  isLoadingTodoTasks = false;

  /** Prevent stale To Do loads when the tab or subordinate toggle is used rapidly. */
  todoTaskLoadSeq = 0;

  /**
   * True after the first To Do Graph pass (assignee + blocked-list fallback) finishes.
   * False while still loading or after a mid-load cancel - used to decide resume vs full reload.
   */
  todoFirstPassComplete = false;

  /** True while Progress=Pending procurement pages are draining in the background. */
  todoProcurementDrainActive = false;

  /** Background drain chunks spent on the current To Do load (capped per load). */
  todoBackgroundDrainChunks = 0;

  /** Prevent stale Comments/HR task loads when task groups are clicked rapidly. */
  hrCommentsLoadSeq = 0;

  userHrTasksLoaded = false;

  selectedSubmitter: string | null = null;

  selectedEFormListId: string | null = null;

  selectedEFormTitle: string | null = null;

  selectedGroupTaskIds: string[] = [];

  superiorMode = false;

  // To Do paging state (cursors, watermarks, scanned lists) lives in TodoFetchService.
  isLoadingMoreTodoTasks = false;

  /** True while any To Do list still has unfetched older rows. */
  get hasMoreTodoTasks(): boolean {
    for (const cursor of this.todoFetch.todoListCursors.values()) {
      if (cursor) return true;
    }
    return false;
  }

  /**
   * True while a procurement list still has Progress=Pending pages left for the background
   * drain. Lists whose first scan stopped early keep their cursor for "Load older" only.
   */
  hasMoreProcurementTodoCursors(): boolean {
    for (const [listName, cursor] of this.todoFetch.todoListCursors.entries()) {
      if (cursor && isTodoProcurementTaskList(listName) && !this.todoFetch.todoProcScanStoppedEarly.has(listName)) {
        return true;
      }
    }
    return false;
  }

  /** True after a full To Do fetch (HR + procurement lists) has populated the task cache. */
  todoScopeCacheValid = false;

  /** Open To Do as the default section after sign-in instead of preloading HR Files. */
  initializeToDoOnLogin(): void {
    this.ui.showToDoSection = true;
    this.ui.showAllFilesSection = false;
    this.ui.hideComments = true;
    Object.assign(this.ui, getTodoTabMobileState());
    void this.loadTodoTasksForCurrentUser();
    // Prefetch HRPersonal group membership (SharePoint REST - does not use Graph slots).
    void this.hr.isUserInHrPersonalAllAccessGroup();
    this.session.startSharePointCachePoll();
    this.view.refreshView();
  }

  toggleToDoSection(): void {
    this.attachments.closePreview();
    // Drop HR Files / All Files folder scope so My Tasks comments are not loaded
    // for the last person opened in HR Files.
    this.allFiles.leaveAllFilesTaskView();
    this.ui.showToDoSection = true;
    this.ui.showAllFilesSection = false;
    this.ui.hideComments = true;
    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    Object.assign(this.ui, getTodoTabMobileState());
    this.comments.invalidateCommentFilters();
    this.ui.syncMobileSearchInput();
    // Instant paint from cache, then finish any interrupted load (never treat a
    // partial/cancelled snapshot as final - that caused intermittent missing tasks).
    this.restoreTodoTasksFromCache();
    if (!this.isLoadingTodoTasks) {
      this.ensureTodoLoadComplete();
    }
    this.view.refreshView();
  }

  /**
   * Pull the next slice of older rows from every To Do list that still has a cursor.
   * Nothing is re-fetched: each list resumes exactly where its first pass stopped.
   */
  async loadMoreTodoTasks(): Promise<void> {
    if (this.isLoadingMoreTodoTasks || !this.hasMoreTodoTasks) return;

    this.isLoadingMoreTodoTasks = true;
    this.view.refreshView();

    try {
      await this.fetchMoreTodoFromCursors(AppConstants.todoLoadMorePagesPerList);
    } catch {
      // Best effort - the button stays available so the user can retry.
    } finally {
      this.isLoadingMoreTodoTasks = false;
      this.view.refreshView();
    }
  }

  /** Resume list cursors for `pagesPerList` more Graph pages and merge into To Do. */
  async fetchMoreTodoFromCursors(
    pagesPerList: number,
    options: { procurementOnly?: boolean } = {},
  ): Promise<void> {
    if (!this.hasMoreTodoTasks) return;

    const token = await this.session.getSharePointToken();
    const siteId = this.session.cachedSiteId;
    const siteListsArr = this.session.cachedSiteLists;
    if (!siteId || siteListsArr.length === 0) return;

    const siteWebUrl = this.session.cachedSiteWebUrl;
    const userEmail = (this.session.currentUser?.email ?? '').toLowerCase();
    const userUpn = (this.session.currentUser?.userPrincipalName ?? '').toLowerCase();

    const pendingLists = [...this.todoFetch.todoListCursors.entries()]
      .filter((entry): entry is [string, string] => !!entry[1])
      .filter(([listName]) =>
        !options.procurementOnly ||
        (isTodoProcurementTaskList(listName) && !this.todoFetch.todoProcScanStoppedEarly.has(listName)),
      );

    if (pendingLists.length === 0) return;

    const fetchOne = async ([listName, cursor]: [string, string]): Promise<any[]> => {
      const list = this.taskLoader.findSiteList(siteListsArr, listName);
      if (!list?.id) {
        this.todoFetch.todoListCursors.delete(listName);
        return [];
      }

      try {
        const raw = isTodoProcurementTaskList(listName)
          ? await this.todoFetch.fetchProcurementTodoPendingItems(
            siteId, list.id, token, listName, cursor, pagesPerList,
          )
          : await this.todoFetch.fetchTodoScanPages(
            siteId, list.id, token, listName, cursor, pagesPerList,
          );

        const mapped: any[] = [];
        for (const item of raw ?? []) {
          const shaped = this.hrTaskMapper.mapSharePointItemToHrTask(
            item, listName, list, siteWebUrl, userEmail, userUpn, false, false,
          );
          if (shaped) mapped.push(shaped);
        }
        return mapped;
      } catch {
        // Drop the cursor so a transient Graph error cannot stall the drain forever.
        this.todoFetch.todoListCursors.delete(listName);
        return [];
      }
    };

    const fetched: any[] = [];
    const concurrency = AppConstants.hrTasksTodoListConcurrency;
    for (let i = 0; i < pendingLists.length; i += concurrency) {
      const batch = pendingLists.slice(i, i + concurrency);
      const results = await Promise.all(batch.map(entry => fetchOne(entry)));
      fetched.push(...results.flat());
    }

    if (fetched.length > 0) {
      const existing = this.todoStore.getAll() ?? [];
      const merged = mergeHrTasks(existing, fetched);
      this.todoStore.replaceAll(merged);
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), { clearLoading: false });
      this.userHrTasksLoaded = true;
      this.view.refreshView();
    }
  }

  /** Restore To Do from cache (memory or persisted snapshot) - instant when switching back or reopening the tab. */
  restoreTodoTasksFromCache(): boolean {
    const cached = this.todoStore.getAll();
    if (!Array.isArray(cached) || cached.length === 0) return false;

    this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(cached));
    this.userHrTasksLoaded = true;
    // Keep spinner state if a load is still running; only clear when idle.
    if (!this.isLoadingTodoTasks) {
      this.todoService.setLoading(false);
    }
    return true;
  }

  /**
   * After painting cache, finish an interrupted To Do load without a full re-fan-out when possible.
   * - First pass incomplete ? full reload.
   * - First pass done but Proc* pages remain ? resume drain only.
   * - Fully complete ? no Graph.
   */
  ensureTodoLoadComplete(): void {
    if (this.isLoadingTodoTasks || this.todoProcurementDrainActive) {
      return;
    }

    if (this.todoScopeCacheValid && !this.hasMoreProcurementTodoCursors()) {
      return;
    }

    if (this.todoFirstPassComplete && this.hasMoreProcurementTodoCursors()) {
      void this.resumeTodoProcurementDrain();
      return;
    }

    void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
  }

  /** Resume Progress=Pending procurement cursors left over after a tab switch cancel. */
  async resumeTodoProcurementDrain(): Promise<void> {
    if (!this.hasMoreProcurementTodoCursors()) {
      this.todoScopeCacheValid = this.todoFirstPassComplete;
      return;
    }

    const loadSeq = ++this.todoTaskLoadSeq;
    // Background only - do not re-show "Loading more tasks..." over an already painted list.
    this.attachments.pauseFolderPrefetch();
    this.view.refreshView();
    try {
      await this.drainProcurementTodoInBackground(loadSeq);
      if (!this.isStaleTodoTaskLoad(loadSeq) && this.todoFirstPassComplete) {
        this.todoScopeCacheValid = !this.hasMoreProcurementTodoCursors();
        void this.hr.warmHrPersonalRootFoldersCache();
      }
    } finally {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.attachments.folderPrefetchPaused = false;
        this.view.refreshView();
      }
    }
  }

  /** Load HR tasks assigned to the logged-in user into the To Do list. */
  async loadTodoTasksForCurrentUser(options: { forceRefresh?: boolean } = {}): Promise<void> {
    if (!this.session.currentUser) return;

    // Per-task To Do cache from Redis; the reads below use it for the instant first paint.
    await this.todoStore.load();

    const forceRefresh = options.forceRefresh === true;
    // A recent full crawl is already in Redis: top it up instead of re-paging every list.
    if (!forceRefresh && (await this.resumeTodoFromSavedCrawl())) return;

    // Stale (even day-old, persisted) tasks paint instantly; the fetch below refreshes them.
    const cached = this.todoStore.getAll();
    const hasCachedTasks = Array.isArray(cached) && cached.length > 0;

    if (hasCachedTasks && !forceRefresh) {
      this.restoreTodoTasksFromCache();
      this.view.refreshView();
    }

    const loadSeq = ++this.todoTaskLoadSeq;
    this.isLoadingTodoTasks = true;
    this.todoBackgroundDrainChunks = 0;
    this.todoFirstPassComplete = false;
    this.todoScopeCacheValid = false;
    // Yield Graph capacity to To Do - idle folder children crawls compete for the same budget.
    this.attachments.pauseFolderPrefetch();
    // A fresh load re-pages each list from the top, so old continuations are stale.
    this.todoFetch.todoListCursors.clear();
    this.todoFetch.todoProcScanStoppedEarly.clear();
    this.todoFetch.todoBlockedListsScanned.clear();

    if (hasCachedTasks) {
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(cached));
      this.userHrTasksLoaded = true;
      this.view.refreshView();
    } else {
      this.todoService.setLoading(true);
      this.view.refreshView();
    }

    // Progressive paints change the store every ~500 ms; upload once when the load ends
    // instead of a burst of PUT /tasks between Graph pages.
    this.todoStore.holdSaves();
    try {
      await this.session.ensureFormConfigLoaded();
      if (this.isStaleTodoTaskLoad(loadSeq)) return;

      const since = new Date();
      since.setUTCDate(since.getUTCDate() - AppConstants.hrTasksRecentFetchDays);
      const sinceIso = since.toISOString();

      // Always progressive + capped Proc* first slice so multiple Pr- tasks paint as they arrive.
      await this.loadHrTasksFromLists(
        false,
        { createdSinceIso: sinceIso },
        {
          updateTodo: true,
          updateCommentItems: false,
          kickoffOlderBackfill: false,
          skipCache: true,
          progressiveTodo: true,
          todoLoadSeq: loadSeq,
          todoAssigneeOnly: true,
          todoScope: true,
        }
      );
      if (this.isStaleTodoTaskLoad(loadSeq)) return;

      // First slice may already show recent assigned rows. Clear the empty-state
      // spinner once progressive paint has started; older Proc* matches still
      // arrive via the background drain without keeping "Loading more tasks...".
      this.userHrTasksLoaded = true;
      this.view.refreshView();

      // Limited fallback for lists found to reject AssignedToLookupId during this load
      // (known ones were scanned in the first pass). Progressive, so each list's cards
      // show as it finishes rather than after the whole fallback.
      if (this.taskQuery.blockedAssigneeLookupListIds.size > 0) {
        const fallbackItems = await this.loadHrTasksFromLists(
          false,
          { createdSinceIso: sinceIso },
          {
            updateTodo: true,
            updateCommentItems: false,
            kickoffOlderBackfill: false,
            skipCache: true,
            progressiveTodo: true,
            todoLoadSeq: loadSeq,
            todoAssigneeOnly: false,
            restrictToAssigneeBlockedLists: true,
            fastLoadPageLimit: AppConstants.hrTasksTodoLoadPageLimit,
            todoScope: true,
          }
        );
        if (!this.isStaleTodoTaskLoad(loadSeq) && fallbackItems.length > 0) {
          const cachedItems = this.todoStore.getAll() ?? [];
          const merged = mergeHrTasks(cachedItems, fallbackItems);
          this.todoStore.replaceAll(merged);
          this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), { clearLoading: false });
          this.view.refreshView();
        }
      }

      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        const folderAssigned = this.collectAssignedProcTasksFromFolderCaches();
        if (folderAssigned.length > 0) {
          const cachedItems = this.todoStore.getAll() ?? [];
          const merged = mergeHrTasks(cachedItems, folderAssigned);
          this.todoStore.replaceAll(merged);
          this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), { clearLoading: false });
          this.view.refreshView();
        }
      }

      if (this.isStaleTodoTaskLoad(loadSeq)) return;

      // First Graph pass done - paint without holding "Loading more tasks...".
      // Remaining Proc* pages continue in the background drain below.
      this.todoFirstPassComplete = true;
      this.isLoadingTodoTasks = false;
      this.todoService.setLoading(false);
      this.view.refreshView();

      await this.drainProcurementTodoInBackground(loadSeq);
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.todoScopeCacheValid = !this.hasMoreProcurementTodoCursors();
        this.userHrTasksLoaded = true;
        // Older Proc* rows left on cursors are fetched by "Load older"; the normal load is done.
        this.persistTodoMeta(true);
        // Warm HR Files people list only after To Do Graph work settles.
        void this.hr.warmHrPersonalRootFoldersCache();
      }
    } catch {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.todoService.setLoading(false);
      }
    } finally {
      this.todoStore.releaseSaves();
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.isLoadingTodoTasks = false;
        this.todoService.setLoading(false);
        this.attachments.folderPrefetchPaused = false;
        this.view.refreshView();
      }
    }
  }

  /** A saved full crawl older than this is redone (top-ups only watch lists that had rows). */
  static readonly TODO_CRAWL_MAX_AGE_MS = 12 * 60 * 60_000;

  /**
   * Reload path: paint the saved To Do and read only what changed since the saved
   * per-list watermarks (about one Graph request per list) instead of a full crawl.
   * Returns false when there is no usable saved crawl, so the caller loads normally.
   */
  async resumeTodoFromSavedCrawl(): Promise<boolean> {
    const meta = this.todoStore.getMeta();
    const cached = this.todoStore.getAll();
    const watermarks = Object.entries(meta?.watermarks ?? {});
    if (!meta?.complete || !cached?.length || watermarks.length === 0) {
      console.info(
        `[todo] full crawl: no usable saved state (meta=${!!meta}, complete=${!!meta?.complete}, ` +
        `tasks=${cached?.length ?? 0}, lists=${watermarks.length})`,
      );
      return false;
    }
    if (Date.now() - (meta.crawledAt ?? 0) > TodoStore.TODO_CRAWL_MAX_AGE_MS) {
      console.info('[todo] full crawl: saved crawl is older than 12 hours');
      return false;
    }
    console.info(`[todo] top-up only: ${watermarks.length} lists, ${cached.length} saved tasks`);

    if (this.todoFetch.todoListWatermarks.size === 0) {
      for (const [listName, watermark] of watermarks) this.todoFetch.todoListWatermarks.set(listName, watermark);
    }
    this.todoFetch.todoListCursors.clear();
    this.todoFetch.todoProcScanStoppedEarly.clear();
    this.todoFirstPassComplete = true;
    this.todoScopeCacheValid = true;
    this.restoreTodoTasksFromCache();
    this.view.refreshView();

    try {
      const token = await this.session.getSharePointToken();
      await this.session.ensureFormConfigLoaded();
      await this.session.ensureSiteMetadata(sharePointConfig.siteHostName, sharePointConfig.sitePath, token);
      await this.refreshTodoTasksIncrementally();
    } catch {
      // The saved list stays on screen; the next poll or refresh tops it up.
    }
    void this.hr.warmHrPersonalRootFoldersCache();
    return true;
  }

  /** Save the per-list watermarks so the next reload can top up instead of re-crawling. */
  persistTodoMeta(fullCrawlFinished: boolean): void {
    if (this.todoFetch.todoListWatermarks.size === 0) return;
    const previous = this.todoStore.getMeta();
    if (!fullCrawlFinished && !previous?.complete) return;
    this.todoStore.setMeta({
      watermarks: Object.fromEntries(this.todoFetch.todoListWatermarks),
      complete: true,
      crawledAt: fullCrawlFinished ? Date.now() : previous?.crawledAt,
    });
  }

  /** Continue Progress=Pending procurement cursors without holding the To Do spinner. */
  async drainProcurementTodoInBackground(loadSeq: number): Promise<void> {
    this.todoProcurementDrainActive = true;
    try {
      // Capped and gated: leftover cursors stay for "Load more tasks" instead of
      // scanning whole Proc* lists in the background on every load.
      while (
        !this.isStaleTodoTaskLoad(loadSeq) &&
        this.hasMoreProcurementTodoCursors() &&
        this.todoBackgroundDrainChunks < AppConstants.todoProcBackgroundMaxChunks &&
        this.session.backgroundGraphAllowed()
      ) {
        this.todoBackgroundDrainChunks += 1;
        await this.fetchMoreTodoFromCursors(AppConstants.todoProcBackgroundPagesPerChunk, {
          procurementOnly: true,
        });
      }
    } catch {
      // Background only.
    } finally {
      if (loadSeq === this.todoTaskLoadSeq) {
        this.todoProcurementDrainActive = false;
      }
    }
  }

  isStaleTodoTaskLoad(loadSeq: number): boolean {
    return loadSeq !== this.todoTaskLoadSeq;
  }

  /** Stop in-flight To Do fetches so HR Files / All Files get Graph quota. */
  cancelBackgroundTodoLoad(): void {
    const interrupted =
      this.isLoadingTodoTasks ||
      this.todoProcurementDrainActive ||
      !this.todoFirstPassComplete ||
      this.hasMoreProcurementTodoCursors() ||
      !this.todoScopeCacheValid;
    this.todoTaskLoadSeq++;
    this.todoProcurementDrainActive = false;
    if (interrupted) {
      // Keep todoFirstPassComplete + cursors so returning to To Do can resume drain
      // instead of treating a partial Redis snapshot as complete.
      this.todoScopeCacheValid = false;
    }
    if (this.isLoadingTodoTasks) {
      this.isLoadingTodoTasks = false;
      this.todoService.setLoading(false);
      this.view.refreshView();
    }
  }

  onSuperiorModeChange(checked: boolean): void {
    this.superiorMode = checked;
    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    this.comments.invalidateCommentFilters();
    if (checked) {
      void this.reloadTodoTasksWithSubordinates();
    } else {
      void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
    }
    this.view.refreshView();
  }

  async reloadTodoTasksWithSubordinates(): Promise<void> {
    if (!this.session.currentUser) return;

    const identifiers = [
      this.session.currentUser.email,
      this.session.currentUser.userPrincipalName,
      this.session.currentUser.username,
    ].map(value => (value ?? '').trim()).filter(Boolean);

    if (!identifiers.length) return;

    const loadSeq = ++this.todoTaskLoadSeq;
    this.isLoadingTodoTasks = true;
    this.attachments.pauseFolderPrefetch();
    this.todoService.setLoading(true);
    this.view.refreshView();

    const since = new Date();
    since.setUTCDate(since.getUTCDate() - AppConstants.hrTasksRecentFetchDays);



    try {
      //resolve managers direct reports frist(this wil be cahced after the first call)
      const Subordinate = await this.subordinaryTaskService.ensureSubordinatesForManager(
        identifiers,
        false
      )
      if (this.isStaleTodoTaskLoad(loadSeq)) return;

       // NEW: nobody reports to this user, so there are no subordinate tasks to find.
      // Skip the scan of every task list (~77 Graph requests / ~55 s) and let the
      // "No employees found under your manager account." message show straight away.
      // `finally` below still clears the loading flags because we return from inside try.
      if(Subordinate.length === 0){
        this.userHrTasksLoaded = true;
        return;
      }
       
      await this.loadHrTasksFromLists(
        false,
        { createdSinceIso: since.toISOString() },
        {
          updateTodo: true,
          updateCommentItems: false,
          kickoffOlderBackfill: false,
          subordinateTasksOnly: true,
          fastLoadPageLimit: AppConstants.hrTasksSubordinateLoadPageLimit,
          progressiveTodo: true,
          todoLoadSeq: loadSeq,
        }
      );
      if (this.isStaleTodoTaskLoad(loadSeq)) return;

      this.userHrTasksLoaded = true;
    } catch {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.todoService.setLoading(false);
      }
    } finally {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.isLoadingTodoTasks = false;
        this.todoService.setLoading(false);
        this.attachments.folderPrefetchPaused = false;
        this.view.refreshView();
      }
    }
  }

  // Receives submitter + eFormListId + eFormTitle from todo-list after a grouped row is clicked.
  onSubmitterSelected(selection: SubmitterSelection | null): void {
    this.selectedSubmitter = selection?.name ?? null;
    this.selectedEFormListId = selection?.eFormListId ?? null;
    this.selectedEFormTitle = selection?.eFormTitle ?? null;
    this.selectedGroupTaskIds = (selection?.taskIds ?? [])
      .map(id => String(id ?? '').trim())
      .filter(Boolean);

    const rawTaskId = String(selection?.taskId ?? '').trim();
    const numericTaskId = Number(rawTaskId);
    this.comments.selectedTask = rawTaskId
      ? { Id: !Number.isNaN(numericTaskId) ? numericTaskId : rawTaskId }
      : null;

    const task = this.todoService.getTaskById(rawTaskId);
    const commentItem = this.comments.commentItems.find(item => String(item.id ?? '').trim() === rawTaskId);
    const groupedItem = this.comments.commentItems.find(
      item =>
        String(item.eFormDetails?.eFormListId ?? '').trim() ===
        String(selection?.eFormListId ?? '').trim(),
    );
    this.comments.selectedCommentListName =
      String((commentItem as { listName?: string })?.listName ?? '').trim() ||
      String(commentItem?.eFormDetails?.listName ?? '').trim() ||
      String(groupedItem?.eFormDetails?.listName ?? (groupedItem as { listName?: string })?.listName ?? '').trim() ||
      String(task?.eFormDetails?.listName ?? '').trim();

    this.comments.invalidateCommentFilters();
    if (selection) {
      if (this.ui.showToDoSection) {
        this.folderTasks.selectedHrPersonalTaskFolder = '';
        this.allFiles.selectedFolderName = '';
      }
      this.ui.hideComments = false;
      Object.assign(this.ui, getCommentsTabMobileState());
      this.comments.closeFormViewer();
      // Comments panel filters locally. Only hit Graph when we have nothing cached
      // (every accordion click used to call loadHrTasks and re-fan-out all lists).
      const hasLocalComments = this.comments.commentItems.length > 0;
      const cachedTasks = this.todoStore.getAll();
      if (!hasLocalComments && Array.isArray(cachedTasks) && cachedTasks.length > 0) {
        this.comments.commentItems = sortCommentItemsByDateDesc(cachedTasks);
        this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
        this.comments.invalidateCommentFilters();
        this.comments.isLoadingComments = false;
        this.comments.commentsMessage = '';
        this.userHrTasksLoaded = true;
      } else if (!hasLocalComments && !this.userHrTasksLoaded && !this.comments.isLoadingComments) {
        this.loadHrTasks();
      }
      if (this.ui.showToDoSection) {
        void this.attachments.loadTodoTaskAttachments(task, selection.eFormTitle ?? '');
      }
    } else {
      this.comments.closeFormViewer();
    }
    this.view.refreshView();
  }

  refreshToDoSection(): void {
    // Subordinate view has its own load path and no per-list watermarks.
    if (this.superiorMode) {
      this.todoStore.invalidate();
      void this.reloadTodoTasksWithSubordinates();
      this.view.refreshView();
      return;
    }

    // Nothing cached yet - nothing to top up, so do the normal first load.
    if (this.todoFetch.todoListWatermarks.size === 0) {
      void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
      this.view.refreshView();
      return;
    }

    void this.refreshTodoTasksIncrementally();

    if (this.selectedSubmitter) {
      this.loadHrTasks();
    }

    this.view.refreshView();
  }

  /**
   * Top up To Do with whatever changed since the last look, keeping everything already
   * collected. Each list is read newest-first and stops at the first row already held,
   * so no previously fetched row is downloaded, re-mapped or re-filtered a second time.
   */
  async refreshTodoTasksIncrementally(): Promise<boolean> {
    if (this.isLoadingMoreTodoTasks || this.isLoadingTodoTasks) return false;

    this.session.lastTodoTopUpAtMs = Date.now();
    this.isLoadingMoreTodoTasks = true;
    this.view.refreshView();

    try {
      const token = await this.session.getSharePointToken();
      const siteId = this.session.cachedSiteId;
      const siteListsArr = this.session.cachedSiteLists;
      if (!siteId || siteListsArr.length === 0) {
        void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
        return false;
      }

      const siteWebUrl = this.session.cachedSiteWebUrl;
      const userEmail = (this.session.currentUser?.email ?? '').toLowerCase();
      const userUpn = (this.session.currentUser?.userPrincipalName ?? '').toLowerCase();
      const watermarks = [...this.todoFetch.todoListWatermarks.entries()];

      const refreshOne = async ([listName, watermark]: [string, string]): Promise<any[]> => {
        const list = siteListsArr.find(
          (l: any) =>
            (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
            (l.displayName ?? '').toLowerCase() === listName.toLowerCase(),
        );
        if (!list?.id) return [];

        const raw = await this.todoFetch.fetchTodoRowsNewerThanWatermark(
          siteId, list.id, token, listName, watermark,
        );
        if (raw.length === 0) return [];

        this.todoFetch.noteTodoWatermark(listName, raw);

        const mapped: any[] = [];
        for (const item of raw) {
          const shaped = this.hrTaskMapper.mapSharePointItemToHrTask(
            item, listName, list, siteWebUrl, userEmail, userUpn, false, false,
          );
          if (shaped) mapped.push(shaped);
        }
        return mapped;
      };

      const fresh: any[] = [];
      const concurrency = AppConstants.hrTasksTodoListConcurrency;
      for (let i = 0; i < watermarks.length; i += concurrency) {
        const batch = watermarks.slice(i, i + concurrency);
        const results = await Promise.all(batch.map(entry => refreshOne(entry).catch(() => [])));
        fresh.push(...results.flat());
      }

      // Merge on top of the cache so previously loaded pages (including anything pulled
      // via "Load older tasks") survive the refresh.
      const existing = this.todoStore.getAll() ?? [];
      const merged = fresh.length > 0 ? mergeHrTasks(existing, fresh) : existing;
      if (fresh.length > 0) {
        this.todoStore.replaceAll(merged);
      }
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged));
      this.userHrTasksLoaded = true;
      this.persistTodoMeta(false);
      return fresh.length > 0;
    } catch {
      // Refresh is best effort - the existing view stays as-is.
      return false;
    } finally {
      this.isLoadingMoreTodoTasks = false;
      this.view.refreshView();
    }
  }

  /** Pull assignee-matched procurement tasks already loaded via All Files folder views. */
  collectAssignedProcTasksFromFolderCaches(): any[] {
    const matchName = this.session.currentUser?.username ?? this.session.currentUser?.userPrincipalName ?? '';
    const matchEmail = (this.session.currentUser?.email ?? this.session.currentUser?.userPrincipalName ?? '').toLowerCase();
    if (!matchName && !matchEmail) return [];

    const byKey = new Map<string, any>();
    for (const entry of this.fileCrawlCache.getEntriesByPrefix('tasks:lib:v7:')) {
      if (!/proctasks/i.test(entry.key)) continue;
      const items = entry.data;
      if (!Array.isArray(items)) continue;

      for (const item of items) {
        if (!isTodoProcurementTaskList(String(item?.listName ?? ''))) continue;
        if (!this.isMappedTaskAssignedToLoggedInUser(item, matchName, matchEmail)) continue;
        byKey.set(getHrTaskKey(item), {
          ...item,
          isAssignedToCurrentUser: true,
        });
      }
    }
    return [...byKey.values()];
  }

  isMappedTaskAssignedToLoggedInUser(
    item: any,
    matchName: string,
    matchEmail: string,
  ): boolean {
    const assignedTo = String(item?.eFormDetails?.assignedTo ?? item?.assignedTo ?? '').trim();
    if (assignedTo && this.userService.matchesAssigneeField(assignedTo, matchName, matchEmail)) {
      return true;
    }

    const rawFields = item?.eFormDetails?.rawFields;
    if (rawFields && typeof rawFields === 'object') {
      return this.hrTaskMapper.isAssignedToUserFromFields(rawFields as Record<string, any>, matchName, matchEmail);
    }

    return false;
  }

  getHrTasksForTodoList<T extends { id?: unknown }>(items: T[]): T[] {
    return items.filter(item => !isSyntheticCommentCard(item));
  }

  /**
   * Remove a completed/delegated task from the To Do Redis snapshot (`tasks:hr-user`)
   * so the next instant paint does not bring it back.
   */
  removeTaskFromTodoCache(taskId: string): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;

    const cached = this.todoStore.getAll();
    if (!Array.isArray(cached) || cached.length === 0) return;

    const next = cached.filter((item: any) => String(item?.id ?? '').trim() !== id);
    if (next.length === cached.length) return;

    this.todoStore.replaceAll(next);
  }

  /** Merge field updates into the To Do Redis snapshot for one task id. */
  patchTaskInTodoCache(taskId: string, patch: (item: any) => any): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;

    const cached = this.todoStore.getAll();
    if (!Array.isArray(cached) || cached.length === 0) return;

    let changed = false;
    const next = cached.map((item: any) => {
      if (String(item?.id ?? '').trim() !== id) return item;
      changed = true;
      return patch(item);
    });
    if (!changed) return;

    this.todoStore.replaceAll(next);
  }

  async loadOlderHrTasks(cutoffIso: string): Promise<void> {
    if (this.folderTasks.hrOlderHistoryLoadInProgress) return;
    this.folderTasks.hrOlderHistoryLoadInProgress = true;
    this.view.refreshView();
    try {
      const older = await this.loadHrTasksFromLists(false, { createdBeforeIso: cutoffIso });
      if (!older.length || this.folderTasks.isFolderTaskCommentsViewActive()) return;
      this.comments.commentItems = mergeHrTasks(this.comments.commentItems, older);
      this.comments.invalidateCommentFilters();
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(this.comments.commentItems));
    } catch { /* ignore */ }
    finally {
      this.folderTasks.hrOlderHistoryLoadInProgress = false;
      this.view.refreshView();
    }
  }

  // LOAD HR TASKS - public entry point
  loadHrTasks(): void {
    if (!this.session.currentUser) {
      this.ui.showModal('Login Required', 'You must log in first to view HR Tasks.', 'error');
      this.comments.commentsMessage = 'Please log in first.';
      return;
    }

    if (!this.ui.showToDoSection && this.folderTasks.isFolderTaskCommentsViewActive()) {
      const folder = String(this.folderTasks.selectedHrPersonalTaskFolder || this.allFiles.selectedFolderName || '').trim();
      const library = this.ui.showAllFilesSection
        ? String(this.attachments.currentLibraryName ?? '').trim()
        : this.hr.targetLibraryName;
      if (folder && library) {
        this.folderTasks.loadAllFilesTasksForLibrary(library, folder);
        return;
      }
    }

    const cached = this.todoStore.getFresh();
    if (cached?.length) {
      this.comments.commentItems = sortCommentItemsByDateDesc(cached);
      this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
      this.comments.invalidateCommentFilters();
      this.comments.isLoadingComments = false;
      this.comments.isLoadingMoreComments = false;
      this.comments.commentsMessage = '';
      this.userHrTasksLoaded = true;
      this.view.refreshView();
      return;
    }

    const loadSeq = ++this.hrCommentsLoadSeq;
    this.comments.isLoadingComments = true;
    this.comments.isLoadingMoreComments = false;
    this.comments.commentItems = [];
    this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
    this.comments.invalidateCommentFilters();
    this.view.refreshView();
    void this.loadHrTasksWithRecentWindow(loadSeq);
  }

  isStaleHrCommentsLoad(loadSeq: number): boolean {
    return loadSeq !== this.hrCommentsLoadSeq;
  }

  async loadHrTasksWithRecentWindow(loadSeq: number): Promise<void> {
    try {
      await this.session.ensureFormConfigLoaded();
      if (this.isStaleHrCommentsLoad(loadSeq)) return;

      const since = new Date();
      since.setUTCDate(since.getUTCDate() - AppConstants.hrTasksRecentFetchDays);
      const sinceIso = since.toISOString();
      await this.loadHrTasksFromLists(
        true,
        { createdSinceIso: sinceIso },
        { hrCommentsLoadSeq: loadSeq }
      );
    } catch {
      if (!this.isStaleHrCommentsLoad(loadSeq)) {
        this.comments.commentsMessage = 'Failed to load HR Tasks.';
        this.comments.isLoadingComments = false;
        this.view.refreshView();
      }
    }
  }

  // LOAD HR TASKS - core async logic
  // All SharePoint lists queried IN PARALLEL via Promise.all().
  // Optional Graph $filter on list item createdDateTime reduces pages
  // for the "recent" and silent-refresh paths.
  async loadHrTasksFromLists(
    updateComponentState = true,
    range: { createdSinceIso?: string; createdBeforeIso?: string } | null = null,
    options: {
      kickoffOlderBackfill?: boolean;
      updateTodo?: boolean;
      updateCommentItems?: boolean;
      subordinateTasksOnly?: boolean;
      fastLoadPageLimit?: number;
      hrCommentsLoadSeq?: number;
      skipCache?: boolean;
      progressiveTodo?: boolean;
      todoLoadSeq?: number;
      /** To Do fast path: AssignedToLookupId only - skip slow page-scan fallback. */
      todoAssigneeOnly?: boolean;
      /** Page-scan only lists where AssignedToLookupId filter previously failed. */
      restrictToAssigneeBlockedLists?: boolean;
      /** To Do panel: include SAP/procurement lists + archive companions (e.g. ProcTasksArchive). */
      todoScope?: boolean;
    } = {}
  ): Promise<any[]> {
    const updateCommentItems = options.updateCommentItems ?? updateComponentState;
    const updateTodo = options.updateTodo ?? updateComponentState;
    const kickoffOlderBackfill = options.kickoffOlderBackfill ?? updateComponentState;
    const subordinateTasksOnly = options.subordinateTasksOnly === true;
    const fastLoadPageLimit = options.fastLoadPageLimit;
    const hrCommentsLoadSeq = options.hrCommentsLoadSeq;
    const progressiveTodo = options.progressiveTodo === true;
    const todoLoadSeq = options.todoLoadSeq;
    const todoAssigneeOnly = options.todoAssigneeOnly === true;
    const restrictToAssigneeBlockedLists = options.restrictToAssigneeBlockedLists === true;
    const todoScope = options.todoScope === true;
    const isStaleCommentsLoad = (): boolean =>
      hrCommentsLoadSeq !== undefined && this.isStaleHrCommentsLoad(hrCommentsLoadSeq);
    const isStaleTodoLoad = (): boolean =>
      todoLoadSeq !== undefined && this.isStaleTodoTaskLoad(todoLoadSeq);

    // Repaints are spaced out so rows trickle in instead of re-sorting the list on every page.
    let lastPartialPublishAt = 0;
    let partialTimer: ReturnType<typeof setTimeout> | null = null;
    let latestPartial: any[] = [];
    const publishPartialTodo = (partialItems: any[], flushNow = false): void => {
      if (!updateTodo || !progressiveTodo || isStaleTodoLoad()) return;
      if (partialItems.length === 0) return;
      latestPartial = partialItems;
      const wait = AppConstants.todoPartialPublishMs - (Date.now() - lastPartialPublishAt);
      if (!flushNow && wait > 0) {
        partialTimer ??= setTimeout(() => {
          partialTimer = null;
          publishPartialTodo(latestPartial, true);
        }, wait);
        return;
      }
      if (partialTimer) {
        clearTimeout(partialTimer);
        partialTimer = null;
      }
      lastPartialPublishAt = Date.now();
      const existing = this.todoStore.getAll() ?? [];
      const merged = existing.length > 0
        ? mergeHrTasks(existing, partialItems)
        : [...partialItems];
      const sorted = [...merged].sort(
        (a: { submittedDate: any }, b: { submittedDate: any }) =>
          new Date(b.submittedDate || 0).getTime() - new Date(a.submittedDate || 0).getTime()
      );
      this.todoStore.replaceAll(sorted);
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(sorted), { clearLoading: false });
      this.userHrTasksLoaded = true;
      this.view.refreshView();
    };

    const isCacheableRecentUserLoad =
      !subordinateTasksOnly &&
      !!range?.createdSinceIso &&
      !range?.createdBeforeIso;

    // Serve recent user tasks from cache unless this is a forced network refresh.
    if (isCacheableRecentUserLoad && !options.skipCache) {
      const cached = this.todoStore.getFresh();
      const cacheOk = !todoScope || this.todoScopeCacheValid;
      if (cached?.length && cacheOk) {
        if (updateCommentItems && !isStaleCommentsLoad()) {
          if (!this.folderTasks.isFolderTaskCommentsViewActive()) {
            this.comments.commentItems = cached;
            this.comments.invalidateCommentFilters();
            this.comments.commentsMessage = '';
            this.comments.isLoadingComments = false;
            this.userHrTasksLoaded = true;
            this.view.refreshView();
          } else {
            this.comments.isLoadingComments = false;
            this.view.refreshView();
          }
        }
        if (updateTodo) {
          this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(cached));
          this.userHrTasksLoaded = true;
        }
        return cached;
      }
    }

    if(subordinateTasksOnly){
      await this.subordinaryTaskService.ensureSubordinatesForManager([
        this.session.currentUser?.email?? '',
        this.session.currentUser?.userPrincipalName??'',
        this.session.currentUser?.username??'',   
      ].map(value => value.trim()).filter(Boolean),false);
    
    }

    try {
      const token = await this.session.getSharePointToken();
      const siteHost = sharePointConfig.siteHostName;
      const sitePath = sharePointConfig.sitePath;

      await this.session.ensureSiteMetadata(siteHost, sitePath, token);

      const siteId: string = this.session.cachedSiteId!;
      const siteWebUrl: string = this.session.cachedSiteWebUrl;
      const siteListsArr: any[] = this.session.cachedSiteLists;

      const userEmail = (this.session.currentUser?.email ?? '').toLowerCase();
      const userUpn = (this.session.currentUser?.userPrincipalName ?? '').toLowerCase();

      if (subordinateTasksOnly) {
        await this.subordinaryTaskService.ensureSubordinatesForManager([
          this.session.currentUser?.email ?? '',
          this.session.currentUser?.userPrincipalName ?? '',
          this.session.currentUser?.username ?? '',
        ].map(value => value.trim()).filter(Boolean), true);
      }

      // Names come from several sources (config, eForms, archive companions); two of them
      // can resolve to the same list, which used to page that list twice in parallel.
      let listsToQuery = this.taskLoader.dedupeTaskListsBySiteList(
        this.taskLoader.getTaskListsToQuery(siteListsArr, {
          includeAllWorkflows: todoScope,
          includeArchiveCompanions: todoScope,
        }),
        siteListsArr,
      );

      if (restrictToAssigneeBlockedLists) {
        listsToQuery = listsToQuery.filter((listName: string) => {
          // Procurement lists are already covered by the pending/newest-first pass;
          // re-scanning them here would just page through their oldest rows again.
          if (isTodoProcurementTaskList(listName)) return false;
          // Lists known to be blocked before this load were already scanned in the first pass.
          if (this.todoFetch.todoBlockedListsScanned.has(listName)) return false;
          const list = this.taskLoader.findSiteList(siteListsArr, listName);
          return !!list?.id && this.taskQuery.blockedAssigneeLookupListIds.has(list.id);
        });
      } else {
        listsToQuery = listsToQuery.sort((a: string, b: string) => {
          const rank = (name: string): number => {
            const key = name.toLowerCase();
            // Quick assignee-filtered lists go first so rows appear straight away; the slow
            // Proc* scans (many pages each) run after them and stream in slice by slice.
            if (key === 'proctasksarchive') return 5;
            if (isTodoProcurementTaskList(name)) return 4;
            const list = this.taskLoader.findSiteList(siteListsArr, name);
            if (!list?.id) return 3;
            if (this.taskQuery.getListFilterFieldStatus(list.id, 'AssignedToLookupId') === true) return 0;
            if (this.taskQuery.blockedAssigneeLookupListIds.has(list.id)) return 3;
            return 1;
          };
          return rank(a) - rank(b);
        });
      }

      if (listsToQuery.length === 0) {
        return [];
      }

      const assigneeLookupId = !subordinateTasksOnly
        ? await this.taskQuery.resolveSharePointUserLookupId(userEmail || userUpn, token)
        : null;

      const listResults: any[] = [];
      const listConcurrency = todoAssigneeOnly || progressiveTodo
        ? AppConstants.hrTasksTodoListConcurrency
        : AppConstants.hrFilesTaskListConcurrency;

      const fetchOneList = (listName: string, onPartial?: (items: any[]) => void) =>
        this.todoFetch.fetchItemsForList(
          listName,
          siteListsArr,
          siteId,
          siteWebUrl,
          token,
          userEmail,
          userUpn,
          range,
          subordinateTasksOnly,
          fastLoadPageLimit,
          assigneeLookupId,
          todoAssigneeOnly,
          todoAssigneeOnly || restrictToAssigneeBlockedLists,
          onPartial,
        );

      if (todoAssigneeOnly && progressiveTodo) {
        const merged: any[] = [];
        let nextIndex = 0;
        const workers = Array.from(
          { length: Math.min(listConcurrency, listsToQuery.length) },
          async () => {
            while (true) {
              const index = nextIndex++;
              if (index >= listsToQuery.length) break;
              // Slices of a long Proc* scan are shown as they land; `merged` only gets the
              // list's rows once it finishes, so nothing is counted twice.
              const items = await fetchOneList(listsToQuery[index], (slice) =>
                publishPartialTodo([...merged, ...slice]),
              );
              merged.push(...items);
              publishPartialTodo(merged);
            }
          },
        );
        await Promise.all(workers);
        if (partialTimer) {
          clearTimeout(partialTimer);
          partialTimer = null;
        }
        publishPartialTodo(merged, true);
        listResults.push(merged);
      } else {
        for (let i = 0; i < listsToQuery.length; i += listConcurrency) {
          const batch = listsToQuery.slice(i, i + listConcurrency);
          const batchResults = await Promise.all(batch.map(listName => fetchOneList(listName)));
          listResults.push(...batchResults);
          publishPartialTodo(listResults.flat());
          if (i + listConcurrency < listsToQuery.length && !progressiveTodo) {
            await sleep(AppConstants.hrTaskListBatchGapMs);
          }
        }
      }

      const allItems = listResults.flat();

      allItems.sort(
        (a: { submittedDate: any; }, b: { submittedDate: any; }) =>
          new Date(b.submittedDate || 0).getTime() - new Date(a.submittedDate || 0).getTime()
      );

      if (isCacheableRecentUserLoad) {
        if (todoScope) {
          if (allItems.length > 0) {
            // Always merge - never replace an earlier pass with a later subset.
            const existing = this.todoStore.getAll() ?? [];
            this.todoStore.replaceAll(
              existing.length > 0 ? mergeHrTasks(existing, allItems) : allItems,
            );
          }
          // Do NOT set todoScopeCacheValid here - progressive To Do still has Proc*
          // drain / blocked-list fallback after this pass. Callers mark complete.
        } else if (allItems.length > 0) {
          const existing = this.todoStore.getAll() ?? [];
          const procItems = existing.filter((item: any) =>
            isTodoProcurementTaskList(String(item.listName ?? ''))
          );
          this.todoStore.replaceAll(
            mergeHrTasks(allItems, procItems),
          );
        }
      }

      if (updateCommentItems) {
        if (isStaleCommentsLoad()) return allItems;
        if (this.folderTasks.isFolderTaskCommentsViewActive()) {
          this.comments.isLoadingComments = false;
          this.view.refreshView();
          return allItems;
        }
        this.comments.commentItems = allItems;
        this.comments.invalidateCommentFilters();
        this.comments.commentsMessage = '';
        this.comments.isLoadingComments = false;
        this.userHrTasksLoaded = true;
        this.view.refreshView();
        if (kickoffOlderBackfill && range?.createdSinceIso && !isStaleCommentsLoad()) {
          void this.loadOlderHrTasks(range.createdSinceIso);
        }
      }

      if (updateTodo) {
        if (!isStaleTodoLoad()) {
          if (allItems.length > 0) {
            const existing = this.todoStore.getAll() ?? [];
            const merged = existing.length > 0 ? mergeHrTasks(existing, allItems) : allItems;
            this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), {
              clearLoading: !progressiveTodo,
            });
          }
          this.userHrTasksLoaded = true;
        }
      }

      return allItems;
    } catch (error: unknown) {
      if (updateCommentItems) {
        if (!isStaleCommentsLoad()) {
          this.comments.commentsMessage = `Failed to load HR Tasks: ${(error as any)?.message || 'Unknown error'}.`;
          this.comments.isLoadingComments = false;
          this.view.refreshView();
        }
      }
      if (updateTodo) {
        this.todoService.setLoading(false);
      }
      return [];
    }
  }
}
