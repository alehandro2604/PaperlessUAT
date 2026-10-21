import { Injectable, Injector, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Subscription } from 'rxjs';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { graphGetWithRetry, normalizeName } from '../microsoft-graph';
import { FormConfigurationService } from '../services/form-configuration.service';
import { getAllFilesTabMobileState, getAttachmentsTabMobileState, getCommentsTabMobileState } from '../components/mobile navigation/mobile-navigation';
import { AttachmentSearchHit, CommentSearchHit } from '../components/All-files/all-files';
import { FileCrawlCacheService } from '../services/file-crawl.service';
import { DriveFilesService } from '../services/drive-files.service';
import { buildQuickCommentItemFromSearchHit } from '../utils/comment-utils';
import { extractTrailingFolderId, extractChildSubjectNamesFromFields } from '../utils/hr-task-matching';
import { hrFolderTaskCacheKey, libFolderTaskCacheKey, driveFolderCacheKey, childSubjectCacheKey } from '../utils/cache-keys';
import { ViewRefreshService } from './view-refresh.service';
import { UiStore } from './ui.store';
import { SessionStore } from './session.store';
import { CommentsStore } from './comments.store';
import { FolderTasksStore } from './folder-tasks.store';
import { TodoStore } from './todo.store';
import { AttachmentsStore } from './attachments.store';
import { HrFilesStore } from './hr-files.store';

/**
 * All Files: reacting to clicks in the All Files component (folder, comment hit, attachment hit), library selection and the ChildSubject dropdown.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class AllFilesStore {
  private readonly injector = inject(Injector);
  private readonly http = inject(HttpClient);
  private readonly formConfigService = inject(FormConfigurationService);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly driveFiles = inject(DriveFilesService);

  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get comments(): CommentsStore { return this.injector.get(CommentsStore); }
  private get folderTasks(): FolderTasksStore { return this.injector.get(FolderTasksStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get attachments(): AttachmentsStore { return this.injector.get(AttachmentsStore); }
  private get hr(): HrFilesStore { return this.injector.get(HrFilesStore); }

  selectedFolderName = '';

  selectedAllFilesFolderId: string | null = null;

  /** Parent folder name when a level-2 ChildSubject folder is selected (for task-cache reuse). */
  selectedAllFilesParentFolderName = '';

  childSubjectSourceFetchedForFolderId: string | null = null;

  /** ChildSubject / ChildSubject2 names keyed by library + trailing eForm id. */
  readonly childSubjectNamesByLibraryEFormId = new Map<string, string[]>();

  /** SharePoint lastModifiedBy for the folder selected in All Files (shown as Submitted by). */
  selectedFolderModifiedBy = '';

  /** Library name ? task list name, sourced from the eForms list via FormConfigurationService. */
  documentLibraryTaskMap: Record<string, string> = {};

  documentLibraryTaskMapSub?: Subscription;

  /**
   * Unblock the ChildSubject accordion without waiting for the full folder task scan.
   * Procurement ChildSubject columns live on the source eForm item (trailing folder id).
   */
  async primeChildSubjectForFolder(
    libraryName: string,
    folderName: string,
    folderId: string,
  ): Promise<void> {
    if (!folderId || !this.ui.showAllFilesSection) return;

    const eFormId = extractTrailingFolderId(folderName);
    if (!eFormId) return;

    const cacheKey = childSubjectCacheKey(libraryName, eFormId);
    const cached = this.childSubjectNamesByLibraryEFormId.get(cacheKey);
    if (cached && cached.length > 0) {
      if (this.selectedAllFilesFolderId === folderId) {
        this.childSubjectSourceFetchedForFolderId = folderId;
        this.ui.allFilesComponent?.setChildSubjectFolderNames(folderId, cached);
      }
      return;
    }

    const names = await this.loadChildSubjectNamesFromSourceEFormDirect(libraryName, folderName, eFormId);
    if (this.selectedAllFilesFolderId !== folderId) return;

    if (names.length > 0) {
      this.childSubjectSourceFetchedForFolderId = folderId;
      this.childSubjectNamesByLibraryEFormId.set(cacheKey, names);
      this.ui.allFilesComponent?.setChildSubjectFolderNames(folderId, names);
    }
    // Empty / failed: leave pending for publishSelectedFolderChildSubjects (task path).
  }

  /** Populate All Files level 2 from ChildSubject + ChildSubject2 on the source eForm row. */
  async publishSelectedFolderChildSubjects(items: any[], done: boolean): Promise<void> {
    const folderId = this.selectedAllFilesFolderId;
    if (!folderId || !this.ui.showAllFilesSection) return;

    let names = this.extractChildSubjectNamesFromTaskItems(items);
    if (names.length === 0 && this.childSubjectSourceFetchedForFolderId !== folderId && items.length > 0) {
      names = await this.loadChildSubjectNamesFromSourceEForm(items);
      this.childSubjectSourceFetchedForFolderId = folderId;
      const libraryName = this.attachments.currentLibraryName ?? '';
      const eFormId = extractTrailingFolderId(this.selectedFolderName);
      if (names.length > 0 && libraryName && eFormId) {
        this.childSubjectNamesByLibraryEFormId.set(childSubjectCacheKey(libraryName, eFormId), names);
      }
    }

    if (names.length > 0) {
      this.ui.allFilesComponent?.setChildSubjectFolderNames(folderId, names);
      return;
    }

    // Don't wipe an early primeChildSubjectForFolder result when tasks have no ChildSubject columns.
    if (done && this.childSubjectSourceFetchedForFolderId !== folderId) {
      this.ui.allFilesComponent?.setChildSubjectFolderNames(folderId, names);
    }
  }

  extractChildSubjectNamesFromTaskItems(items: any[]): string[] {
    const names: string[] = [];
    const seen = new Set<string>();
    const selectedName = normalizeName(this.selectedFolderName);

    for (const item of items) {
      const fields = item?.eFormDetails?.rawFields ?? {};
      names.push(...extractChildSubjectNamesFromFields(fields, selectedName, seen));
    }
    return names;
  }

  /** ChildSubject columns live on the source eForm list item, not always on the task row. */
  async loadChildSubjectNamesFromSourceEForm(items: any[]): Promise<string[]> {
    const folderEFormId = extractTrailingFolderId(this.selectedFolderName);
    const seedTask =
      items.find((item) => String(item?.eFormDetails?.eFormListId ?? '').trim() === folderEFormId) ??
      items[0];
    if (!seedTask) return [];

    const eFormListId = String(seedTask?.eFormDetails?.eFormListId ?? folderEFormId ?? '').trim();
    if (!eFormListId) return [];

    const libraryName = this.attachments.currentLibraryName ?? '';
    const taskListName = String(seedTask?.eFormDetails?.listName ?? '').trim();
    return this.loadChildSubjectNamesFromSourceEFormDirect(libraryName, this.selectedFolderName, eFormListId, taskListName);
  }

  /** Resolve ChildSubject names from the source eForm item by trailing folder / eForm id. */
  async loadChildSubjectNamesFromSourceEFormDirect(
    libraryName: string,
    folderName: string,
    eFormListId: string,
    taskListName = '',
  ): Promise<string[]> {
    if (!eFormListId) return [];

    const form =
      this.formConfigService.getFormByDocLibrary(libraryName) ??
      (taskListName
        ? this.formConfigService.snapshot.find(
          (entry) => String(entry.taskListName ?? '').trim().toLowerCase() === taskListName.toLowerCase()
        )
        : undefined);
    if (!form) return [];

    try {
      const token = await this.session.getSharePointToken();
      await this.session.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );
      if (!this.session.cachedSiteId) return [];

      const sourceListName = this.parseSharePointListNameFromUrl(form.url) || String(form.prefix ?? '').trim();
      const sourceList = this.session.cachedSiteLists.find((list: any) =>
        (list.name ?? '').toLowerCase() === sourceListName.toLowerCase() ||
        (list.displayName ?? '').toLowerCase() === sourceListName.toLowerCase()
      );
      if (!sourceList?.id) return [];

      const response: any = await graphGetWithRetry(
        this.http,
        `/sites/${this.session.cachedSiteId}/lists/${sourceList.id}/items/${eFormListId}?$expand=fields`,
        token,
        AppConstants.graphFileListingTimeoutMs,
      );
      const fields = (response?.fields ?? {}) as Record<string, unknown>;
      return extractChildSubjectNamesFromFields(fields, normalizeName(folderName), new Set());
    } catch (err) {
      console.warn('[All Files] Could not load ChildSubject columns from source eForm item.', err);
      return [];
    }
  }

  parseSharePointListNameFromUrl(url: string): string {
    return this.formConfigService.parseSharePointListNameFromUrl(url);
  }

  toggleAllFilesSection(): void {
    this.attachments.closePreview();
    this.todo.cancelBackgroundTodoLoad();
    this.folderTasks.cancelPendingFolderTaskLoads();
    this.folderTasks.selectedHrPersonalTaskFolder = '';
    this.ui.allFilesMounted = true;
    this.ui.showAllFilesSection = true;
    this.ui.showToDoSection = false;
    this.ui.hideComments = true;
    this.todo.selectedSubmitter = null;
    this.todo.selectedEFormListId = null;
    this.todo.selectedEFormTitle = null;
    this.selectedFolderName = '';
    this.attachments.showUserFile = false;
    this.attachments.userFiles = [];
    this.attachments._folderMapDirty = true;
    this.attachments.folderStack = [];
    this.attachments.attachmentBrowsingRoot = null;
    this.attachments.currentFolderId = null;
    this.attachments.currentFolderWebUrl = '';
    this.attachments.currentFolderName = '';
    this.attachments.userRootFolderId = null;
    Object.assign(this.ui, getAllFilesTabMobileState());
    this.comments.invalidateCommentFilters();
    this.ui.syncMobileSearchInput();
    this.view.refreshView();
  }

  /** Resolve the SharePoint tasks list name for a document library. */
  getTaskListForLibrary(libraryName: string): string | undefined {
    const map = this.documentLibraryTaskMap;
    if (!libraryName || !map) return undefined;

    if (map[libraryName]) return map[libraryName];

    // Case-insensitive fallback: DocLibraryName casing in the eForms list may differ
    // from the SharePoint drive (library) name.
    const target = libraryName.toLowerCase();
    const matchKey = Object.keys(map).find((key) => key.toLowerCase() === target);
    return matchKey ? map[matchKey] : undefined;
  }

  /** Called when a file/folder is clicked in the All Files view. */
  onAllFileSelected(file: any): void {
    if (!file) return;
    this.comments.pendingCommentFocusId = null;
    this.comments.highlightedCommentId = null;
    this.attachments.pendingAttachmentFocusId = null;
    this.attachments.highlightedAttachmentId = null;

    // Folder click: load that folder's children into Attachments (not the whole library).
    if (file.isfolder) {
      const driveId = file.driveId ?? this.attachments.currentLibraryDriveId;
      if (driveId) {
        this.selectedAllFilesFolderId = String(file.id ?? '');
        this.selectedAllFilesParentFolderName = String(file.parentFolderName ?? '').trim();
        this.childSubjectSourceFetchedForFolderId = null;
        this.ui.allFilesComponent?.beginChildSubjectLoad(this.selectedAllFilesFolderId);
        this.attachments.pauseFolderPrefetch();
        this.attachments.currentLibraryDriveId = driveId;
        this.attachments.attachmentBrowsingRoot = { id: file.id, name: file.name, webUrl: file.webUrl ?? '' };
        this.attachments.folderStack = [];
        this.attachments.currentFolderName = file.name;
        // Show Attachments immediately - tasks load in the background for Comments.
        this.ui.hideComments = false;
        Object.assign(this.ui, getAttachmentsTabMobileState());
        void this.attachments.loadDriveFolderContents(driveId, file.id, file.webUrl, file.name);

        const libName = file.libraryName ?? this.attachments.currentLibraryName;
        if (libName) {
          this.selectedFolderModifiedBy = String(file.modifiedby ?? '').trim();
          // ChildSubject accordion: fetch source eForm in parallel (don't wait for task scan).
          void this.primeChildSubjectForFolder(libName, String(file.name || ''), this.selectedAllFilesFolderId);
          this.folderTasks.loadAllFilesTasksForLibrary(libName, String(file.name || ''));
        }

        return;
      }
      this.attachments.userFiles = [{
        id: file.id,
        name: file.name,
        webUrl: file.webUrl ?? '',
        parentId: file.parentId ?? undefined,
        isFolder: true,
        modifiedBy: file.modifiedby,
        fileExtension: file.fileExtension ?? undefined,
        mimeType: '',
        fileCategory: file.fileCategory ?? undefined,
        fileIcon: file.fileIcon ?? undefined,
      }];
      this.attachments._folderMapDirty = true;
      this.attachments.showUserFile = true;
      this.view.refreshView();
      return;
    }

    // Non-folder: show single file entry in attachments panel
    this.attachments.userFiles = [{
      id: file.id,
      name: file.name,
      webUrl: file.webUrl ?? '',
      parentId: file.parentId ?? undefined,
      lastModifiedDateTime: file.modified ?? file.lastModifiedDateTime,
      size: file.size,
      isFolder: !!file.isfolder,
      modifiedBy: file.modifiedby,
      fileExtension: file.fileExtension ?? undefined,
      mimeType: '',
      fileCategory: file.fileCategory ?? undefined,
      fileIcon: file.fileIcon ?? undefined,

    }];
    this.attachments._folderMapDirty = true;
    this.attachments.showUserFile = true;
    this.view.refreshView();
  }

  /**
   * All Files → Comments search row: open the related folder and focus that
   * comment/task in the Comments panel (Live behaviour).
   */
  onAllFilesCommentSelected(hit: CommentSearchHit): void {
    if (!hit?.folder) return;

    const file = hit.folder;
    const taskId = String(hit.taskId ?? '').trim();
    const driveId = file.driveId ?? this.attachments.currentLibraryDriveId;
    if (!driveId) return;

    const libName = String(file.libraryName ?? this.attachments.currentLibraryName ?? '').trim();
    const folderId = String(file.id ?? '');
    const folderName = String(file.name || '');
    const folderWebUrl = file.webUrl ?? '';

    // Stop background Comments crawling so the click isn't competing with Graph.
    this.ui.allFilesComponent?.pauseCommentTasksWarm();

    this.comments.pendingCommentFocusId = taskId || null;
    this.comments.highlightedCommentId = taskId || null;
    this.attachments.pendingAttachmentFocusId = null;
    this.attachments.highlightedAttachmentId = null;
    this.selectedAllFilesFolderId = folderId;
    this.selectedAllFilesParentFolderName = String(file.parentFolderName ?? '').trim();
    this.childSubjectSourceFetchedForFolderId = null;
    this.attachments.currentLibraryDriveId = driveId;
    this.attachments.attachmentBrowsingRoot = { id: file.id, name: file.name, webUrl: folderWebUrl };
    this.attachments.folderStack = [];
    this.attachments.currentFolderName = folderName;
    this.selectedFolderModifiedBy = String(file.modifiedby ?? '').trim();
    this.ui.hideComments = false;
    Object.assign(this.ui, getCommentsTabMobileState());

    // Paint only the clicked comment immediately — never scan the full search index here.
    this.seedCommentsFromSearchHit(hit);
    this.comments.highlightedCommentId = taskId || null;
    this.comments.pendingCommentFocusId = taskId || null;
    this.comments.focusPendingCommentIfNeeded();
    // Keep focus armed for when the full folder Comments list replaces the seed.
    this.comments.pendingCommentFocusId = taskId || null;

    // Defer Attachments + full folder Comments so the UI stays responsive on click.
    setTimeout(() => {
      this.attachments.pauseFolderPrefetch();
      this.ui.allFilesComponent?.beginChildSubjectLoad(folderId);
      void this.attachments.loadDriveFolderContents(driveId, folderId, folderWebUrl, folderName);
      if (libName) {
        void this.primeChildSubjectForFolder(libName, folderName, folderId);
        this.folderTasks.loadAllFilesTasksForLibrary(libName, folderName);
      }
      // Resume indexing after the click-driven Graph work has had time to start.
      setTimeout(() => this.ui.allFilesComponent?.resumeCommentTasksWarm(), 2500);
    }, 0);
  }

  /**
   * All Files → Attachments search row: open the related folder and highlight
   * that file in the Attachments panel (Live behaviour).
   */
  onAllFilesAttachmentSelected(hit: AttachmentSearchHit): void {
    if (!hit?.folder) return;

    const file = hit.folder;
    const fileId = String(hit.fileId ?? '').trim();
    const driveId = file.driveId ?? this.attachments.currentLibraryDriveId;
    if (!driveId) return;

    const libName = String(file.libraryName ?? this.attachments.currentLibraryName ?? '').trim();
    const folderId = String(file.id ?? '');
    const folderName = String(file.name || '');
    const folderWebUrl = file.webUrl ?? '';

    this.comments.pendingCommentFocusId = null;
    this.comments.highlightedCommentId = null;
    this.attachments.pendingAttachmentFocusId = fileId || null;
    this.attachments.highlightedAttachmentId = fileId || null;
    this.selectedAllFilesFolderId = folderId;
    this.selectedAllFilesParentFolderName = String(file.parentFolderName ?? '').trim();
    this.childSubjectSourceFetchedForFolderId = null;
    this.attachments.currentLibraryDriveId = driveId;
    this.attachments.attachmentBrowsingRoot = { id: file.id, name: file.name, webUrl: folderWebUrl };
    this.attachments.folderStack = [];
    this.attachments.currentFolderName = folderName;
    this.selectedFolderModifiedBy = String(file.modifiedby ?? hit.author ?? '').trim();
    this.ui.hideComments = false;
    Object.assign(this.ui, getAttachmentsTabMobileState());

    // Instant paint: show the clicked file while the full folder listing loads.
    this.attachments.userFiles = [{
      id: hit.fileId,
      name: hit.fileName,
      webUrl: hit.webUrl ?? '',
      lastModifiedDateTime: hit.date || undefined,
      isFolder: false,
      modifiedBy: hit.author || undefined,
      fileIcon: hit.fileIcon || undefined,
    }];
    this.attachments._folderMapDirty = true;
    this.attachments.showUserFile = true;
    this.attachments.isLoadingUserFiles = false;
    this.view.refreshView();

    setTimeout(() => {
      this.attachments.pauseFolderPrefetch();
      this.ui.allFilesComponent?.beginChildSubjectLoad(folderId);
      void this.attachments.loadDriveFolderContents(driveId, folderId, folderWebUrl, folderName).then(() => {
        this.attachments.focusPendingAttachmentIfNeeded();
      });
      if (libName) {
        void this.primeChildSubjectForFolder(libName, folderName, folderId);
        this.folderTasks.loadAllFilesTasksForLibrary(libName, folderName);
      }
    }, 0);
  }

  /** Instant Comments paint: only the clicked search hit (full folder load follows async). */
  seedCommentsFromSearchHit(hit: CommentSearchHit): void {
    const folderName = String(hit.folder?.name ?? '').trim();
    if (!folderName) return;

    this.todo.selectedSubmitter = null;
    this.todo.selectedEFormListId = null;
    this.todo.selectedEFormTitle = null;
    this.selectedFolderName = folderName;
    this.ui.hideComments = false;

    const libName = String(hit.folder?.libraryName ?? hit.libraryName ?? '').trim();
    const listName = (libName && this.getTaskListForLibrary(libName)) || libName || 'Tasks';
    const item = buildQuickCommentItemFromSearchHit(hit, listName);

    this.comments.isLoadingComments = false;
    this.comments.isLoadingMoreComments = false;
    this.comments.commentsMessage = '';
    this.comments.commentItems = [item];
    this.comments.visibleItemCount = this.comments.commentsInitialPageSize;
    this.comments.invalidateCommentFilters(true);
    this.view.refreshView();
  }

  /**
   * Quietly cache Comments for an All Files folder during Comments search.
   * Does not change the open Comments panel — only warms tasks:lib/hr caches.
   */
  onEnsureFolderComments(req: { libraryName: string; folderName: string }): void {
    const libraryName = String(req?.libraryName ?? '').trim();
    const folderName = String(req?.folderName ?? '').trim();
    if (!libraryName || !folderName) return;
    void this.silentWarmFolderCommentsForSearch(libraryName, folderName);
  }

  /**
   * Quietly cache Attachments for an All Files folder during Attachments search.
   * Does not change the open Attachments panel — only warms files:folder caches.
   */
  onEnsureFolderAttachments(req: { driveId: string; folderId: string }): void {
    const driveId = String(req?.driveId ?? '').trim();
    const folderId = String(req?.folderId ?? '').trim();
    if (!driveId || !folderId) return;
    void this.silentWarmFolderAttachmentsForSearch(driveId, folderId);
  }

  async silentWarmFolderAttachmentsForSearch(
    driveId: string,
    folderId: string,
  ): Promise<void> {
    const cacheKey = driveFolderCacheKey(driveId, folderId);
    try {
      const existing = this.fileCrawlCache.getStale(cacheKey);
      if (Array.isArray(existing)) {
        this.ui.allFilesComponent?.refreshAttachmentSearchFromCaches();
        return;
      }

      const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
      if (hydrated && Array.isArray(hydrated.data)) {
        this.fileCrawlCache.set(cacheKey, hydrated.data);
        this.ui.allFilesComponent?.refreshAttachmentSearchFromCaches();
        if (hydrated.fresh) return;
      }

      if (this.attachments.folderPrefetchInFlight.has(cacheKey)) return;
      this.attachments.folderPrefetchInFlight.add(cacheKey);
      try {
        const token = await this.session.getSharePointToken();
        const mapped = await this.driveFiles.fetchFolderChildren(driveId, folderId, token);
        this.fileCrawlCache.set(cacheKey, mapped);
        this.ui.allFilesComponent?.refreshAttachmentSearchFromCaches();
      } finally {
        this.attachments.folderPrefetchInFlight.delete(cacheKey);
      }
    } catch (err) {
      console.warn(`[Attachments search] Failed to warm folder ${folderId}`, err);
    }
  }

  async silentWarmFolderCommentsForSearch(
    libraryName: string,
    folderName: string,
  ): Promise<void> {
    try {
      if (normalizeName(libraryName) === normalizeName(this.hr.targetLibraryName)) {
        const cacheKey = hrFolderTaskCacheKey(folderName);
        const existing = this.fileCrawlCache.getStale(cacheKey);
        if (Array.isArray(existing) && existing.length > 0) {
          this.ui.allFilesComponent?.refreshCommentSearchFromCaches();
          return;
        }
        const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
        if (hydrated && Array.isArray(hydrated.data)) {
          this.fileCrawlCache.set(cacheKey, hydrated.data);
          this.ui.allFilesComponent?.refreshCommentSearchFromCaches();
        }
        // Skip Graph for HR during search warm — opening the folder loads Comments normally.
        return;
      }

      const mappedTaskList = this.getTaskListForLibrary(libraryName);
      if (!mappedTaskList) return;

      const cacheKey = libFolderTaskCacheKey(mappedTaskList, folderName);
      const existing = this.fileCrawlCache.getStale(cacheKey);
      if (Array.isArray(existing) && existing.length > 0) {
        this.ui.allFilesComponent?.refreshCommentSearchFromCaches();
        return;
      }

      const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
      if (hydrated && Array.isArray(hydrated.data)) {
        this.fileCrawlCache.set(cacheKey, hydrated.data);
        this.ui.allFilesComponent?.refreshCommentSearchFromCaches();
        if (hydrated.fresh) return;
      }

      await this.folderTasks.loadDocumentLibraryTasksForFolder(
        mappedTaskList,
        folderName,
        this.folderTasks.allFilesFolderTaskLoadSeq,
        { silent: true },
      );
    } catch (err) {
      console.warn(`[Comments search] Failed to warm "${folderName}"`, err);
    }
  }

  /** Remember which library is active in All Files - do not preload attachments. */
  onLibrarySelected(event: { name: string; driveId: string } | null): void {
    if (!event) return;
    const { name: libraryName, driveId } = event;

    this.attachments.currentLibraryName = libraryName;
    this.attachments.currentLibraryDriveId = driveId;
    this.attachments.folderStack = [];
    this.attachments.attachmentBrowsingRoot = null;
    this.attachments.currentFolderId = null;
    this.attachments.currentFolderWebUrl = '';
    this.attachments.currentFolderName = '';
    this.attachments.userRootFolderId = null;
    this.attachments.showUserFile = false;
    this.attachments.userFiles = [];
    this.attachments._folderMapDirty = true;

    // Tasks and attachments load only when a folder is clicked
    this.selectedFolderName = '';
    this.selectedFolderModifiedBy = '';
    this.comments.commentItems = [];
    this.comments.commentsMessage = '';
    this.ui.hideComments = true;
    this.comments.invalidateCommentFilters();
    this.view.refreshView();
  }

  refreshAllFilesSection(): void {
    // Only the open folder is re-read. Folder listings and Comments are shared in Redis,
    // so nothing is deleted up front: each re-fetch overwrites its own key in place.
    void this.ui.allFilesComponent?.refreshAllFilesData(true);

    if (this.selectedFolderName && this.attachments.currentLibraryName) {
      if (normalizeName(this.attachments.currentLibraryName) === normalizeName(this.hr.targetLibraryName)) {
        // HR person folders merge on soft refresh, so keep the clean re-crawl here.
        this.folderTasks.invalidateFolderTaskCache(this.attachments.currentLibraryName, this.selectedFolderName);
        this.folderTasks.loadAllFilesTasksForLibrary(this.attachments.currentLibraryName, this.selectedFolderName);
      } else {
        this.folderTasks.softRefreshOpenFolderTaskCache();
      }
    }

    const stackFolder = this.attachments.folderStack[this.attachments.folderStack.length - 1];
    const folderId =
      this.attachments.currentFolderId ||
      this.attachments.attachmentBrowsingRoot?.id ||
      stackFolder?.id ||
      '';
    const folderWebUrl =
      this.attachments.currentFolderWebUrl ||
      this.attachments.attachmentBrowsingRoot?.webUrl ||
      stackFolder?.webUrl ||
      '';
    const folderName =
      this.attachments.currentFolderName ||
      this.attachments.attachmentBrowsingRoot?.name ||
      stackFolder?.name ||
      '';
    if (this.attachments.currentLibraryDriveId && folderId) {
      void this.attachments.loadDriveFolderContents(
        this.attachments.currentLibraryDriveId,
        folderId,
        folderWebUrl,
        folderName,
        { skipCache: true },
      );
    }

    this.view.refreshView();
  }

  /** Leave HR Files tab context - drop All Files navigation without losing hrPersonalDriveId. */
  leaveAllFilesTaskView(): void {
    if (
      !this.hr.hrPersonalDriveId &&
      this.attachments.currentLibraryDriveId &&
      normalizeName(this.attachments.currentLibraryName ?? '') === normalizeName(this.hr.targetLibraryName)
    ) {
      this.hr.hrPersonalDriveId = this.attachments.currentLibraryDriveId;
    }

    this.folderTasks.cancelPendingFolderTaskLoads();
    this.attachments.currentLibraryDriveId = null;
    this.attachments.currentLibraryName = null;
    this.attachments.currentFolderId = null;
    this.attachments.currentFolderWebUrl = '';
    this.attachments.currentFolderName = '';
    this.attachments.folderStack = [];
    this.attachments.attachmentBrowsingRoot = null;
    this.selectedFolderName = '';
    this.selectedFolderModifiedBy = '';
    this.folderTasks.selectedHrPersonalTaskFolder = '';
    this.attachments.showUserFile = false;
    this.attachments.userFiles = [];
    this.attachments._folderMapDirty = true;
    this.comments.commentItems = [];
    this.comments.commentsMessage = '';
    this.comments.isLoadingComments = false;
    this.comments.isLoadingMoreComments = false;
    this.comments.invalidateCommentFilters();
  }
}
