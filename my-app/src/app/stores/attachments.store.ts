import { Injectable, Injector, inject, signal } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { AppConstants } from '../app.constants';
import { graphGetWithRetry, toGraphPath, normalizeName } from '../microsoft-graph';
import { FileCrawlCacheService } from '../services/file-crawl.service';
import { SiteMetadataService } from '../services/site-metadata.service';
import { AppDropdownOption } from '../components/app-dropdown/app-dropdown.component';
import { DriveFilesService, DriveFile } from '../services/drive-files.service';
import { extractTrailingFolderId } from '../utils/hr-task-matching';
import { driveFolderCacheKey } from '../utils/cache-keys';
import { ViewRefreshService } from './view-refresh.service';
import { UiStore } from './ui.store';
import { SessionStore } from './session.store';
import { FolderTasksStore } from './folder-tasks.store';
import { TodoStore } from './todo.store';
import { HrFilesStore } from './hr-files.store';
import { AllFilesStore } from './all-files.store';

/**
 * The Attachments panel: the files of the open folder/task, folder navigation (open / Back), search, the file preview and the HR task-list selector.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class AttachmentsStore {
  private readonly injector = inject(Injector);
  private readonly http = inject(HttpClient);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly siteMetadataService = inject(SiteMetadataService);
  private readonly driveFiles = inject(DriveFilesService);

  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get folderTasks(): FolderTasksStore { return this.injector.get(FolderTasksStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get hr(): HrFilesStore { return this.injector.get(HrFilesStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  showUserFile = false;

  isLoadingUserFiles = false;

  userFileProgressMessage = '';

  userFileError = '';

  userFileWarning = '';

  currentFolderId: string | null = null;

  currentFolderWebUrl = '';

  currentFolderName = '';

  folderStack: Array<{ id: string; name: string; webUrl: string }> = [];

  /** Top-level folder opened from HR Files / All Files - back never goes above this. */
  attachmentBrowsingRoot: { id: string; name: string; webUrl: string } | null = null;

  currentLibraryDriveId: string | null = null;

  currentLibraryName: string | null = null;

  /** After Attachments-search click: highlight this file id in the Attachments panel. */
  pendingAttachmentFocusId: string | null = null;

  highlightedAttachmentId: string | null = null;

  allowedHrTaskLists: string[] = [];

  selectedHrTaskList: string = '';

  get isUsingHrTaskList(): boolean { return this.selectedHrTaskList !== ''; }

  /** File shown in the preview box (set when a row is hovered or clicked). */
  previewFile = signal<DriveFile | null>(null);

  /**
   * Rows shown in the Attachments card (the type lives in DriveFilesService).
   * Written through a setter so that whenever the list is replaced - new folder,
   * Back, a new task/file, another section - a preview of a file that is no longer
   * listed closes by itself. A refresh of the same folder keeps the preview open.
   */
  _userFiles: DriveFile[] = [];

  get userFiles(): DriveFile[] { return this._userFiles; }

  set userFiles(files: DriveFile[]) {
    this._userFiles = files;
    const shown = this.previewFile();
    if (shown && !files.some(f => f.id === shown.id)) this.previewFile.set(null);
  }

  /** Hide the preview (Back, opening a folder, switching section or tab). */
  closePreview(): void {
    this.previewFile.set(null);
  }

  /** Prevent stale folder results when users click folders quickly. */
  driveFolderLoadSeq = 0;

  /** Guards To Do attachment lookups so a slow lookup cannot paint over a newer card click. */
  todoAttachmentLoadSeq = 0;

  /** Document library name (normalized) -> drive id. */
  libraryDriveIdCache = new Map<string, string | null>();

  /** `${driveId}|${folderName}` -> eForm root folder. */
  todoTaskFolderCache = new Map<string, { id: string; name: string; webUrl: string }>();

  /** In-flight silent prefetches keyed by `files:folder:{driveId}:{folderId}`. */
  readonly folderPrefetchInFlight = new Set<string>();

  folderPrefetchTimer: ReturnType<typeof setTimeout> | null = null;

  /** When true, idle prefetch is suspended so user clicks / task loads get Graph capacity. */
  folderPrefetchPaused = false;

  userRootFolderId: string | null = null;

  /** Item count shown in the attachments header - matches what filteredAttachments displays. */
  get loadedHrPersonalItemCount(): number {
    if (this.currentLibraryDriveId || this.isUsingHrTaskList) {
      return this.filteredAttachments.length;
    }
    return this.visibleFiles.length;
  }

  readonly hrTaskListOptions: AppDropdownOption[] = [
    { value: '', label: '-- Select a list --' },
    { value: 'ProcEForm', label: 'ProcEForm' },
  ];

  // Alias used by the template
  get attachmentItems() { return this.userFiles; }

  // FOLDER MAP - cached, rebuilt only when userFiles changes
  _folderMap = new Map<string, typeof this.userFiles>();

  _folderMapDirty = true;

  refreshAttachmentFolderMap(): void {
    if (!this._folderMapDirty) return;
    this._folderMap.clear();
    for (const f of this.userFiles) {
      const key = f.parentId ?? '';
      if (!this._folderMap.has(key)) this._folderMap.set(key, []);
      this._folderMap.get(key)!.push(f);
    }
    this._folderMapDirty = false;
  }

  get visibleFiles() {
    this.refreshAttachmentFolderMap();
    const key = this.currentFolderId ?? this.userRootFolderId ?? '';
    return this._folderMap.get(key) ?? [];
  }

  /** Show Back while browsing an opened folder tree. */
  get showBackButtonInAttachments(): boolean {
    return this.showUserFile && !!this.attachmentBrowsingRoot;
  }

  /** True when Back can move up to a parent folder within the opened tree. */
  get canGoBackInAttachments(): boolean {
    return this.folderStack.length > 0;
  }

  get hasLoadedPersonalFiles(): boolean {
    return !this.isUsingHrTaskList && this.userRootFolderId !== null;
  }

  userFileLoadingTimeoutId: ReturnType<typeof setTimeout> | null = null;

  attachmentSearch = '';

  get filteredAttachments() {
    // For HR Task lists (All Files), show all userFiles directly
    if (this.isUsingHrTaskList) return this.userFiles;

    // Library browsing loads one folder level at a time via Graph - userFiles is already scoped
    const source = this.currentLibraryDriveId ? this.userFiles : this.visibleFiles;
    return this.applyAttachmentSearch(source);
  }

  applyAttachmentSearch(source: typeof this.userFiles) {
    if (!this.attachmentSearch.trim()) return source;
    const query = this.attachmentSearch.toUpperCase().trim();
    return source.filter(item => {
      if (item.isFolder) return false;
      if (item.name?.toUpperCase().includes(query.toUpperCase())) return true;
      if (item.modifiedBy?.toUpperCase().includes(query.toUpperCase())) return true;
      return false;
    });
  }

  onAttachmentSearchChange(event: Event): void {
    this.attachmentSearch = (event.target as HTMLInputElement).value;
    this.view.refreshView();
  }

  /**
   * To Do: show the eForm folder for the clicked task in Attachments, same as
   * clicking that folder in All Files. The folder is a root folder of the document
   * library linked to the task list (eForms config), named after the task Title
   * ("Pr-Test-21" = title + eFormListId).
   */
  async loadTodoTaskAttachments(task: any, eFormTitle: string): Promise<void> {
    const seq = ++this.todoAttachmentLoadSeq;
    const isStale = () => seq !== this.todoAttachmentLoadSeq || !this.ui.showToDoSection;

    const listName = String(task?.eFormDetails?.listName ?? task?.listName ?? '').trim();
    const rawFields = (task?.eFormDetails?.rawFields ?? {}) as Record<string, unknown>;
    const folderName = String(rawFields['Title'] ?? eFormTitle ?? '').trim();
    const eFormListId = String(task?.eFormDetails?.eFormListId ?? '').trim();
    const libraryName = this.getLibraryForTaskList(listName);

    this.userFileError = '';
    this.userFileWarning = '';
    this.highlightedAttachmentId = null;
    this.pendingAttachmentFocusId = null;

    if (!libraryName || !folderName) {
      this.showTodoAttachmentsMessage('No attachments folder is linked to this task.');
      return;
    }

    this.userFiles = [];
    this._folderMapDirty = true;
    this.folderStack = [];
    this.currentFolderId = null;
    this.currentFolderName = folderName;
    this.showUserFile = true;
    this.isLoadingUserFiles = true;
    this.userFileProgressMessage = 'Loading attachments...';
    this.view.refreshView();

    try {
      const token = await this.session.getSharePointToken();
      if (isStale()) return;
      const driveId = await this.resolveLibraryDriveId(libraryName, token);
      if (isStale()) return;
      if (!driveId) {
        this.showTodoAttachmentsMessage(`Could not find the "${libraryName}" document library.`);
        return;
      }

      const folder = await this.findTodoTaskFolder(driveId, folderName, eFormListId, token);
      if (isStale()) return;
      if (!folder) {
        this.showTodoAttachmentsMessage(`No attachments folder "${folderName}" found in ${libraryName}.`);
        return;
      }

      this.currentLibraryDriveId = driveId;
      this.attachmentBrowsingRoot = { ...folder };
      this.folderStack = [];
      await this.loadDriveFolderContents(driveId, folder.id, folder.webUrl, folder.name);
    } catch (err: any) {
      if (isStale()) return;
      console.warn('[To Do] Failed to load task attachments', err);
      this.showTodoAttachmentsMessage(err?.message || 'Failed to load attachments.');
    }
  }

  showTodoAttachmentsMessage(message: string): void {
    this.userFiles = [];
    this._folderMapDirty = true;
    this.currentFolderId = null;
    this.attachmentBrowsingRoot = null;
    this.folderStack = [];
    this.showUserFile = true;
    this.isLoadingUserFiles = false;
    this.userFileProgressMessage = '';
    this.userFileWarning = message;
    this.view.refreshView();
  }

  /** Reverse of getTaskListForLibrary: task list name -> document library name. */
  getLibraryForTaskList(listName: string): string | undefined {
    const target = normalizeName(listName);
    if (!target) return undefined;
    return Object.keys(this.allFiles.documentLibraryTaskMap).find(
      (library) => normalizeName(this.allFiles.documentLibraryTaskMap[library]) === target,
    );
  }

  async resolveLibraryDriveId(libraryName: string, token: string): Promise<string | null> {
    const key = normalizeName(libraryName);
    if (this.libraryDriveIdCache.has(key)) return this.libraryDriveIdCache.get(key) ?? null;

    const site = await this.siteMetadataService.resolve(token);
    const drives: any = await graphGetWithRetry(
      this.http,
      `/sites/${site.siteId}/drives?$select=id,name,webUrl`,
      token,
    );
    for (const drive of drives?.value ?? []) {
      if (!drive?.id) continue;
      this.libraryDriveIdCache.set(normalizeName(drive.name), drive.id);
      // Drive name can differ from the library URL segment (e.g. spaces removed).
      const urlName = decodeURIComponent(String(drive.webUrl ?? '').split('/').pop() ?? '');
      if (urlName && !this.libraryDriveIdCache.has(normalizeName(urlName))) {
        this.libraryDriveIdCache.set(normalizeName(urlName), drive.id);
      }
    }
    if (!this.libraryDriveIdCache.has(key)) this.libraryDriveIdCache.set(key, null);
    return this.libraryDriveIdCache.get(key) ?? null;
  }

  /** Exact root-folder lookup by Title; falls back to a root folder ending in "-<eFormListId>". */
  async findTodoTaskFolder(
    driveId: string,
    folderName: string,
    eFormListId: string,
    token: string,
  ): Promise<{ id: string; name: string; webUrl: string } | null> {
    const cacheKey = `${driveId}|${folderName.toLowerCase()}`;
    const cached = this.todoTaskFolderCache.get(cacheKey);
    if (cached) return cached;

    let found: { id: string; name: string; webUrl: string } | null = null;
    try {
      const encoded = folderName.split('/').map(encodeURIComponent).join('/');
      const item: any = await graphGetWithRetry(
        this.http,
        `/drives/${driveId}/root:/${encoded}?$select=id,name,webUrl,folder`,
        token,
        AppConstants.graphFileListingTimeoutMs,
      );
      if (item?.id && item?.folder) found = { id: item.id, name: item.name, webUrl: item.webUrl ?? '' };
    } catch { /* not found by exact name — scan root folders below */ }

    if (!found && eFormListId) {
      let nextPath: string | null =
        `/drives/${driveId}/root/children?$select=id,name,webUrl,folder&$top=999`;
      while (nextPath && !found) {
        const page: any = await graphGetWithRetry(
          this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs,
        );
        const match = (page?.value ?? []).find(
          (f: any) => !!f?.folder && extractTrailingFolderId(String(f.name ?? '')) === eFormListId,
        );
        if (match) found = { id: match.id, name: match.name, webUrl: match.webUrl ?? '' };
        nextPath = toGraphPath(page?.['@odata.nextLink']);
      }
    }

    // Only remember hits, so a folder created after the first click is still found later.
    if (found) this.todoTaskFolderCache.set(cacheKey, found);
    return found;
  }

  focusPendingAttachmentIfNeeded(): void {
    const id = String(this.pendingAttachmentFocusId ?? '').trim();
    if (!id) return;
    this.highlightedAttachmentId = id;
    this.pendingAttachmentFocusId = null;
    this.view.refreshView();
    setTimeout(() => {
      const el = document.getElementById(`attachment-item-${id}`);
      el?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 50);
  }

  /** Load one-level children of a drive folder and show them in Attachments. */
  async loadDriveFolderContents(
    driveId: string,
    folderId: string,
    folderWebUrl?: string,
    folderName?: string,
    options: { skipCache?: boolean } = {},
  ): Promise<void> {
    if (!driveId || !folderId) return;
    const seq = ++this.driveFolderLoadSeq;
    const cacheKey = driveFolderCacheKey(driveId, folderId);

    // Document-library browsing - do not mix with the HR Personal drive tree.
    this.userRootFolderId = null;

    let paintedFromStaleCache = false;
    if (!options.skipCache) {
      const cached = this.fileCrawlCache.get(cacheKey);
      if (cached) {
        if (seq !== this.driveFolderLoadSeq) return;
        this.applyDriveFolderContentsToUi(cached, folderId, folderWebUrl, folderName);
        return;
      }
      // Expired or persisted-from-last-session listing: paint it instantly,
      // then fall through so fresh SharePoint data replaces it below.
      const stale = this.fileCrawlCache.getStale(cacheKey);
      if (stale) {
        if (seq !== this.driveFolderLoadSeq) return;
        this.applyDriveFolderContentsToUi(stale, folderId, folderWebUrl, folderName);
        paintedFromStaleCache = true;
      } else {
        // Same as All Files / HR Comments: shared Redis before spinner + Graph.
        const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
        if (seq !== this.driveFolderLoadSeq) return;
        if (hydrated && Array.isArray(hydrated.data)) {
          this.applyDriveFolderContentsToUi(hydrated.data, folderId, folderWebUrl, folderName);
          paintedFromStaleCache = true;
          if (hydrated.fresh) return;
        }
      }
    }

    // If openHr already painted this folder from cache, never flash loading.
    if (
      !paintedFromStaleCache &&
      this.currentFolderId === folderId &&
      Array.isArray(this.userFiles) &&
      this.userFiles.length > 0 &&
      !this.isLoadingUserFiles
    ) {
      paintedFromStaleCache = true;
    }

    if (!paintedFromStaleCache) {
      this.isLoadingUserFiles = true;
      this.userFileProgressMessage = 'Loading folder contents...';
    }
    this.userFileError = '';
    this.userFileWarning = '';
    this.view.refreshView();
    try {
      const token = await this.session.getSharePointToken();
      const mapped = await this.driveFiles.fetchFolderChildren(driveId, folderId, token);
      if (seq !== this.driveFolderLoadSeq) return;

      this.fileCrawlCache.set(cacheKey, mapped);
      this.applyDriveFolderContentsToUi(mapped, folderId, folderWebUrl, folderName);
    } catch (err: any) {
      if (seq !== this.driveFolderLoadSeq) return;
      this.userFileError = err?.message || 'Failed to load folder contents';
    } finally {
      if (seq !== this.driveFolderLoadSeq) return;
      this.isLoadingUserFiles = false;
      this.view.refreshView();
    }
  }

  applyDriveFolderContentsToUi(
    mapped: any[],
    folderId: string,
    folderWebUrl?: string,
    folderName?: string,
  ): void {
    this.userFiles = mapped;
    this._folderMapDirty = true;
    this.currentFolderId = folderId;
    this.currentFolderWebUrl = folderWebUrl ?? '';
    this.currentFolderName = folderName ?? this.currentFolderName ?? '';
    this.showUserFile = true;
    this.userFileProgressMessage = '';
    this.isLoadingUserFiles = false;
    this.userFileError = '';
    if (folderId === this.allFiles.selectedAllFilesFolderId) {
      this.ui.allFilesComponent?.refreshChildFoldersForFolder(folderId);
    }
    this.ui.allFilesComponent?.refreshAttachmentSearchFromCaches();
    this.focusPendingAttachmentIfNeeded();
    this.view.refreshView();
  }

  pauseFolderPrefetch(): void {
    this.folderPrefetchPaused = true;
    if (this.folderPrefetchTimer) {
      clearTimeout(this.folderPrefetchTimer);
      this.folderPrefetchTimer = null;
    }
    if (this.folderTasks.hrTaskPrefetchTimer) {
      clearTimeout(this.folderTasks.hrTaskPrefetchTimer);
      this.folderTasks.hrTaskPrefetchTimer = null;
    }
    if (this.folderTasks.hrPersonHoverPrefetchTimer) {
      clearTimeout(this.folderTasks.hrPersonHoverPrefetchTimer);
      this.folderTasks.hrPersonHoverPrefetchTimer = null;
    }
    // Resume after click-path work; stay paused while To Do still owns Graph quota.
    setTimeout(() => {
      if (this.todo.isLoadingTodoTasks) return;
      this.folderPrefetchPaused = false;
    }, 8000);
  }

  applyDriveItems(items: any[], rootFolderId: string): void {
    this.userFiles = items
      .filter((item: any) => !!item.file || !!item.folder)
      .map((item: any) => {
        const file = this.driveFiles.toDriveFile(item, rootFolderId);
        // Items directly under the drive root report the root's real id; the folder
        // map looks them up under 'root', so point them there.
        const parentPath = item.parentReference?.path ?? '';
        if (rootFolderId === 'root' && /\/root:?$/i.test(parentPath)) file.parentId = 'root';
        return file;
      });

    this._folderMapDirty = true;
  }

  openFolder(item: { id: string; name: string; webUrl: string }): void {
    this.closePreview();
    this.attachmentSearch = '';

    // All Files / library browsing: fetch subfolder contents from SharePoint
    if (this.currentLibraryDriveId) {
      this.pushCurrentFolderOntoStack();
      void this.loadDriveFolderContents(
        this.currentLibraryDriveId,
        item.id,
        item.webUrl,
        item.name,
      );

      // Doc-library folders map Title ? tasks. HRPersonal subfolders must NOT
      // reload Comments (tasks stay scoped to the person root already open).
      if (this.shouldReloadTasksForAttachmentFolder(item.name)) {
        this.folderTasks.loadAllFilesTasksForLibrary(this.currentLibraryName!, String(item.name || ''));
      }
      return;
    }

    // Personal files: full tree already loaded - filter locally by parentId
    this.pushCurrentFolderOntoStack();
    this.currentFolderId = item.id;
    this.currentFolderWebUrl = item.webUrl;
    this.currentFolderName = item.name;
    this.view.refreshView();
  }

  /** True when drilling into an Attachments subfolder should also refresh Comments tasks. */
  shouldReloadTasksForAttachmentFolder(folderName: string): boolean {
    const libName = String(this.currentLibraryName ?? '').trim();
    if (!libName || !String(folderName ?? '').trim()) return false;
    // HRPersonal person tree: Comments already loaded for the root person.
    if (normalizeName(libName) === normalizeName(this.hr.targetLibraryName)) {
      return false;
    }
    // All Files document libraries: folder name is the task Title key.
    return this.ui.showAllFilesSection && !!this.allFiles.getTaskListForLibrary(libName);
  }

  pushCurrentFolderOntoStack(): void {
    if (!this.currentFolderId) return;
    this.folderStack.push({
      id: this.currentFolderId,
      name: this.currentFolderName || this.allFiles.selectedFolderName || '',
      webUrl: this.currentFolderWebUrl,
    });
  }

  goBackToParent(): void {
    this.closePreview();
    this.navigateToParentFolder();
  }

  navigateToParentFolder(): void {
    this.attachmentSearch = '';

    if (this.currentLibraryDriveId) {
      if (this.folderStack.length > 0) {
        const parent = this.folderStack.pop()!;
        void this.loadDriveFolderContents(
          this.currentLibraryDriveId,
          parent.id,
          parent.webUrl,
          parent.name,
        );
        // Only re-sync Comments for All Files doc-library folders - not HRPersonal.
        if (this.shouldReloadTasksForAttachmentFolder(parent.name)) {
          this.folderTasks.loadAllFilesTasksForLibrary(this.currentLibraryName!, String(parent.name || ''));
        }
        return;
      }

      this.resetToAttachmentBrowsingRoot();
      return;
    }

    if (this.folderStack.length > 0) {
      const parent = this.folderStack.pop()!;
      this.currentFolderId = parent.id;
      this.currentFolderWebUrl = parent.webUrl;
      this.currentFolderName = parent.name;
      this.view.refreshView();
      return;
    }

    if (this.attachmentBrowsingRoot) {
      this.currentFolderId = this.attachmentBrowsingRoot.id;
      this.currentFolderWebUrl = this.attachmentBrowsingRoot.webUrl;
      this.currentFolderName = this.attachmentBrowsingRoot.name;
      this.view.refreshView();
      return;
    }

    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.view.refreshView();
  }

  /** At the entry folder, close attachments; otherwise reload that folder's contents. */
  resetToAttachmentBrowsingRoot(): void {
    if (!this.attachmentBrowsingRoot || !this.currentLibraryDriveId) {
      this.currentFolderId = null;
      this.currentFolderWebUrl = '';
      this.currentFolderName = '';
      this.view.refreshView();
      return;
    }

    if (this.currentFolderId === this.attachmentBrowsingRoot.id) {
      return;
    }

    void this.loadDriveFolderContents(
      this.currentLibraryDriveId,
      this.attachmentBrowsingRoot.id,
      this.attachmentBrowsingRoot.webUrl,
      this.attachmentBrowsingRoot.name,
    );

    // Only re-sync Comments for All Files doc-library folders - not HRPersonal.
    if (this.shouldReloadTasksForAttachmentFolder(this.attachmentBrowsingRoot.name)) {
      this.folderTasks.loadAllFilesTasksForLibrary(
        this.currentLibraryName!,
        String(this.attachmentBrowsingRoot.name || ''),
      );
    }
  }

  onHrTaskListSelected(listName: string): void {
    if (listName && this.allowedHrTaskLists.includes(listName)) {
      this.selectedHrTaskList = listName;
      // Show the file panel when a list is selected
      this.showUserFile = true;
      this.view.refreshView();
    }
  }

  loadHrTaskListItems(): void {
    if (!this.session.currentUser || !this.selectedHrTaskList) {
      this.userFileError = 'Please log in and select an HR Task list first.';
      this.showUserFile = false;
      return;
    }
    if (this.isLoadingUserFiles) return;

    this.showUserFile = true;
    this.isLoadingUserFiles = true;
    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.userFileProgressMessage = `Loading items from ${this.selectedHrTaskList}...`;
    this.userFileError = '';
    this.userFileWarning = '';
    this.userFiles = [];
    this._folderMapDirty = true;
    this.startUserFileLoadingWatchdog();
    this.view.refreshView();

    void this.loadHrTaskListItemsInternal();
  }

  async loadHrTaskListItemsInternal(): Promise<void> {
    try {
      const token = await this.session.getSharePointToken();
      const listName = this.selectedHrTaskList;

      if (!listName || !this.allowedHrTaskLists.includes(listName)) {
        this.userFileError = 'Invalid HR task list selected.';
        this.view.refreshView();
        return;
      }

      const site = await this.siteMetadataService.resolve(token);
      if (!site?.siteId) throw new Error('Could not find site');

      this.userFileProgressMessage = `Finding ${listName} list...`;
      this.view.refreshView();

      // Match by name (more reliable than $filter) against the shared list metadata.
      const allLists = site.lists;
      const targetList = allLists.find((l: any) =>
        (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
        (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
      );

      if (!targetList?.id) {
        this.userFileError = `Could not find list "${listName}". Available lists: ${allLists.map((l: any) => l.name).join(', ')}`;
        this.view.refreshView();
        return;
      }

      this.userFileProgressMessage = `Loading items from ${listName}...`;
      this.view.refreshView();

      // Load items with pagination (hard cap - never drain a multi-thousand-item list here)
      const allItems: any[] = [];
      let nextUrl: string | null = `/sites/${site.siteId}/lists/${targetList.id}/items?$expand=fields&$top=200`;
      let pages = 0;
      const maxPages = 10;

      while (nextUrl && pages < maxPages) {
        this.userFileProgressMessage = `Loading items from ${listName}... (${allItems.length} loaded)`;
        this.view.refreshView();

        const page: any = await graphGetWithRetry(
          this.http, nextUrl, token, AppConstants.graphFileListingTimeoutMs,
        );

        if (page?.value) {
          allItems.push(...page.value);
        }

        nextUrl = toGraphPath(page?.['@odata.nextLink']);
        pages += 1;
      }

      this.userFiles = allItems.map((item: any) => {
        const fields = item.fields || {};
        const title = fields.Title || fields.Title0 || fields.Name || fields.Description || `Item ${item.id}`;
        return {
          id: item.id,
          name: title,
          webUrl: `${site.siteWebUrl}/Lists/${targetList.name}/DispForm.aspx?ID=${item.id}`,
          lastModifiedDateTime: fields.Modified || fields.Created,
          isFolder: false,
        };
      });

      this._folderMapDirty = true;
      this.userFileProgressMessage = '';
      this.userFileError = '';
      if (this.userFiles.length === 0) {
        this.userFileWarning = `No items found in "${listName}".`;
      }
      this.view.refreshView();
    } catch (error: unknown) {
      this.handleUserFilesError(error);
    } finally {
      if (this.isLoadingUserFiles) {
        this.isLoadingUserFiles = false;
        this.clearUserFileLoadingWatchdog();
        this.view.refreshView();
      }
    }
  }

  // the error handler is used to handle the errors that occur when loading the user files
  handleUserFilesError(error: unknown): void {
    this.clearUserFileLoadingWatchdog();
    if ((error as any)?.name === 'TimeoutError') {
      this.userFileError = 'SharePoint request timed out. Please try again.';
    } else if (error instanceof HttpErrorResponse) {
      const msg = error?.error?.error?.message || error?.error?.message || error.message;
      this.userFileError = `Graph error ${error.status}: ${msg}`;
    } else if (error instanceof Error) {
      this.userFileError = error.message;
    } else {
      this.userFileError = 'Failed to load files from SharePoint.';
    }
    this.isLoadingUserFiles = false;
    this.view.refreshView();
  }

  //what this does is it will timeout the user file loading if it takes too long
  startUserFileLoadingWatchdog(): void {
    this.clearUserFileLoadingWatchdog();
    this.userFileLoadingTimeoutId = setTimeout(() => {
      if (this.hr.isLoadingHrFilesList) {
        this.userFileError = 'SharePoint request timed out. Please try again.';
        this.hr.isLoadingHrFilesList = false;
        this.view.refreshView();
        return;
      }
      if (this.isLoadingUserFiles) {
        this.userFileError = 'SharePoint request timed out. Please try again.';
        this.isLoadingUserFiles = false;
        this.view.refreshView();
      }
    }, AppConstants.userFileLoadingWatchdogMs);
  }

  clearUserFileLoadingWatchdog(): void {
    if (this.userFileLoadingTimeoutId !== null) {
      clearTimeout(this.userFileLoadingTimeoutId);
      this.userFileLoadingTimeoutId = null;
    }
  }
}
