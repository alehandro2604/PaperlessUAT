import { Component, OnInit, OnDestroy, ChangeDetectorRef, Output, EventEmitter, Input, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { LoadingScreenComponent } from '../loading screen/loading-screen';
import { HttpClient } from '@angular/common/http';
import { AuthService } from '../../services/auth.service';
import { DocumentService } from '../../services/document.service';
import { sharePointConfig } from '../../sharepoint.config';
import { graphGetWithRetry, toGraphPath, normalizeName } from '../../microsoft-graph';
import { extractTrailingFolderId } from '../../utils/hr-task-matching';
import { AppConstants } from '../../app.constants';
import { getFileExtension, getFileCategory, getFileIcon } from '../../file-utils';
import { FormsModule } from '@angular/forms';
import { FormConfigurationService } from '../../services/form-configuration.service';
import { FileCrawlCacheService } from '../../services/file-crawl.service';
import { AppDropdownComponent, AppDropdownOption } from '../app-dropdown/app-dropdown.component';
import { Subscription } from 'rxjs';

export interface SpFile {
  createdBy?: string;
  /** SharePoint list column "Assigned To" (matches Live folder person). */
  assignedTo?: string;
  id: string;
  name: string;
  webUrl: string;
  isfolder: boolean;
  size?: number;
  modifiedby?: string;
  modified?: string | boolean;
  fileIcon?: string;
  fileCategory?: string;
  fileExtension?: string;
  driveId?: string;
  libraryName?: string;
  parentFolderId?: string;
  parentFolderName?: string;
  parentFolderWebUrl?: string;
}

/** One Comments-search hit: a task/comment row, not a folder. */
export interface CommentSearchHit {
  taskId: string;
  libraryName: string;
  title: string;
  author: string;
  date: string;
  body: string;
  folder: SpFile;
}

/** One Attachments-search hit: Live-style row with folder title + file link. */
export interface AttachmentSearchHit {
  fileId: string;
  fileName: string;
  fileIcon: string;
  webUrl: string;
  libraryName: string;
  title: string;
  author: string;
  date: string;
  body: string;
  folder: SpFile;
}

interface LibraryTaskItem {
  libraryName: string;
  id: string;
  fields: Record<string, unknown>;
}

interface GraphDriveResponse {
  value: Array<{
    id: string;
    name: string;
    webUrl: string;
  }>;
}

interface GraphDriveItemsResponse {
  value: Array<{
    id: string;
    name: string;
    webUrl: string;
    size?: number;
    folder?: any;
    file?: any;
    createdBy?: {
      user?: {
        displayName?: string;
        email?: string;
      };
    };
    lastModifiedBy?: {
      user?: {
        displayName?: string;
        email?: string;
      };
    };
    lastModifiedDateTime?: string;
    parentReference?: {
      driveId?: string;
      id?: string;
      path?: string;
      name?: string;
    };
    listItem?: {
      fields?: Record<string, unknown>;
    };
  }>;
  '@odata.nextLink'?: string;
}

/** Cached Files-tab listing for one document library (supports resume via nextLink). */
interface LibraryFilesCacheEntry {
  items: SpFile[];
  nextLink: string | null;
  driveId: string | null;
  complete: boolean;
}

/** Cached Comments-search task pages across document-library task lists. */
interface CommentTasksCacheEntry {
  tasks: LibraryTaskItem[];
  pendingLibraries: string[];
  currentNextLink: string | null;
  currentLibrary: string | null;
  hasMore: boolean;
  listIds: Record<string, string>;
  siteId: string | null;
}

/**
 * Same lightweight root listing style as HR Files (all-users):
 * $select only — no $expand=listItem — then page with @odata.nextLink.
 */
const DRIVE_CHILDREN_QUERY =
  '?$select=id,name,webUrl,size,folder,file,createdBy,lastModifiedBy,lastModifiedDateTime&$top=10000';
const PAGE_SIZE = 10000;
/** Comment-search pages — smaller than drive crawls so the first matches paint quickly. */
const COMMENT_TASKS_PAGE_SIZE = 200;
/** Idle warm stops here so we don't preload tens of thousands of tasks into memory. */
const COMMENT_TASKS_WARM_MAX = 1500;
/** Hard cap for an active Comments search crawl (prevents 30k+ freezes). */
const COMMENT_TASKS_SEARCH_MAX = 8000;
/** Cap on keyword rows drawn from the task crawl so broad words don't flood the list. */
const COMMENT_SEARCH_CRAWL_HITS_MAX = 500;
/** Graph drive-search pages (50 each) per library for Attachments keyword search. */
const ATTACHMENT_SEARCH_MAX_PAGES = 3;
/** Throttle Comments search list rebuilds while background pages arrive. */
const COMMENT_TASKS_SEARCH_UI_THROTTLE_MS = 400;
/** Only the SharePoint columns Comments search needs (keeps Graph + memory light). */
const COMMENT_TASKS_FIELDS_SELECT = [
  'Title', 'Name', 'Comment', 'Comments', 'Notes', 'Description', 'Body', 'field_10',
  'EmployeeName', 'Employee', 'Requestor', 'SubmittedBy', 'Submitter', 'Author',
  'CreatedBy', 'CustomCreatedBy', 'AssignedTo',
  'CustomCreatedDate', 'Created', 'SubmittedDate', 'CustomModifiedDate', 'Modified',
  'Folder', 'FolderName', 'DocumentFolder', 'RelatedFolder', 'FileLeafRef', 'FileDirRef', 'Path',
  'eFormListId', 'ListId', 'field_11',
  'RequestorComment', 'ApproverComment', 'Approver1Comment', 'Approver2Comment', 'Approver3Comment',
].join(',');
/** Pause between Graph pages / libraries so SharePoint throttling (429) stays rare. */
const GRAPH_PAGE_GAP_MS = 350;
/** Tighter gap while searching/warming Comments (still polite to Graph). */
const COMMENT_TASKS_GAP_MS = 100;
/** Debounce search keystrokes before filtering / background paging. */
const SEARCH_DEBOUNCE_MS = 250;
/** Folder cards rendered per page in the All / Files grid. */
const RENDER_PAGE_SIZE = 100;
/** Start rendering the next page when this close (px) to the bottom of the list. */
const RENDER_AHEAD_PX = 800;
/** Debounce folder-prefetch emissions so progressive paint doesn't spam the parent. */
const PREFETCH_EMIT_DEBOUNCE_MS = 1500;
/** Memory-only Comments search index (never Redis — payloads hit HTTP 413). */
const COMMENT_TASKS_CACHE_KEY = 'tasks:all-files-comments:v2-mem';
const DRIVES_CACHE_KEY = 'files:drives:v1';

@Component({
  selector: 'app-all-files',
  standalone: true,
  imports: [CommonModule, FormsModule, AppDropdownComponent, LoadingScreenComponent],
  templateUrl: './all-files.html',
  styleUrls: ['./all-files.css']
})

export class AllFilesComponent implements OnInit, OnDestroy {//component that displays all files from Sharepoint
  /** Full, unfiltered list from SharePoint (source of truth). */
  private allFiles: SpFile[] = [];
  /** Files from all libraries for search across all libraries. */
  private allLibrariesFiles: SpFile[] = [];
  /** Files currently displayed (filtered view). */
  private _files: SpFile[] = [];
  get files(): SpFile[] {
    return this._files;
  }
  set files(value: SpFile[]) {
    this._files = value;
    this.visibleFiles = value.slice(0, this.renderLimit);
  }
  /**
   * Cards actually in the DOM. "All" can hold 20k+ folders; rendering them at
   * once froze the tab for seconds, so cards are added in pages on scroll.
   */
  visibleFiles: SpFile[] = [];
  private renderLimit = RENDER_PAGE_SIZE;
  @ViewChild('contentInner') private contentInner?: ElementRef<HTMLElement>;
  isLoading = false;//loading state
  /** Loading indicator for the "All" view which aggregates folders across libraries. */
  isLoadingAllLibraries = false;
  /** True for the whole all-libraries crawl (including progressive paint) — blocks prefetch. */
  allLibrariesCrawlActive = false;
  error: string | null = null;//error message
  documentLibraries: any[] = [];
  selectedLibrary: string = '';
  @Output() fileSelected: EventEmitter<SpFile> = new EventEmitter<SpFile>();
  /** Comments filter hit selected — parent opens that comment in the Comments panel. */
  @Output() commentSelected = new EventEmitter<CommentSearchHit>();
  /** Attachments filter hit selected — parent opens that folder and focuses the file. */
  @Output() attachmentSelected = new EventEmitter<AttachmentSearchHit>();
  /** Asks parent to quietly cache Comments for All Files folders during search. */
  @Output() ensureFolderComments = new EventEmitter<{ libraryName: string; folderName: string }>();
  /** Asks parent to quietly cache Attachments for All Files folders during search. */
  @Output() ensureFolderAttachments = new EventEmitter<{ driveId: string; folderId: string }>();
  @Output() librarySelected: EventEmitter<{ name: string; driveId: string }> = new EventEmitter();
  /** Visible folders for the parent to idle-prefetch into the per-folder contents cache. */
  @Output() foldersForPrefetch = new EventEmitter<Array<{ driveId: string; folderId: string }>>();
  /** Emits when the Select status filter changes (filters Comments in the parent). */
  @Output() approvalFilterChange = new EventEmitter<string>();
  searchTerm = '';
  selectedFileId: string | null = null;
  /** Active Comments-search row (task id). */
  selectedCommentTaskId: string | null = null;
  /** Active Attachments-search row (file id). */
  selectedAttachmentFileId: string | null = null;
  /** Comment/task hits for the Comments filter search (Live-style rows). */
  commentHits: CommentSearchHit[] = [];
  /** File hits for the Attachments filter search (Live-style rows). */
  attachmentHits: AttachmentSearchHit[] = [];

  /** Level-2 subfolders: expanded folder ids, per-folder loading state, and loaded children. */
  expandedFolderIds = new Set<string>();
  loadingChildFolderIds = new Set<string>();
  private childFoldersByFolderId = new Map<string, SpFile[]>();
  private allowedChildFolderNamesByFolderId = new Map<string, Set<string>>();
  private pendingChildSubjectFolderIds = new Set<string>();

  /** Real HR task comments, passed down from AppComponent so Comments search can look at actual content. */
  @Input() commentItems: Array<{
    id: string;
    name: string;
    webUrl: string;
    description?: string;
    submittedBy?: string;
    status?: string;
    lastModifiedDateTime?: string;
    eFormDetails?: any;
  }> = [];

  /** Bound from parent so status stays in sync with the Users tab filter. */
  @Input() selectedApprovalFilter = '';

  /** SharePoint tasks from every document-library task list — Comments search, 100/page. */
  private allLibraryTasks: LibraryTaskItem[] = [];
  isLoadingCommentTasks = false;
  /** True for any Comments index Graph pull, including quiet ones with no spinner. */
  private commentTasksFetchInFlight = false;
  isLoadingMoreCommentTasks = false;
  hasMoreCommentTasks = false;
  /** True while search auto-pages remaining tasks in the background. */
  isSearchingMoreCommentTasks = false;
  /** True while Comments keyword search pages every task list in the background. */
  isCrawlingCommentTasks = false;
  /** True while Attachments search warms matching folder contents. */
  isSearchingMoreAttachments = false;
  /** Attachment hits from Graph drive search (file names + document text, every folder). */
  private graphAttachmentHits: AttachmentSearchHit[] = [];
  /** Folder lookups for Comments search, rebuilt when the folder lists change. */
  private taskFolderIndexByName = new Map<string, SpFile[]>();
  private taskFolderIndexById = new Map<string, SpFile[]>();
  /** First folder per lowercase name (allLibrariesFiles before allFiles). */
  private folderIndexByKey = new Map<string, SpFile>();
  private taskFolderIndexLibs: SpFile[] | null = null;
  private taskFolderIndexFiles: SpFile[] | null = null;
  private taskFolderIndexSize = -1;
  /** Lowercased search text per task/comment object, so rebuilds don't re-parse HTML. */
  private commentHaystackCache = new WeakMap<object, string>();
  /** True while search auto-pages remaining files in the background. */
  isSearchingMoreFiles = false;
  /** True while Microsoft Graph drive search is running for the current query. */
  isGraphSearching = false;
  /** Folder hits from Graph drive search (merged into Files/All filtered views). */
  private graphSearchFolders: SpFile[] = [];
  private graphSearchSeq = 0;
  get hasLoadedCommentTasks(): boolean {
    return this.allLibraryTasks.length > 0;
  }
  get allLibraryTasksCount(): number {
    return this.allLibraryTasks.length;
  }
  private documentLibraryTaskMap: Record<string, string> = {};

  private commentTasksSiteId: string | null = null;
  private commentTaskListIds: Record<string, string> = {};
  private commentTasksPendingLibraries: string[] = [];
  private commentTasksCurrentNextLink: string | null = null;
  private commentTasksCurrentLibrary: string | null = null;
  /** Prevents duplicate idle warm of the Comments search index. */
  private commentTasksWarmStarted = false;
  private commentTasksWarmAborted = false;
  /** True while the user is interacting — warm crawl yields to keep clicks responsive. */
  private commentTasksWarmPaused = false;

  /** Files pagination (selected library) — same page cursor pattern as HR Files. */
  isLoadingMoreFiles = false;
  hasMoreFiles = false;
  private filesNextLink: string | null = null;
  private filesCurrentDriveId: string | null = null;

  private static readonly ALL_ROOT_CACHE_KEY = 'files:all-root:v5-paged';
  /** Libraries whose All-root crawl finished completely (safe to treat as no Load more). */
  private fullyLoadedAllLibraries = new Set<string>();

  private searchDebounceTimer: ReturnType<typeof setTimeout> | null = null;
  private prefetchEmitTimer: ReturnType<typeof setTimeout> | null = null;
  private commentSearchFilterTimer: ReturnType<typeof setTimeout> | null = null;
  private lastPrefetchSignature = '';
  private readonly subscriptions = new Subscription();
  private viewDestroyed = false;

  constructor(
    private http: HttpClient,
    private authService: AuthService,
    private documentService: DocumentService,
    private cdr: ChangeDetectorRef,
    private formConfig: FormConfigurationService,
    private fileCrawlCache: FileCrawlCacheService,
  ) {}

  /**
   * Runs change detection only while the view is alive. Async crawls finish after
   * the user navigates away; detectChanges on a torn-down view throws.
   */
  private refreshView(): void {
    if (this.viewDestroyed) return;
    // The app runs without zone.js: markForCheck also schedules a full app pass, so the
    // list still paints if this component-only pass ran too early or was interrupted.
    this.cdr.markForCheck();
    this.cdr.detectChanges();
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private normalizeCacheKeyPart(value: string): string {
    return String(value ?? '').trim().toLowerCase();
  }

  private libraryFilesCacheKey(libraryName: string): string {
    return `files:lib:v5-paged:${this.normalizeCacheKeyPart(libraryName)}`;
  }

  private readLibraryFilesCache(libraryName: string): LibraryFilesCacheEntry | null {
    const raw = this.fileCrawlCache.get<LibraryFilesCacheEntry | SpFile[]>(
      this.libraryFilesCacheKey(libraryName)
    );
    if (!raw) return null;
    // Legacy array shape from older cache entries.
    if (Array.isArray(raw)) {
      return { items: raw, nextLink: null, driveId: raw.find((f) => f.driveId)?.driveId ?? null, complete: true };
    }
    if (!Array.isArray(raw.items)) return null;
    return raw;
  }

  private writeLibraryFilesCache(
    libraryName: string,
    items: SpFile[],
    nextLink: string | null,
    driveId: string | null,
  ): void {
    const entry: LibraryFilesCacheEntry = {
      items,
      nextLink,
      driveId,
      complete: !nextLink,
    };
    this.fileCrawlCache.set(this.libraryFilesCacheKey(libraryName), entry);
  }

  private seedLibraryCachesFromAllRoot(items: SpFile[]): void {
    const byLibrary = new Map<string, SpFile[]>();
    for (const item of items) {
      const name = item.libraryName;
      if (!name) continue;
      const list = byLibrary.get(name) ?? [];
      list.push(item);
      byLibrary.set(name, list);
    }
    for (const [libraryName, libItems] of byLibrary) {
      // Only mark complete when that library's All crawl finished — never wipe a
      // Files-tab nextLink from a partial All snapshot.
      if (!this.fullyLoadedAllLibraries.has(libraryName)) continue;
      const existing = this.readLibraryFilesCache(libraryName);
      // Already saved complete (shared Redis entry) — rewriting it would re-upload the same listing.
      if (existing?.complete) continue;
      if (existing && !existing.complete && (existing.items?.length ?? 0) > libItems.length) {
        continue;
      }
      const driveId = libItems.find((f) => f.driveId)?.driveId ?? null;
      this.writeLibraryFilesCache(libraryName, libItems, null, driveId);
    }
  }

  private async getDrivesCached(token: string): Promise<GraphDriveResponse> {
    const cached = this.fileCrawlCache.get<GraphDriveResponse>(DRIVES_CACHE_KEY);
    if (cached?.value?.length) return cached;

    const siteId = sharePointConfig.siteId;
    const drivesResponse = await graphGetWithRetry(
      this.http,
      `/sites/${siteId}/drives`,
      token,
      AppConstants.graphFileListingTimeoutMs
    ) as GraphDriveResponse;
    this.fileCrawlCache.set(DRIVES_CACHE_KEY, drivesResponse);
    return drivesResponse;
  }

  private applyLibraryFilesToView(
    items: SpFile[],
    driveId: string | null,
    nextLink: string | null,
  ): void {
    this.allFiles = items;
    this.filesCurrentDriveId = driveId;
    this.filesNextLink = nextLink;
    this.hasMoreFiles = !!nextLink;
    this.applyFilters();
    if (driveId) {
      try {
        this.librarySelected.emit({ name: this.selectedLibrary, driveId });
      } catch { /* ignore */ }
    }
  }

  /** Keep Graph nextLink usable even when toGraphPath cannot strip the host. */
  private toNextPath(nextLink?: string | null): string | null {
    if (!nextLink) return null;
    const path = toGraphPath(nextLink);
    if (path) return path;
    const match = nextLink.match(/^https?:\/\/graph\.microsoft\.com\/[^/]+(\/.*)$/i);
    if (match?.[1]) return match[1];
    if (nextLink.startsWith('/')) return nextLink;
    return nextLink;
  }

  private invalidateAllFilesCrawlCache(): void {
    this.fileCrawlCache.invalidate(AllFilesComponent.ALL_ROOT_CACHE_KEY);
    this.fileCrawlCache.invalidate(COMMENT_TASKS_CACHE_KEY);
    this.fileCrawlCache.invalidate(DRIVES_CACHE_KEY);
    this.fullyLoadedAllLibraries.clear();
    if (this.selectedLibrary) {
      this.fileCrawlCache.invalidate(this.libraryFilesCacheKey(this.selectedLibrary));
    }
    for (const library of this.documentLibraries) {
      this.fileCrawlCache.invalidate(this.libraryFilesCacheKey(library.name));
    }
  }

  ngOnInit(): void {
    // Subscribe here, not in the constructor: documentLibraries$ replays synchronously
    // when forms are already loaded, and a cached crawl would run detectChanges before
    // the view has its component context (TypeError reading 'filterType' of null).
    this.subscriptions.add(
      this.formConfig.documentLibraries$.subscribe((libraries) => {
        this.documentLibraries = libraries;
        if (libraries.length > 0 && !this.selectedLibrary) {
          this.selectedLibrary = libraries[0].name;
          // Default view is "All", which stays empty until the user searches,
          // so nothing is crawled at startup. See runSearch().
        }
      }),
    );
    this.subscriptions.add(
      this.formConfig.documentLibraryTaskMap$.subscribe((map) => {
        this.documentLibraryTaskMap = map ?? {};
        // Comments search uses per-folder All Files caches — do not warm a global 30k+ index.
      }),
    );
  }

  ngOnDestroy(): void {
    this.viewDestroyed = true;
    this.subscriptions.unsubscribe();
    this.commentTasksWarmAborted = true;
    if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
    if (this.prefetchEmitTimer) clearTimeout(this.prefetchEmitTimer);
    if (this.commentSearchFilterTimer) clearTimeout(this.commentSearchFilterTimer);
  }

  async loadAllLibrariesFiles(force = false): Promise<void> {
    if ((this.isLoadingAllLibraries || this.allLibrariesCrawlActive) && !force) return;

    if (force) {
      this.fileCrawlCache.invalidate(AllFilesComponent.ALL_ROOT_CACHE_KEY);
    } else {
      const cached = this.fileCrawlCache.get<SpFile[]>(AllFilesComponent.ALL_ROOT_CACHE_KEY);
      // An empty snapshot is a crawl that found nothing — load again rather than show a blank list.
      if (Array.isArray(cached) && cached.length > 0) {
        this.allLibrariesFiles = cached;
        // ALL_ROOT is only written after a finished crawl — mark those libraries complete.
        for (const file of this.allLibrariesFiles) {
          if (file.libraryName) this.fullyLoadedAllLibraries.add(file.libraryName);
        }
        this.seedLibraryCachesFromAllRoot(this.allLibrariesFiles);
        if (this.filterType !== 'files') {
          this.applyFilters();
        }
        this.isLoadingAllLibraries = false;
        this.allLibrariesCrawlActive = false;
        this.refreshView();
        return;
      }
    }

    this.isLoadingAllLibraries = true;
    this.allLibrariesCrawlActive = true;
    this.error = null;
    this.refreshView();
    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token) return;

      const drivesResponse = await this.getDrivesCached(token);

      // Keep the current list on screen while reloading; each library swaps in its
      // own rows as they arrive. Clearing up front made whole libraries (e.g. Call
      // for Applications) vanish mid-refresh, and for good when one fetch failed.
      const previousFiles = this.allLibrariesFiles;
      let anyLibraryFailed = false;
      const drivesByName = new Map(
        (drivesResponse.value ?? []).map((d) => [d.name, d])
      );

      // Sequential (like HR Files) — parallel library hits caused Graph 429s.
      for (const library of this.documentLibraries) {
        const drive = drivesByName.get(library.name);
        if (!drive) continue;

        // Prefer a complete per-library cache so switching libraries / remounts stay instant.
        // Pull this library's shared Redis entry into memory first (no-op if already loaded).
        await this.fileCrawlCache.hydrateFromPersistent(this.libraryFilesCacheKey(library.name));
        const libCached = this.readLibraryFilesCache(library.name);
        if (libCached?.complete && libCached.items.length > 0) {
          this.fullyLoadedAllLibraries.add(library.name);
          const tagged = libCached.items.map((item) => ({
            ...item,
            libraryName: library.name,
            driveId: item.driveId ?? libCached.driveId ?? drive.id,
          }));
          const others = this.allLibrariesFiles.filter((f) => f.libraryName !== library.name);
          this.allLibrariesFiles = [...others, ...tagged];
          if (this.filterType !== 'files') {
            this.applyFilters();
            this.isLoadingAllLibraries = false;
            this.refreshView();
          }
          continue;
        }

        try {
          const items = await this.fetchDriveRootChildrenPages(
            drive.id,
            token,
            library.name,
            (partial) => {
              // Progressive paint: merge this library's pages into the visible list.
              const others = this.allLibrariesFiles.filter((f) => f.libraryName !== library.name);
              this.allLibrariesFiles = [...others, ...partial];
              if (this.filterType !== 'files') {
                this.applyFilters();
                this.isLoadingAllLibraries = false;
                this.refreshView();
              }
            }
          );
          const others = this.allLibrariesFiles.filter((f) => f.libraryName !== library.name);
          this.allLibrariesFiles = [...others, ...items];
          this.fullyLoadedAllLibraries.add(library.name);
          this.writeLibraryFilesCache(library.name, items, null, drive.id);
        } catch (err) {
          console.error(`Failed to load files from library ${library.name}:`, err);
          anyLibraryFailed = true;
          // Put back what we had for this library rather than a partial/empty list.
          const previousForLibrary = previousFiles.filter((f) => f.libraryName === library.name);
          const current = this.allLibrariesFiles.filter((f) => f.libraryName === library.name);
          if (previousForLibrary.length > current.length) {
            const others = this.allLibrariesFiles.filter((f) => f.libraryName !== library.name);
            this.allLibrariesFiles = [...others, ...previousForLibrary];
          }
        }

        // Pace libraries so Graph doesn't throttle the session.
        await this.sleep(GRAPH_PAGE_GAP_MS);
      }

      // Drop rows for libraries no longer configured (list was not cleared up front).
      const configured = new Set(this.documentLibraries.map((l) => l.name));
      this.allLibrariesFiles = this.allLibrariesFiles.filter(
        (f) => !f.libraryName || configured.has(f.libraryName)
      );

      // ALL_ROOT means "finished crawl" — never cache a list missing a failed library.
      // Memory only: it merges every library, so it is not shareable. Each library's own
      // listing is saved once (shared, access-checked) by writeLibraryFilesCache instead.
      if (!anyLibraryFailed && this.allLibrariesFiles.length > 0) {
        this.fileCrawlCache.setMemoryOnly(AllFilesComponent.ALL_ROOT_CACHE_KEY, this.allLibrariesFiles);
      }
      this.seedLibraryCachesFromAllRoot(this.allLibrariesFiles);
      if (anyLibraryFailed && this.allLibrariesFiles.length === 0) {
        this.error = 'Could not load folders from SharePoint. Please refresh to try again.';
      }
    } catch (err: any) {
      console.error('Failed to load all libraries files:', err);
      if (this.allLibrariesFiles.length === 0) {
        this.error = 'Could not load folders from SharePoint. Please refresh to try again.';
      }
    } finally {
      if (this.filterType !== 'files') {
        this.applyFilters();
      }
      this.isLoadingAllLibraries = false;
      this.allLibrariesCrawlActive = false;
      this.refreshView();
    }
  }

  /**
   * Same pattern as AppComponent.fetchHrPersonalRootFolders:
   * page with $top=100 + @odata.nextLink, paint after each page, retry on 429.
   */
  private async fetchDriveRootChildrenPages(
    driveId: string,
    token: string,
    libraryName: string,
    onPage?: (items: SpFile[]) => void,
  ): Promise<SpFile[]> {
    const collected: SpFile[] = [];
    let nextPath: string | null =
      `/drives/${driveId}/root/children${DRIVE_CHILDREN_QUERY}`;

    while (nextPath) {
      const page = await graphGetWithRetry(
        this.http,
        nextPath,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as GraphDriveItemsResponse;

      for (const item of page.value ?? []) {
        const spFile = this.mapToSpFile(item, driveId);
        spFile.libraryName = libraryName;
        collected.push(spFile);
      }
      onPage?.(collected.slice());
      nextPath = this.toNextPath(page['@odata.nextLink']);
      if (nextPath) await this.sleep(GRAPH_PAGE_GAP_MS);
    }

    return collected;
  }

  /** Pause idle Comments indexing while the user opens a search hit. */
  pauseCommentTasksWarm(): void {
    this.commentTasksWarmPaused = true;
  }

  /** Resume idle Comments indexing after interaction settles. */
  resumeCommentTasksWarm(): void {
    if (!this.commentTasksWarmPaused) return;
    this.commentTasksWarmPaused = false;
    void this.continueCommentTasksWarmInBackground();
  }

  /**
   * Idle-warm the Comments search index (Redis → memory → Graph) so the first
   * Comments search feels as instant as All / Files.
   */
  private async warmCommentTasksIndex(): Promise<void> {
    if (this.commentTasksWarmStarted || this.commentTasksWarmAborted) return;
    this.commentTasksWarmStarted = true;
    try {
      await this.loadAllLibraryTasksForCommentsSearch({ quiet: true });
      if (!this.commentTasksWarmAborted && !this.commentTasksWarmPaused) {
        void this.continueCommentTasksWarmInBackground();
      }
    } catch (err) {
      console.warn('Comments search warm failed:', err);
      this.commentTasksWarmStarted = false;
    }
  }

  /** Keep paging Comments tasks in the background until the shared cache is complete. */
  private async continueCommentTasksWarmInBackground(): Promise<void> {
    if (!this.hasMoreCommentTasks || this.commentTasksWarmAborted) return;

    try {
      while (this.hasMoreCommentTasks && !this.commentTasksWarmAborted) {
        if (this.allLibraryTasks.length >= COMMENT_TASKS_WARM_MAX) {
          // Leave hasMore true so an active search can keep paging past the warm cap.
          break;
        }
        if (
          this.commentTasksWarmPaused ||
          this.commentTasksFetchInFlight ||
          this.isCrawlingCommentTasks
        ) {
          await this.sleep(COMMENT_TASKS_GAP_MS);
          continue;
        }
        const token = await this.authService.acquireSharePointToken();
        if (!token) break;
        await this.fetchNextCommentTasksPage(token);
        if (this.filterType === 'comments') this.refreshView();
        await this.sleep(COMMENT_TASKS_GAP_MS);
      }
    } catch (err) {
      console.warn('Background Comments warm failed:', err);
    }
  }

  /** Apply a Comments-search cache entry into memory and refresh the UI if needed. */
  private applyCommentTasksCacheEntry(cached: CommentTasksCacheEntry): void {
    this.allLibraryTasks = cached.tasks ?? [];
    this.commentTasksPendingLibraries = [...(cached.pendingLibraries ?? [])];
    this.commentTasksCurrentNextLink = cached.currentNextLink ?? null;
    this.commentTasksCurrentLibrary = cached.currentLibrary ?? null;
    this.hasMoreCommentTasks = !!cached.hasMore;
    this.commentTaskListIds = { ...(cached.listIds ?? {}) };
    this.commentTasksSiteId = cached.siteId ?? sharePointConfig.siteId;
    if (this.filterType === 'comments') {
      this.applyFilters();
      this.refreshView();
    }
  }

  /** Resolves list IDs once, then pages tasks — memory cache first (never Redis for this index). */
  private async loadAllLibraryTasksForCommentsSearch(
    options: { quiet?: boolean } = {},
  ): Promise<void> {
    if (this.isLoadingCommentTasks || this.commentTasksFetchInFlight) return;
    if (this.allLibraryTasks.length > 0) {
      if (this.filterType === 'comments') this.applyFilters();
      return;
    }

    // Drop legacy Redis snapshots that caused HTTP 413 (full-field 30k+ task blobs).
    this.fileCrawlCache.invalidate('tasks:all-files-comments:v1');

    const fresh = this.fileCrawlCache.get<CommentTasksCacheEntry>(COMMENT_TASKS_CACHE_KEY);
    if (fresh?.tasks?.length) {
      this.applyCommentTasksCacheEntry(fresh);
      return;
    }

    const stale = this.fileCrawlCache.getStale<CommentTasksCacheEntry>(COMMENT_TASKS_CACHE_KEY);
    if (stale?.tasks?.length) {
      this.applyCommentTasksCacheEntry(stale);
      if (!fresh) {
        void this.refreshCommentTasksFromGraph({ quiet: true });
      }
      return;
    }

    await this.refreshCommentTasksFromGraph(options);
  }

  /** Cold Graph crawl for the Comments search index. */
  private async refreshCommentTasksFromGraph(
    options: { quiet?: boolean } = {},
  ): Promise<void> {
    if (this.isLoadingCommentTasks || this.commentTasksFetchInFlight) return;

    // Quiet (background) pulls must not flip the visible spinner flag — the
    // template uses it to hide results, which made folders vanish mid-refresh.
    this.commentTasksFetchInFlight = true;
    const showSpinner = !options.quiet && this.filterType === 'comments';
    if (showSpinner) {
      this.isLoadingCommentTasks = true;
      this.hasMoreCommentTasks = false;
      this.refreshView();
    }

    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token) return;

      // Reset crawl cursor only when starting a fresh Graph pull with no data yet.
      if (this.allLibraryTasks.length === 0) {
        await this.ensureCommentTaskListIds(token);
        this.commentTasksPendingLibraries = Object.keys(this.commentTaskListIds);
        this.commentTasksCurrentNextLink = null;
        this.commentTasksCurrentLibrary = null;
      } else if (Object.keys(this.commentTaskListIds).length === 0) {
        await this.ensureCommentTaskListIds(token);
      }

      if (this.allLibraryTasks.length === 0) {
        await this.fetchNextCommentTasksPage(token);
        this.persistCommentTasksCache();
      }
    } catch (err) {
      console.error('Failed to load library tasks for comment search:', err);
    } finally {
      this.commentTasksFetchInFlight = false;
      this.isLoadingCommentTasks = false;
      this.refreshView();
    }
  }

  private persistCommentTasksCache(): void {
    const entry: CommentTasksCacheEntry = {
      tasks: this.allLibraryTasks,
      pendingLibraries: [...this.commentTasksPendingLibraries],
      currentNextLink: this.commentTasksCurrentNextLink,
      currentLibrary: this.commentTasksCurrentLibrary,
      hasMore: this.hasMoreCommentTasks,
      listIds: { ...this.commentTaskListIds },
      siteId: this.commentTasksSiteId,
    };
    // Memory only — full Comments index is far too large for Redis (HTTP 413).
    this.fileCrawlCache.setMemoryOnly(COMMENT_TASKS_CACHE_KEY, entry);
  }

  /** Keep only columns used for Comments search / folder resolve / hit preview. */
  private stripCommentTaskFields(fields: Record<string, unknown>): Record<string, unknown> {
    const out: Record<string, unknown> = {};
    for (const key of COMMENT_TASKS_FIELDS_SELECT.split(',')) {
      if (fields[key] != null && fields[key] !== '') {
        out[key] = fields[key];
      }
    }
    return out;
  }

  /** Rebuild Comments search hits without blocking every Graph page. */
  private scheduleCommentSearchFilter(force = false): void {
    if (this.filterType !== 'comments') return;
    if (force) {
      if (this.commentSearchFilterTimer) {
        clearTimeout(this.commentSearchFilterTimer);
        this.commentSearchFilterTimer = null;
      }
      this.applyFilters();
      this.refreshView();
      return;
    }
    if (this.commentSearchFilterTimer) return;
    this.commentSearchFilterTimer = setTimeout(() => {
      this.commentSearchFilterTimer = null;
      this.applyFilters();
      this.refreshView();
    }, COMMENT_TASKS_SEARCH_UI_THROTTLE_MS);
  }

  private async ensureCommentTaskListIds(token: string): Promise<void> {
    if (Object.keys(this.commentTaskListIds).length > 0) return;

    this.commentTasksSiteId = sharePointConfig.siteId;
    const listsResponse = await graphGetWithRetry(
      this.http,
      `/sites/${this.commentTasksSiteId}/lists?$select=id,name,displayName`,
      token,
      AppConstants.graphFileListingTimeoutMs
    ) as { value?: Array<{ id: string; name: string; displayName?: string }> };
    const siteLists = listsResponse.value ?? [];

    for (const [libraryName, taskListName] of Object.entries(this.documentLibraryTaskMap)) {
      const listObj = siteLists.find((l) =>
        (l.name ?? '').toLowerCase() === taskListName.toLowerCase() ||
        (l.displayName ?? '').toLowerCase() === taskListName.toLowerCase()
      );
      if (listObj?.id) this.commentTaskListIds[libraryName] = listObj.id;
    }
  }

  async loadMoreCommentTasks(): Promise<void> {
    if (this.isLoadingMoreCommentTasks || !this.hasMoreCommentTasks) return;
    this.isLoadingMoreCommentTasks = true;
    this.refreshView();
    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token) return;
      await this.fetchNextCommentTasksPage(token);
    } catch (err) {
      console.error('Failed to load more comment tasks:', err);
    } finally {
      this.isLoadingMoreCommentTasks = false;
      this.refreshView();
    }
  }

  /**
   * When the user searches Comments, keep paging in the background until every
   * task page is loaded (or the search is cleared) so they don't need Load more.
   * Matches appear progressively via applyFilters after each page.
   */
  private async continueCommentTasksSearchInBackground(): Promise<void> {
    if (this.isCrawlingCommentTasks) return;
    if (!this.searchTerm.trim()) return;

    this.isCrawlingCommentTasks = true;
    this.refreshView();
    try {
      if (this.allLibraryTasks.length === 0) {
        await this.loadAllLibraryTasksForCommentsSearch({ quiet: true });
      }
      while (
        this.hasMoreCommentTasks &&
        this.searchTerm.trim() &&
        this.filterType === 'comments' &&
        this.allLibraryTasks.length < COMMENT_TASKS_SEARCH_MAX
      ) {
        if (this.commentTasksWarmPaused) {
          await this.sleep(COMMENT_TASKS_GAP_MS);
          continue;
        }
        const token = await this.authService.acquireSharePointToken();
        if (!token) break;
        await this.fetchNextCommentTasksPage(token);
        this.refreshView();
        await this.sleep(COMMENT_TASKS_GAP_MS);
      }
    } catch (err) {
      console.error('Background comment-task search failed:', err);
    } finally {
      this.isCrawlingCommentTasks = false;
      this.scheduleCommentSearchFilter(true);
    }
  }

  /**
   * When searching Files, auto-load remaining pages in the background so matches
   * beyond the first page are found without clicking Load more.
   */
  private async continueFilesSearchInBackground(): Promise<void> {
    if (this.isSearchingMoreFiles) return;
    if (!this.searchTerm.trim() || !this.hasMoreFiles) return;

    this.isSearchingMoreFiles = true;
    this.refreshView();
    try {
      while (this.hasMoreFiles && this.searchTerm.trim() && this.filterType === 'files') {
        const before = this.allFiles.length;
        await this.loadMoreFiles();
        if (this.allFiles.length === before) break;
        await this.sleep(GRAPH_PAGE_GAP_MS);
      }
    } catch (err) {
      console.error('Background file search failed:', err);
    } finally {
      this.isSearchingMoreFiles = false;
      this.refreshView();
    }
  }

  private async fetchNextCommentTasksPage(token: string): Promise<void> {
    if (!this.commentTasksCurrentNextLink) {
      const nextLibrary = this.commentTasksPendingLibraries.shift();
      if (!nextLibrary) {
        this.hasMoreCommentTasks = false;
        return;
      }
      const listId = this.commentTaskListIds[nextLibrary];
      this.commentTasksCurrentLibrary = nextLibrary;
      this.commentTasksCurrentNextLink = this.buildCommentTasksListPath(listId, true);
    }

    const path = this.toNextPath(this.commentTasksCurrentNextLink) ?? this.commentTasksCurrentNextLink!;
    let response: { value?: Array<{ id: string; fields?: Record<string, unknown> }>; '@odata.nextLink'?: string };
    try {
      response = await graphGetWithRetry(
        this.http,
        path,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as { value?: Array<{ id: string; fields?: Record<string, unknown> }>; '@odata.nextLink'?: string };
    } catch (err) {
      // Some lists reject fields($select=...) — retry once with a full fields expand.
      const listId = this.commentTaskListIds[this.commentTasksCurrentLibrary ?? ''];
      if (listId && path.includes('fields($select=')) {
        this.commentTasksCurrentNextLink = this.buildCommentTasksListPath(listId, false);
        response = await graphGetWithRetry(
          this.http,
          this.commentTasksCurrentNextLink,
          token,
          AppConstants.graphFileListingTimeoutMs
        ) as { value?: Array<{ id: string; fields?: Record<string, unknown> }>; '@odata.nextLink'?: string };
      } else {
        throw err;
      }
    }

    const libraryName = this.commentTasksCurrentLibrary!;
    for (const item of response.value ?? []) {
      this.allLibraryTasks.push({
        libraryName,
        id: item.id,
        fields: this.stripCommentTaskFields(item.fields ?? {}),
      });
    }

    this.commentTasksCurrentNextLink = this.toNextPath(response['@odata.nextLink']);
    if (!this.commentTasksCurrentNextLink) this.commentTasksCurrentLibrary = null;

    this.hasMoreCommentTasks = !!this.commentTasksCurrentNextLink || this.commentTasksPendingLibraries.length > 0;
    this.persistCommentTasksCache();
    this.scheduleCommentSearchFilter();
  }

  private buildCommentTasksListPath(listId: string, selectFields: boolean): string {
    const expand = selectFields
      ? `fields($select=${COMMENT_TASKS_FIELDS_SELECT})`
      : 'fields';
    return (
      `/sites/${this.commentTasksSiteId}/lists/${listId}/items` +
      `?$expand=${expand}&$top=${COMMENT_TASKS_PAGE_SIZE}`
    );
  }

  /** Emit when a file or folder is clicked in the All Files list. */
  selectFile(file: SpFile): void {
    this.selectedFileId = file.id;
    this.selectedCommentTaskId = null;
    this.selectedAttachmentFileId = null;
    this.fileSelected.emit(file);
  }

  /** Comments-search row click: open the related folder and focus that comment. */
  selectCommentHit(hit: CommentSearchHit): void {
    this.selectedCommentTaskId = hit.taskId;
    this.selectedAttachmentFileId = null;
    this.selectedFileId = hit.folder?.id ?? null;
    this.commentSelected.emit(hit);
  }

  /** Attachments-search row click: open the related folder and focus that file. */
  selectAttachmentHit(hit: AttachmentSearchHit): void {
    this.selectedAttachmentFileId = hit.fileId;
    this.selectedCommentTaskId = null;
    this.selectedFileId = hit.folder?.id ?? null;
    this.attachmentSelected.emit(hit);
  }

  /** Level-1 card click: select folder for Attachments/Comments only. */
  onParentCardClick(file: SpFile): void {
    this.selectFile(file);
  }

  isFileSelected(file: SpFile): boolean {
    return this.selectedFileId === file.id;
  }

  trackFile(_: number, file: SpFile): string {
    return `${file.driveId ?? ''}:${file.id}`;
  }

  /** Renders the next page of folder cards once the list is scrolled near its end. */
  onContentScroll(event: Event): void {
    if (this.visibleFiles.length >= this._files.length) return;
    const el = event.target as HTMLElement;
    if (el.scrollTop + el.clientHeight < el.scrollHeight - RENDER_AHEAD_PX) return;
    this.renderLimit += RENDER_PAGE_SIZE;
    this.visibleFiles = this._files.slice(0, this.renderLimit);
  }

  /** New search / filter / library: go back to the first page and the top of the list. */
  private resetRenderedPage(): void {
    this.renderLimit = RENDER_PAGE_SIZE;
    this.visibleFiles = this._files.slice(0, this.renderLimit);
    this.contentInner?.nativeElement.scrollTo({ top: 0 });
  }

  /** Keep result cards in the DOM across rebuilds instead of redrawing all of them. */
  trackCommentHit(_: number, hit: CommentSearchHit): string {
    return hit.taskId;
  }

  trackAttachmentHit(_: number, hit: AttachmentSearchHit): string {
    return hit.fileId;
  }

  isCommentHitSelected(hit: CommentSearchHit): boolean {
    return this.selectedCommentTaskId === hit.taskId;
  }

  isAttachmentHitSelected(hit: AttachmentSearchHit): boolean {
    return this.selectedAttachmentFileId === hit.fileId;
  }

  // ── Level-2 children folders ──────────────────────────────────

  isFolderExpanded(folder: SpFile): boolean {
    return this.expandedFolderIds.has(folder.id);
  }

  isLoadingChildFolders(folder: SpFile): boolean {
    return this.loadingChildFolderIds.has(folder.id);
  }

  /** Show expand chevron only when ChildSubject / ChildSubject2 name(s) exist. */
  hasChildSubjectDropdown(folder: SpFile): boolean {
    if (!folder.isfolder || !folder.driveId) return false;
    if (this.pendingChildSubjectFolderIds.has(folder.id)) return false;
    const allowedNames = this.allowedChildFolderNamesByFolderId.get(folder.id);
    return !!allowedNames && allowedNames.size > 0;
  }

  getChildFolders(folder: SpFile): SpFile[] {
    const allowedNames = this.allowedChildFolderNamesByFolderId.get(folder.id);
    if (!allowedNames || allowedNames.size === 0) return [];

    const resolved = this.resolveChildSubjectFolders(folder, allowedNames);
    const physical = this.childFoldersByFolderId.get(folder.id) ?? [];
    return this.uniqueFolders([
      ...resolved,
      ...physical.filter((child) => this.folderNameMatchesAllowed(child.name, allowedNames)),
    ]);
  }

  private folderNameMatchesAllowed(fileName: string, allowedNames: Set<string>): boolean {
    const normalizedName = this.normalizeFolderName(fileName);
    if (allowedNames.has(normalizedName)) return true;
    const fileId = this.extractTrailingFolderId(fileName);
    if (!fileId) return false;
    for (const allowed of allowedNames) {
      if (this.extractTrailingFolderId(allowed) === fileId) return true;
    }
    return false;
  }

  /** Wait for the clicked task's ChildSubject columns before showing level-2 folders. */
  beginChildSubjectLoad(folderId: string): void {
    if (!folderId) return;
    this.allowedChildFolderNamesByFolderId.delete(folderId);
    this.childFoldersByFolderId.delete(folderId);
    this.expandedFolderIds.delete(folderId);
    this.pendingChildSubjectFolderIds.add(folderId);
    this.refreshView();
  }

  /** Display each folder named by ChildSubject/ChildSubject2 once — no other subfolders. */
  setChildSubjectFolderNames(folderId: string, folderNames: string[]): void {
    if (!folderId) return;
    const uniqueNames = new Set(
      (folderNames ?? [])
        .map((name) => this.normalizeFolderName(name))
        .filter(Boolean)
    );
    this.allowedChildFolderNamesByFolderId.set(folderId, uniqueNames);
    this.pendingChildSubjectFolderIds.delete(folderId);

    if (uniqueNames.size === 0) {
      this.expandedFolderIds.delete(folderId);
      this.childFoldersByFolderId.delete(folderId);
    } else {
      // Do not pre-fill childFoldersByFolderId here — an empty array blocks loadChildFolders
      // on expand. Resolve from drive cache / Graph when ChildSubject names arrive.
      const parent = this.findFolderById(folderId);
      if (parent) {
        void this.hydrateChildFolders(parent);
      }
    }
    this.refreshView();
  }

  /** Called when Attachments finishes loading a folder's drive children into cache. */
  refreshChildFoldersForFolder(folderId: string): void {
    const parent = this.findFolderById(folderId);
    if (parent) {
      void this.hydrateChildFolders(parent);
    }
  }

  private findFolderById(folderId: string): SpFile | undefined {
    return (
      this.files.find((file) => file.id === folderId) ??
      this.allLibrariesFiles.find((file) => file.id === folderId)
    );
  }

  /** Match ChildSubject names to folders in the same library (root or nested). */
  private resolveChildSubjectFolders(parent: SpFile, allowedNames: Set<string>): SpFile[] {
    const libraryName = parent.libraryName;
    if (!libraryName) return [];

    const allowedIds = new Set<string>();
    for (const name of allowedNames) {
      const id = this.extractTrailingFolderId(name);
      if (id) allowedIds.add(id);
    }

    return this.allLibrariesFiles
      .filter((file) => {
        if (!file.isfolder || file.libraryName !== libraryName || file.id === parent.id) {
          return false;
        }
        const normalizedName = this.normalizeFolderName(file.name);
        if (allowedNames.has(normalizedName)) return true;
        const fileId = this.extractTrailingFolderId(file.name);
        return !!fileId && allowedIds.has(fileId);
      })
      .map((file) => ({
        ...file,
        parentFolderId: parent.id,
        parentFolderName: parent.name,
      }));
  }

  private extractTrailingFolderId(value: string): string {
    return String(value ?? '').trim().match(/(?:^|[-_\s])(\d+)\s*$/)?.[1] ?? '';
  }

  private normalizeFolderName(value: string): string {
    return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  private uniqueFolders(folders: SpFile[]): SpFile[] {
    const seen = new Set<string>();
    return folders.filter((folder) => {
      const key = `${folder.driveId ?? ''}:${folder.id || this.normalizeFolderName(folder.name)}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  /** Chevron toggle — only when ChildSubject / ChildSubject2 folders exist. */
  toggleFolderChildren(folder: SpFile, event?: Event): void {
    event?.stopPropagation();
    if (!this.hasChildSubjectDropdown(folder)) return;

    if (this.expandedFolderIds.has(folder.id)) {
      this.expandedFolderIds.delete(folder.id);
      this.refreshView();
      return;
    }
    this.expandedFolderIds.clear();
    this.expandedFolderIds.add(folder.id);
    void this.hydrateChildFolders(folder);
    this.refreshView();
  }

  /** Resolve ChildSubject folders from library root + filtered physical drive children. */
  private async hydrateChildFolders(folder: SpFile): Promise<void> {
    const allowedNames = this.allowedChildFolderNamesByFolderId.get(folder.id);
    if (!allowedNames || allowedNames.size === 0 || !folder.driveId) return;

    const cacheKey = this.driveFolderCacheKey(folder.driveId, folder.id);
    const cached = this.fileCrawlCache.get<any[]>(cacheKey);
    if (Array.isArray(cached)) {
      this.childFoldersByFolderId.set(
        folder.id,
        this.filterAllowedChildFolders(this.mapCachedChildrenToSpFiles(cached, folder), allowedNames)
      );
      this.refreshView();
      return;
    }

    if (this.getChildFolders(folder).length === 0) {
      await this.loadChildFolders(folder);
    }
  }

  private filterAllowedChildFolders(children: SpFile[], allowedNames: Set<string>): SpFile[] {
    return children.filter((child) => this.folderNameMatchesAllowed(child.name, allowedNames));
  }

  /** Same key AppComponent uses for its idle folder prefetch, so we reuse warmed entries. */
  private driveFolderCacheKey(driveId: string, folderId: string): string {
    return `files:folder:${driveId}:${folderId}`;
  }

  /**
   * Load one level of subfolders for a folder card — same loading-state pattern
   * as loadFiles(): cache first (memory, then shared Redis), then Graph.
   */
  private async loadChildFolders(folder: SpFile): Promise<void> {
    if (!folder.driveId || this.loadingChildFolderIds.has(folder.id)) return;

    const allowedNames = this.allowedChildFolderNamesByFolderId.get(folder.id);
    if (!allowedNames || allowedNames.size === 0) return;

    const cacheKey = this.driveFolderCacheKey(folder.driveId, folder.id);

    // 1) Parent (AppComponent) may have already prefetched / warmed this folder.
    let cached = this.fileCrawlCache.get<any[]>(cacheKey);
    if (!Array.isArray(cached)) {
      const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
      if (hydrated && Array.isArray(hydrated.data)) {
        cached = hydrated.data;
      }
    }
    if (Array.isArray(cached)) {
      this.childFoldersByFolderId.set(
        folder.id,
        this.filterAllowedChildFolders(this.mapCachedChildrenToSpFiles(cached, folder), allowedNames)
      );
      this.refreshView();
      return;
    }

    this.loadingChildFolderIds.add(folder.id);
    this.refreshView();
    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token) return;

      const itemsResponse = await graphGetWithRetry(
        this.http,
        `/drives/${folder.driveId}/items/${folder.id}/children${DRIVE_CHILDREN_QUERY}`,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as GraphDriveItemsResponse;

      const rawItems = itemsResponse.value ?? [];

      // Warm the shared Attachments cache (same shape as AppComponent.fetchDriveFolderChildrenMapped).
      const mappedForCache = rawItems
        .filter((item: any) => !!item.file || !!item.folder)
        .map((item: any) => {
          const ext = item.file ? getFileExtension(item.name) : '';
          const category = item.file ? getFileCategory(ext) : '';
          const icon = item.file ? getFileIcon(category, ext) : '';
          return {
            id: item.id,
            name: item.name,
            webUrl: item.webUrl ?? '',
            parentId: folder.id,
            lastModifiedDateTime: item.lastModifiedDateTime,
            size: item.size,
            isFolder: !!item.folder,
            modifiedBy:
              item.lastModifiedBy?.user?.displayName ||
              item.lastModifiedBy?.user?.email ||
              undefined,
            fileExtension: ext,
            mimeType: item.file?.mimeType,
            fileCategory: category,
            fileIcon: icon,
          };
        });
      this.fileCrawlCache.set(cacheKey, mappedForCache);

      const children = rawItems
        .filter((item) => !!(item as { folder?: unknown }).folder)
        .map((item) => {
          const child = this.mapToSpFile(item, folder.driveId!);
          child.libraryName = folder.libraryName;
          child.parentFolderId = folder.id;
          child.parentFolderName = folder.name;
          return child;
        });
      const filtered = this.filterAllowedChildFolders(children, allowedNames);
      this.childFoldersByFolderId.set(folder.id, filtered);
    } catch (err) {
      console.error(`Failed to load subfolders of ${folder.name}:`, err);
    } finally {
      this.loadingChildFolderIds.delete(folder.id);
      this.refreshView();
    }
  }

  /** Parent's prefetch cache stores mapped items with `isFolder` — convert to SpFile. */
  private mapCachedChildrenToSpFiles(items: any[], parent: SpFile): SpFile[] {
    return items
      .filter((item) => !!item.isFolder || !!item.isfolder)
      .map((item) => ({
        id: item.id,
        name: item.name,
        webUrl: item.webUrl ?? '',
        isfolder: true,
        size: item.size,
        modifiedby: item.modifiedBy ?? item.modifiedby,
        modified: item.lastModifiedDateTime ?? item.modified,
        fileIcon: 'folder.png',
        fileCategory: 'folder',
        driveId: parent.driveId,
        libraryName: parent.libraryName,
        parentFolderId: parent.id,
        parentFolderName: parent.name,
      } as SpFile));
  }

  async loadFiles(): Promise<void> {
    this.isLoading = true;
    this.error = null;
    this.selectedFileId = null;
    this.filesNextLink = null;
    this.hasMoreFiles = false;
    this.isLoadingMoreFiles = false;

    // 1) Prefer per-library cache so switching libraries preserves Load more / nextLink.
    await this.fileCrawlCache.hydrateFromPersistent(this.libraryFilesCacheKey(this.selectedLibrary));
    const cached = this.readLibraryFilesCache(this.selectedLibrary);
    if (cached?.items?.length) {
      this.applyLibraryFilesToView(
        cached.items,
        cached.driveId,
        cached.complete ? null : cached.nextLink,
      );
      this.isLoading = false;
      this.refreshView();
      return;
    }

    // 2) All-root memory — only when that library was fully crawled (no missing pages).
    if (this.fullyLoadedAllLibraries.has(this.selectedLibrary)) {
      const fromAllRoot = this.allLibrariesFiles.filter(
        (f) => f.libraryName === this.selectedLibrary
      );
      if (fromAllRoot.length > 0) {
        const driveId = fromAllRoot.find((f) => f.driveId)?.driveId ?? null;
        this.applyLibraryFilesToView(fromAllRoot, driveId, null);
        this.writeLibraryFilesCache(this.selectedLibrary, fromAllRoot, null, driveId);
        this.isLoading = false;
        this.refreshView();
        return;
      }
    }

    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token) {
        throw new Error('Unable to acquire SharePoint token. Please ensure you are logged in.');
      }

      const drivesResponse = await this.getDrivesCached(token);
      const drive = drivesResponse.value?.find((d: any) =>
        d.name === this.selectedLibrary
      );

      if (!drive) {
        throw new Error(`Document library '${this.selectedLibrary}' not found`);
      }

      // First page only (100) — paint immediately like HR Files; more via Load more / search.
      const itemsResponse = await graphGetWithRetry(
        this.http,
        `/drives/${drive.id}/root/children${DRIVE_CHILDREN_QUERY}`,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as GraphDriveItemsResponse;

      const items = (itemsResponse.value ?? []).map((item) => this.mapToSpFile(item, drive.id));
      const nextLink = this.toNextPath(itemsResponse['@odata.nextLink']);
      this.applyLibraryFilesToView(items, drive.id, nextLink);
      this.writeLibraryFilesCache(this.selectedLibrary, items, nextLink, drive.id);
    } catch (err: any) {
      this.error = err.message || 'Failed to load files';
    } finally {
      this.isLoading = false;
      this.refreshView();
    }
  }

  /** Load the next page of files for the selected library (HR-style nextLink paging). */
  async loadMoreFiles(): Promise<void> {
    if (this.isLoadingMoreFiles || !this.hasMoreFiles || !this.filesNextLink) return;

    this.isLoadingMoreFiles = true;
    this.refreshView();
    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token) return;

      const path = this.toNextPath(this.filesNextLink) ?? this.filesNextLink;
      const itemsResponse = await graphGetWithRetry(
        this.http,
        path,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as GraphDriveItemsResponse;

      const driveId = this.filesCurrentDriveId ?? '';
      const more = (itemsResponse.value ?? []).map((item) => this.mapToSpFile(item, driveId));
      this.allFiles = [...this.allFiles, ...more];

      this.filesNextLink = this.toNextPath(itemsResponse['@odata.nextLink']);
      this.hasMoreFiles = !!this.filesNextLink;
      this.writeLibraryFilesCache(
        this.selectedLibrary,
        this.allFiles,
        this.filesNextLink,
        this.filesCurrentDriveId,
      );

      this.applyFilters();
    } catch (err) {
      console.error('Failed to load more files:', err);
    } finally {
      this.isLoadingMoreFiles = false;
      this.refreshView();
    }
  }

  private mapDriveItemToSpFile(item: GraphDriveItemsResponse['value'][number], driveId: string, libraryName: string): SpFile {
    const extension = getFileExtension(item.name);
    const category = getFileCategory(extension);
    const fields = item.listItem?.fields ?? {};
    return {
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      isfolder: !!(item as { folder?: unknown; }).folder,
      size: item.size,
      modifiedby: item.lastModifiedBy?.user?.displayName ||
        item.lastModifiedBy?.user?.email,
      modified: item.lastModifiedDateTime,
      fileIcon: getFileIcon(category, extension),
      fileCategory: category,
      fileExtension: extension,
      driveId,
      libraryName,
      createdBy: this.extractCreatedBy(item),
      assignedTo: this.extractAssignedTo(fields),
    };
  }

  private mapToSpFile(item: any, driveId: string): SpFile {
    const extension = getFileExtension(item.name);
    const category = getFileCategory(extension);
    const fields = item.listItem?.fields ?? {};
    return {
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      isfolder: !!item.folder,
      size: item.size,
      modifiedby: item.lastModifiedBy?.user?.displayName ||
        item.lastModifiedBy?.user?.email,
      modified: item.lastModifiedDateTime,
      fileIcon: getFileIcon(category, extension),
      fileCategory: category,
      fileExtension: extension,
      driveId: item.parentReference?.driveId ?? driveId,
      libraryName: this.selectedLibrary ?? undefined,
      createdBy: this.extractCreatedBy(item),
      assignedTo: this.extractAssignedTo(fields),
    };
  }

  private extractCreatedBy(item: {
    createdBy?: { user?: { displayName?: string; email?: string } };
  }): string | undefined {
    const name = item.createdBy?.user?.displayName?.trim()
      || item.createdBy?.user?.email?.trim();
    return name || undefined;
  }

  /** Primary "Assigned To" person column from the document library list item. */
  private extractAssignedTo(fields: Record<string, unknown>): string | undefined {
    const raw =
      fields['AssignedTo'] ??
      fields['Assigned'] ??
      fields['AssignedTo0'] ??
      fields['AssignedToPerson'];
    const name = this.extractPersonName(raw);
    return name || undefined;
  }

  private extractPersonName(val: unknown): string {
    if (!val) return '';
    if (typeof val === 'string') return val.trim();
    if (Array.isArray(val) && val.length > 0) {
      return val
        .map((p) => this.extractPersonName(p))
        .filter(Boolean)
        .join(', ');
    }
    if (typeof val === 'object') {
      // Prefer first non-empty field — empty LookupValue must not block Email/UPN.
      const obj = val as Record<string, unknown>;
      const nestedUser = (obj['user'] ?? obj['User'] ?? {}) as Record<string, unknown>;
      const candidates = [
        obj['LookupValue'],  //obj is the fields object from the SharePoint list item so what t hese are doing is looking for the value of the LookupValue field in the SharePoint list item
        obj['displayName'],
        obj['DisplayName'],
        obj['Title'],
        obj['EMail'],
        obj['Email'],
        obj['email'],
        obj['UserPrincipalName'],
        obj['userPrincipalName'],
        obj['name'],
        obj['Name'],
        nestedUser['displayName'],
        nestedUser['email'],
        nestedUser['userPrincipalName'],
      ];
      for (const candidate of candidates) {
        const text = String(candidate ?? '').trim();
        if (text) return text;
      }
      return '';
    }
    return '';
  }

  filterType: 'all' | 'files' | 'attachments' | 'comments' = 'all';
  readonly filterTypeOptions: AppDropdownOption[] = [
    { value: 'all', label: 'All' },
    { value: 'files', label: 'Files' },
    { value: 'attachments', label: 'Attachments' },
    { value: 'comments', label: 'Comments' },

  ];

  readonly approvalFilterOptions: AppDropdownOption[] = [
    { value: '', label: 'All statuses', tone: 'neutral' },
    { value: 'approved', label: 'Complete', tone: 'complete' },
    { value: 'pending', label: 'Pending', tone: 'pending' },
    { value: 'general-comments', label: 'General Comments', tone: 'comment' },
    { value: 'new-attachments', label: 'New Attachments', tone: 'attachment' },
    { value: 'rfa', label: 'Request For Action', tone: 'action' },
  ];

  onApprovalFilterChange(value: string): void {
    this.approvalFilterChange.emit(value ?? '');
  }

  get libraryDropdownOptions(): AppDropdownOption[] {
    return this.documentLibraries.map((lib: any) => ({
      value: lib.name,
      label: this.removeEformFromLibraryName(lib.displayName ?? lib.name),
    }));
  }

  onFilterTypeChange(type: string): void {
    this.filterType = type as 'all' | 'files' | 'attachments' | 'comments';
    this.resetRenderedPage();
    this.error = null; // a failure in one view must not hide the list in another
    this.graphSearchSeq++;
    this.graphSearchFolders = [];
    this.graphAttachmentHits = [];
    this.isGraphSearching = false;

    if (this.filterType === 'files') {
      void this.loadFiles().then(() => {
        if (this.searchTerm.trim()) {
          void this.continueFilesSearchInBackground();
          void this.searchFoldersViaGraph(this.searchTerm.trim(), 'files');
        }
      });
      return;
    }
    if (this.filterType === 'all') {
      this.applyFilters();
      if (this.searchTerm.trim()) {
        void this.searchAllLibraries(this.searchTerm.trim());
      }
      return;
    }
    if (this.filterType === 'attachments') {
      this.applyFilters();
      if (this.searchTerm.trim()) {
        void this.ensureAttachmentsForMatchingFolders(this.searchTerm.trim());
        void this.searchAttachmentsViaGraph(this.searchTerm.trim());
      }
      return;
    }
    if (this.filterType === 'comments') {
      this.applyFilters();
      if (this.searchTerm.trim()) {
        void this.ensureCommentsForMatchingFolders(this.searchTerm.trim());
        void this.continueCommentTasksSearchInBackground();
      }
      return;
    }
    this.applyFilters();
  }

  searchFiles(searchTerm: string): void {
    this.searchTerm = searchTerm ?? '';
    if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
    this.searchDebounceTimer = setTimeout(() => {
      this.searchDebounceTimer = null;
      this.resetRenderedPage();
      this.runSearch(this.searchTerm);
    }, SEARCH_DEBOUNCE_MS);
  }

  private runSearch(searchTerm: string): void {
    this.searchTerm = searchTerm ?? '';
    const term = this.searchTerm.trim();

    if (!term) {
      this.graphSearchSeq++;
      this.graphSearchFolders = [];
      this.graphAttachmentHits = [];
      this.isGraphSearching = false;
    }

    if (this.filterType === 'attachments') {
      this.graphAttachmentHits = [];
      this.applyFilters();
      if (term) {
        void this.ensureAttachmentsForMatchingFolders(term);
        void this.searchAttachmentsViaGraph(term);
      }
      return;
    }
    if (this.filterType === 'comments') {
      this.applyFilters();
      if (term) {
        void this.ensureCommentsForMatchingFolders(term);
        void this.continueCommentTasksSearchInBackground();
      }
      return;
    }
    if (this.filterType === 'files') {
      this.applyFilters();
      if (term) {
        void this.continueFilesSearchInBackground();
        void this.searchFoldersViaGraph(term, 'files');
      }
      return;
    }
    if (this.filterType === 'all') {
      this.applyFilters();
      if (term) {
        void this.searchAllLibraries(term);
      }
      return;
    }
    this.applyFilters();
  }

  /**
   * "All" loads nothing until the user searches. Graph search answers by folder
   * name straight away; the root-folder list (cached after the first time) is
   * loaded in the background so matches on Assigned To / library also appear.
   */
  private async searchAllLibraries(term: string): Promise<void> {
    void this.searchFoldersViaGraph(term, 'all');
    await this.loadAllLibrariesFiles();
    if (this.filterType === 'all' && this.searchTerm.trim()) {
      this.applyFilters();
    }
  }

  /** Clear Comments search UI; folder task caches are owned by the parent. */
  refreshCommentSearchCache(): void {
    this.commentHits = [];
    this.selectedCommentTaskId = null;
    if (this.filterType === 'comments') {
      this.applyFilters();
      if (this.searchTerm.trim()) {
        void this.ensureCommentsForMatchingFolders(this.searchTerm.trim());
        void this.continueCommentTasksSearchInBackground();
      }
    }
  }

  /** Called by parent after quietly caching folder Comments during search. */
  refreshCommentSearchFromCaches(): void {
    if (this.filterType !== 'comments' || !this.searchTerm.trim()) return;
    this.applyFilters();
    this.refreshView();
  }

  /** Called by parent after quietly caching folder Attachments during search. */
  refreshAttachmentSearchFromCaches(): void {
    if (this.filterType !== 'attachments' || !this.searchTerm.trim()) return;
    this.applyFilters();
    this.refreshView();
  }

  /** Reload the current All Files view (folders, library files, and search caches). */
  async refreshAllFilesData(force = true): Promise<void> {
    this.allLibraryTasks = [];
    this.filesNextLink = null;
    this.hasMoreFiles = false;
    this.commentTasksPendingLibraries = [];
    this.commentTasksCurrentNextLink = null;
    this.commentTasksCurrentLibrary = null;
    this.hasMoreCommentTasks = false;
    this.commentTasksWarmStarted = false;
    this.invalidateAllFilesCrawlCache();

    if (this.filterType === 'files') {
      await this.loadFiles();
      return;
    }

    // "All" with no search shows nothing, so there is nothing to reload.
    if (this.filterType === 'all' && !this.searchTerm.trim()) {
      this.applyFilters();
      return;
    }

    await this.loadAllLibrariesFiles(force);

    if (this.filterType === 'attachments' && this.searchTerm.trim()) {
      this.applyFilters();
      void this.ensureAttachmentsForMatchingFolders(this.searchTerm.trim());
      void this.searchAttachmentsViaGraph(this.searchTerm.trim());
      return;
    }

    if (this.filterType === 'comments' && this.searchTerm.trim()) {
      this.applyFilters();
      void this.ensureCommentsForMatchingFolders(this.searchTerm.trim());
      void this.continueCommentTasksSearchInBackground();
    }
  }

  private applyFilters(): void {
    const term = this.searchTerm.trim().toLowerCase();

    if (this.filterType !== 'comments') {
      this.commentHits = [];
      this.selectedCommentTaskId = null;
    }
    if (this.filterType !== 'attachments') {
      this.attachmentHits = [];
      this.selectedAttachmentFileId = null;
    }

    switch (this.filterType) {
      case 'all':
        this.applyAllFilter(term);
        break;
      case 'files':
        this.applyFilesFilter(term);
        break;
      case 'attachments':
        this.applyAttachmentsFilter(term);
        break;
      case 'comments':
        this.applyCommentsFilter(term);
        break;
    }
    this.emitFoldersForPrefetch();
    this.refreshView();
  }

  private emitFoldersForPrefetch(): void {
    // Don't compete with active Graph work — prefetch was a major 429 source.
    if (
      this.allLibrariesCrawlActive ||
      this.isLoadingAllLibraries ||
      this.isLoading ||
      this.isLoadingMoreFiles ||
      this.isSearchingMoreFiles ||
      this.isSearchingMoreCommentTasks ||
      this.isCrawlingCommentTasks ||
      this.isSearchingMoreAttachments ||
      this.commentTasksFetchInFlight ||
      this.isLoadingMoreCommentTasks
    ) {
      return;
    }

    const folders = this.files
      .filter((file) => file.isfolder && !!file.driveId)
      .slice(0, 8)
      .map((file) => ({ driveId: file.driveId!, folderId: file.id }));
    if (folders.length === 0) return;

    const signature = folders.map((f) => `${f.driveId}:${f.folderId}`).join('|');
    if (signature === this.lastPrefetchSignature) return;

    if (this.prefetchEmitTimer) clearTimeout(this.prefetchEmitTimer);
    this.prefetchEmitTimer = setTimeout(() => {
      this.prefetchEmitTimer = null;
      this.lastPrefetchSignature = signature;
      this.foldersForPrefetch.emit(folders);
    }, PREFETCH_EMIT_DEBOUNCE_MS);
  }

  /** All: show folders across every library (AllFiles order), filtered by folder reference. */
  private applyAllFilter(term: string): void {
    // Nothing is listed until the user searches (keeps startup fast).
    if (!term) {
      this.files = [];
      return;
    }
    let folders = this.allLibrariesFiles.filter(
      (file) => file.isfolder && this.hasFolderDisplayName(file)
    );
    if (term) {
      folders = folders.filter((folder) => this.folderMatchesTerm(folder, term));
      folders = this.mergeUniqueFolders(
        folders,
        this.graphSearchFolders.filter((f) => this.hasFolderDisplayName(f))
      );
    }
    this.files = this.sortByLibraryOrder(folders);
  }

  /** Files: show only root folders from the selected library (attachments live inside folders). */
  private applyFilesFilter(term: string): void {
    let items = this.allFiles.filter((file) => file.isfolder && this.hasFolderDisplayName(file));
    if (term) {
      items = items.filter((folder) => this.folderMatchesTerm(folder, term));
      const graphForLibrary = this.graphSearchFolders.filter(
        (f) =>
          this.hasFolderDisplayName(f) &&
          (!f.libraryName || f.libraryName === this.selectedLibrary)
      );
      items = this.mergeUniqueFolders(items, graphForLibrary);
    }
    this.files = items;
  }

  private hasFolderDisplayName(folder: SpFile): boolean {
    return String(folder?.name ?? '').trim().length > 0;
  }

  private mergeUniqueFolders(primary: SpFile[], extra: SpFile[]): SpFile[] {
    const map = new Map<string, SpFile>();
    for (const folder of [...primary, ...extra]) {
      if (!folder?.id) continue;
      const key = `${folder.libraryName ?? ''}:${folder.id}`;
      if (!map.has(key)) map.set(key, folder);
    }
    return [...map.values()];
  }

  /**
   * Server-side drive search so matches past the first loaded page still appear.
   * File hits resolve to their root folder so the list stays folder-only.
   */
  private async searchFoldersViaGraph(
    term: string,
    mode: 'files' | 'all',
  ): Promise<void> {
    const query = term.trim();
    if (query.length < 2) {
      this.graphSearchFolders = [];
      this.isGraphSearching = false;
      this.applyFilters();
      return;
    }

    const seq = ++this.graphSearchSeq;
    this.isGraphSearching = true;
    this.refreshView();

    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token || seq !== this.graphSearchSeq) return;

      const drivesResponse = await this.getDrivesCached(token);
      const drivesByName = new Map(
        (drivesResponse.value ?? []).map((d) => [d.name, d])
      );

      const libraries = mode === 'files'
        ? this.documentLibraries.filter((lib: { name: string }) => lib.name === this.selectedLibrary)
        : this.documentLibraries;

      const found: SpFile[] = [];
      for (const library of libraries) {
        if (seq !== this.graphSearchSeq) return;
        const drive = drivesByName.get(library.name);
        if (!drive) continue;

        const folders = await this.searchDriveFolders(drive.id, library.name, query, token);
        if (seq !== this.graphSearchSeq) return;
        found.push(...folders);
        // Progressive paint after each library.
        this.graphSearchFolders = this.mergeUniqueFolders([], found);
        this.applyFilters();
        this.refreshView();
        if (mode === 'all') await this.sleep(GRAPH_PAGE_GAP_MS);
      }

      if (seq !== this.graphSearchSeq) return;
      this.graphSearchFolders = this.mergeUniqueFolders([], found);
      this.applyFilters();
    } catch (err) {
      console.error('Graph folder search failed:', err);
    } finally {
      if (seq === this.graphSearchSeq) {
        this.isGraphSearching = false;
        this.refreshView();
      }
    }
  }

  private async searchDriveFolders(
    driveId: string,
    libraryName: string,
    query: string,
    token: string,
  ): Promise<SpFile[]> {
    const escaped = query.replace(/'/g, "''");
    const select =
      '$select=id,name,webUrl,folder,file,size,createdBy,lastModifiedBy,lastModifiedDateTime,parentReference';
    let nextPath: string | null =
      `/drives/${driveId}/root/search(q='${escaped}')?${select}&$top=50`;

    const folderMap = new Map<string, SpFile>();
    const rootFolderNamesNeeded = new Set<string>();
    let pages = 0;

    while (nextPath && pages < 3) {
      pages++;
      const response = await graphGetWithRetry(
        this.http,
        nextPath,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as GraphDriveItemsResponse;

      for (const item of response.value ?? []) {
        const itemName = String(item.name ?? '').trim();
        if (!itemName) continue;

        const rootName = this.extractRootFolderNameFromPath(item.parentReference?.path);
        const isFolder = !!item.folder;

        if (isFolder && !rootName) {
          // Root-level folder hit — only keep if the folder name actually matches the query.
          const mapped = this.mapDriveItemToSpFile(item, driveId, libraryName);
          mapped.isfolder = true;
          mapped.fileIcon = 'folder.png';
          mapped.fileCategory = 'folder';
          if (this.hasFolderDisplayName(mapped) && this.folderMatchesTerm(mapped, query)) {
            folderMap.set(mapped.id, mapped);
          }
          continue;
        }

        if (isFolder && rootName && rootName === itemName) {
          const mapped = this.mapDriveItemToSpFile(item, driveId, libraryName);
          mapped.isfolder = true;
          mapped.fileIcon = 'folder.png';
          mapped.fileCategory = 'folder';
          if (this.hasFolderDisplayName(mapped) && this.folderMatchesTerm(mapped, query)) {
            folderMap.set(mapped.id, mapped);
          }
          continue;
        }

        // File / nested hit → resolve its root folder (content matched the query).
        if (rootName?.trim()) {
          rootFolderNamesNeeded.add(rootName.trim());
        }
      }

      nextPath = this.toNextPath(response['@odata.nextLink']);
      if (nextPath) await this.sleep(GRAPH_PAGE_GAP_MS);
    }

    // Resolve root folders from local cache first, then Graph path lookup.
    for (const name of rootFolderNamesNeeded) {
      if (!name.trim()) continue;
      const existing =
        this.findCachedRootFolder(libraryName, name) ??
        [...folderMap.values()].find((f) => f.name === name);
      if (existing && this.hasFolderDisplayName(existing)) {
        folderMap.set(existing.id, existing);
        continue;
      }
      const fetched = await this.fetchRootFolderByName(driveId, libraryName, name, token);
      if (fetched && this.hasFolderDisplayName(fetched)) {
        folderMap.set(fetched.id, fetched);
      }
    }

    return [...folderMap.values()].filter((f) => this.hasFolderDisplayName(f));
  }

  private findCachedRootFolder(libraryName: string, folderName: string): SpFile | null {
    const matchName = (f: SpFile) =>
      f.isfolder &&
      f.name === folderName &&
      (!f.libraryName || f.libraryName === libraryName);

    return this.allFiles.find(matchName)
      ?? this.allLibrariesFiles.find(matchName)
      ?? null;
  }

  private async fetchRootFolderByName(
    driveId: string,
    libraryName: string,
    folderName: string,
    token: string,
  ): Promise<SpFile | null> {
    try {
      const encoded = folderName.split('/').map(encodeURIComponent).join('/');
      const item = await graphGetWithRetry(
        this.http,
        `/drives/${driveId}/root:/${encoded}?$select=id,name,webUrl,folder,file,size,createdBy,lastModifiedBy,lastModifiedDateTime,parentReference`,
        token,
        AppConstants.graphFileListingTimeoutMs
      );
      if (!item?.folder) return null;
      const mapped = this.mapDriveItemToSpFile(item, driveId, libraryName);
      mapped.isfolder = true;
      mapped.fileIcon = 'folder.png';
      mapped.fileCategory = 'folder';
      return mapped;
    } catch {
      return null;
    }
  }

  /** First path segment under /root: — the document-library root folder name. */
  private extractRootFolderNameFromPath(path: string | undefined): string | null {
    if (!path) return null;
    const marker = '/root:';
    const idx = path.indexOf(marker);
    if (idx < 0) return null;
    const rest = path.slice(idx + marker.length).replace(/^\//, '');
    if (!rest) return null;
    const first = rest.split('/').filter(Boolean)[0];
    if (!first) return null;
    try {
      return decodeURIComponent(first);
    } catch {
      return first;
    }
  }

  /**
   * Attachments: search files already loaded for All Files folders (folder contents
   * caches), same idea as Comments — Live-style attachment rows, not folder cards.
   */
  private applyAttachmentsFilter(term: string): void {
    if (!term) {
      this.files = [];
      this.attachmentHits = [];
      this.isSearchingMoreAttachments = false;
      return;
    }
    const localHits = this.getAttachmentHitsFromLoadedFolders(term);
    const seen = new Set(localHits.map((h) => h.fileId));
    const graphHits = this.graphAttachmentHits.filter((h) => h.fileId && !seen.has(h.fileId));
    this.attachmentHits = [...localHits, ...graphHits].sort((a, b) => this.compareAttachmentHits(a, b));
    this.files = [];
  }

  private compareAttachmentHits(a: AttachmentSearchHit, b: AttachmentSearchHit): number {
    const dateA = new Date(a.date || 0).getTime();
    const dateB = new Date(b.date || 0).getTime();
    if (dateA !== dateB) return dateB - dateA;
    return (a.fileName ?? '').localeCompare(b.fileName ?? '');
  }

  /**
   * Keyword search inside attachments: SharePoint drive search matches file names and
   * document text, so hits come from every folder — not only folders whose name matches.
   */
  private async searchAttachmentsViaGraph(term: string): Promise<void> {
    const query = term.trim();
    if (query.length < 2) return;

    const seq = ++this.graphSearchSeq;
    this.graphAttachmentHits = [];
    this.isGraphSearching = true;
    this.refreshView();

    try {
      const token = await this.authService.acquireSharePointToken();
      if (!token || seq !== this.graphSearchSeq) return;

      const drivesResponse = await this.getDrivesCached(token);
      const drivesByName = new Map(
        (drivesResponse.value ?? []).map((d) => [d.name, d])
      );

      for (const library of this.documentLibraries) {
        if (seq !== this.graphSearchSeq) return;
        const drive = drivesByName.get(library.name);
        if (!drive) continue;

        const hits = await this.searchDriveAttachments(drive.id, library.name, query, token, seq);
        if (seq !== this.graphSearchSeq) return;
        // Progressive paint after each library.
        this.graphAttachmentHits = [...this.graphAttachmentHits, ...hits];
        this.applyFilters();
        await this.sleep(GRAPH_PAGE_GAP_MS);
      }
    } catch (err) {
      console.error('Graph attachment search failed:', err);
    } finally {
      if (seq === this.graphSearchSeq) {
        this.isGraphSearching = false;
        this.refreshView();
      }
    }
  }

  private async searchDriveAttachments(
    driveId: string,
    libraryName: string,
    query: string,
    token: string,
    seq: number,
  ): Promise<AttachmentSearchHit[]> {
    const escaped = query.replace(/'/g, "''");
    const select =
      '$select=id,name,webUrl,folder,file,size,createdBy,lastModifiedBy,lastModifiedDateTime,parentReference';
    let nextPath: string | null =
      `/drives/${driveId}/root/search(q='${escaped}')?${select}&$top=50`;

    const fileItems: Array<{ item: GraphDriveItemsResponse['value'][number]; rootName: string }> = [];
    let pages = 0;
    while (nextPath && pages < ATTACHMENT_SEARCH_MAX_PAGES) {
      pages++;
      const response = await graphGetWithRetry(
        this.http,
        nextPath,
        token,
        AppConstants.graphFileListingTimeoutMs
      ) as GraphDriveItemsResponse;
      if (seq !== this.graphSearchSeq) return [];

      for (const item of response.value ?? []) {
        if (item.folder || !item.file) continue;
        const rootName = this.extractRootFolderNameFromPath(item.parentReference?.path)?.trim();
        if (rootName) fileItems.push({ item, rootName });
      }

      nextPath = this.toNextPath(response['@odata.nextLink']);
      if (nextPath) await this.sleep(GRAPH_PAGE_GAP_MS);
    }

    // Resolve each root folder once (local cache first, then Graph path lookup).
    const folderByName = new Map<string, SpFile | null>();
    const hits: AttachmentSearchHit[] = [];
    for (const { item, rootName } of fileItems) {
      if (seq !== this.graphSearchSeq) return [];
      if (!folderByName.has(rootName)) {
        const folder =
          this.findCachedRootFolder(libraryName, rootName) ??
          await this.fetchRootFolderByName(driveId, libraryName, rootName, token);
        folderByName.set(
          rootName,
          folder && this.hasFolderDisplayName(folder)
            ? { ...folder, driveId: folder.driveId || driveId, libraryName: folder.libraryName || libraryName }
            : null,
        );
      }
      const folder = folderByName.get(rootName);
      if (!folder) continue;
      const mapped = this.mapDriveItemToSpFile(item, driveId, libraryName);
      hits.push(this.toAttachmentSearchHit(mapped, folder, this.findRelatedCommentBodyForFolder(folder)));
    }
    return hits;
  }

  /**
   * Comments: search tasks already loaded for All Files folders (folder-task caches
   * + the open Comments panel), same idea as Live — not a separate 30k Graph crawl.
   */
  private applyCommentsFilter(term: string): void {
    if (!term) {
      this.files = [];
      this.commentHits = [];
      this.isSearchingMoreCommentTasks = false;
      return;
    }
    this.commentHits = this.getCommentHitsFromLoadedFolderComments(term);
    this.files = [];
  }

  /**
   * For All Files folders matching the query, ask the parent to quietly cache their
   * Attachments so search results fill in (same pattern as Comments search).
   */
  private async ensureAttachmentsForMatchingFolders(rawTerm: string): Promise<void> {
    const term = rawTerm.trim().toLowerCase();
    if (!term || this.filterType !== 'attachments') return;

    if (this.allLibrariesFiles.length === 0) {
      await this.loadAllLibrariesFiles();
    }

    const matchingFolders = this.allLibrariesFiles
      .filter(
        (f) =>
          f.isfolder &&
          this.hasFolderDisplayName(f) &&
          !!f.driveId &&
          this.folderMatchesTerm(f, term),
      )
      .slice(0, 20);

    this.isSearchingMoreAttachments = matchingFolders.length > 0;
    this.refreshView();

    try {
      for (const folder of matchingFolders) {
        if (this.filterType !== 'attachments' || this.searchTerm.trim().toLowerCase() !== term) {
          return;
        }
        const driveId = String(folder.driveId ?? '').trim();
        const folderId = String(folder.id ?? '').trim();
        if (!driveId || !folderId) continue;
        if (this.hasFolderAttachmentsCache(driveId, folderId)) continue;
        this.ensureFolderAttachments.emit({ driveId, folderId });
      }
    } finally {
      setTimeout(() => {
        if (this.filterType === 'attachments') {
          this.isSearchingMoreAttachments = false;
          this.applyFilters();
          this.refreshView();
        }
      }, 400);
    }
  }

  /**
   * For All Files folders matching the query, ask the parent to quietly cache their
   * Comments so search results fill in (Live searches comments on those files).
   */
  private async ensureCommentsForMatchingFolders(rawTerm: string): Promise<void> {
    const term = rawTerm.trim().toLowerCase();
    if (!term || this.filterType !== 'comments') return;

    if (this.allLibrariesFiles.length === 0) {
      await this.loadAllLibrariesFiles();
    }

    const matchingFolders = this.allLibrariesFiles
      .filter((f) => f.isfolder && this.hasFolderDisplayName(f) && this.folderMatchesTerm(f, term))
      .slice(0, 20);

    this.isSearchingMoreCommentTasks = matchingFolders.length > 0;
    this.refreshView();

    try {
      for (const folder of matchingFolders) {
        if (this.filterType !== 'comments' || this.searchTerm.trim().toLowerCase() !== term) {
          return;
        }
        const libraryName = String(folder.libraryName ?? '').trim();
        const folderName = String(folder.name ?? '').trim();
        if (!libraryName || !folderName) continue;
        if (this.hasFolderCommentsCache(libraryName, folderName)) continue;
        this.ensureFolderComments.emit({ libraryName, folderName });
      }
    } finally {
      // Parent refreshes hits as each folder lands; clear banner shortly after requests.
      setTimeout(() => {
        if (this.filterType === 'comments') {
          this.isSearchingMoreCommentTasks = false;
          this.applyFilters();
          this.refreshView();
        }
      }, 400);
    }
  }

  private hasFolderAttachmentsCache(driveId: string, folderId: string): boolean {
    const key = `files:folder:${driveId}:${folderId}`;
    const data = this.fileCrawlCache.getStale(key);
    return Array.isArray(data);
  }

  private parseDriveFolderCacheKey(key: string): { driveId: string; folderId: string } | null {
    const prefix = 'files:folder:';
    if (!key.startsWith(prefix)) return null;
    const rest = key.slice(prefix.length);
    const lastColon = rest.lastIndexOf(':');
    if (lastColon <= 0) return null;
    const driveId = rest.slice(0, lastColon).trim();
    const folderId = rest.slice(lastColon + 1).trim();
    if (!driveId || !folderId) return null;
    return { driveId, folderId };
  }

  /**
   * Live-style attachment rows: folder title + author/date + file link (+ optional body
   * from related folder comments). Includes files whose name matches, or every file in
   * a folder whose name matches the query.
   */
  private getAttachmentHitsFromLoadedFolders(term: string): AttachmentSearchHit[] {
    const query = term.trim().toLowerCase();
    if (!query) return [];

    const hits: AttachmentSearchHit[] = [];
    const seen = new Set<string>();

    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('files:folder:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      const parsed = this.parseDriveFolderCacheKey(key);
      if (!parsed) continue;

      const folder =
        this.allLibrariesFiles.find(
          (f) =>
            f.isfolder &&
            f.id === parsed.folderId &&
            (!f.driveId || f.driveId === parsed.driveId),
        ) ??
        this.allFiles.find(
          (f) =>
            f.isfolder &&
            f.id === parsed.folderId &&
            (!f.driveId || f.driveId === parsed.driveId),
        ) ??
        null;
      if (!folder || !this.hasFolderDisplayName(folder)) continue;

      const folderMatches = this.folderMatchesTerm(folder, query);
      const relatedBody = this.findRelatedCommentBodyForFolder(folder);

      for (const item of data) {
        if (item?.isFolder === true || item?.isfolder === true) continue;
        const fileName = String(item?.name ?? '').trim();
        if (!fileName) continue;

        const fileMatches = this.textMatchesQuery(fileName, query);
        if (!fileMatches && !folderMatches) continue;

        const fileId = String(item?.id ?? '').trim();
        if (!fileId || seen.has(fileId)) continue;
        seen.add(fileId);

        hits.push(this.toAttachmentSearchHit(item, folder, relatedBody));
      }
    }

    return hits.sort((a, b) => this.compareAttachmentHits(a, b));
  }

  private textMatchesQuery(value: string, query: string): boolean {
    const haystack = String(value ?? '').toLowerCase();
    if (!haystack || !query) return false;
    if (haystack.includes(query)) return true;
    const tokens = query.split(/[\s._@+/\\-]+/).filter((t) => t.length > 0);
    return tokens.length > 1 && tokens.every((token) => haystack.includes(token));
  }

  private toAttachmentSearchHit(
    item: any,
    folder: SpFile,
    relatedBody: string,
  ): AttachmentSearchHit {
    const fileName = String(item?.name ?? '').trim();
    const ext = getFileExtension(fileName);
    const category = getFileCategory(ext);
    const icon = String(item?.fileIcon ?? '').trim() || getFileIcon(category, ext) || 'files.png';
    return {
      fileId: String(item?.id ?? '').trim(),
      fileName,
      fileIcon: icon,
      webUrl: String(item?.webUrl ?? '').trim(),
      libraryName: String(folder.libraryName ?? '').trim(),
      title: folder.name,
      author: String(item?.modifiedBy ?? item?.modifiedby ?? folder.modifiedby ?? '').trim(),
      date: String(item?.lastModifiedDateTime ?? item?.modified ?? '').trim(),
      body: relatedBody,
      folder,
    };
  }

  /** Prefer a loaded folder comment body so Live-style cards can show a details section. */
  private findRelatedCommentBodyForFolder(folder: SpFile): string {
    const folderNorm = this.normalizeCacheKeyPart(folder.name);
    if (!folderNorm) return '';

    const tryItems = (items: any[]): string => {
      for (const item of items) {
        const body = this.stripHtmlToText(
          String(
            item?.eFormDetails?.comment ??
            item?.eFormDetails?.commentHtml ??
            item?.eFormDetails?.body ??
            item?.description ??
            '',
          ),
        );
        if (body) return body;
      }
      return '';
    };

    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('tasks:lib:v7:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      if (!key.endsWith(`:${folderNorm}`)) continue;
      const body = tryItems(data);
      if (body) return body;
    }
    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('tasks:hr-folder:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      if (!key.endsWith(`:${folderNorm}`)) continue;
      const body = tryItems(data);
      if (body) return body;
    }
    return '';
  }

  private hasFolderCommentsCache(libraryName: string, folderName: string): boolean {
    const folderNorm = this.normalizeCacheKeyPart(folderName);
    const libNorm = this.normalizeCacheKeyPart(libraryName);
    const taskList = this.documentLibraryTaskMap[libraryName];
    const listNorm = this.normalizeCacheKeyPart(taskList ?? '');

    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('tasks:lib:v7:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      if (!key.endsWith(`:${folderNorm}`)) continue;
      if (listNorm && key.includes(`:${listNorm}:`)) return true;
      if (key.includes(`:${libNorm}:`)) return true;
      // Folder name match is enough — list name varies by mapping.
      if (key.endsWith(`:${folderNorm}`)) return true;
    }
    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('tasks:hr-folder:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      if (key.endsWith(`:${folderNorm}`)) return true;
    }
    return false;
  }

  private getCommentHitsFromLoadedFolderComments(term: string): CommentSearchHit[] {
    const query = term.trim().toLowerCase();
    if (!query) return [];

    const hits: CommentSearchHit[] = [];
    const seen = new Set<string>();

    const addMapped = (item: any, folder: SpFile | null, folderTitle: string) => {
      if (!item || !this.mappedCommentMatchesTerm(item, query)) return;
      const taskId = String(item.id ?? '').trim();
      if (!taskId || seen.has(taskId)) return;
      const resolvedFolder =
        folder ??
        this.findFolderByNormalizedName(this.normalizeCacheKeyPart(folderTitle));
      if (!resolvedFolder) return;
      seen.add(taskId);
      hits.push(this.toCommentSearchHit(item, resolvedFolder));
    };

    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('tasks:lib:v7:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      const folderTitle = this.folderTitleFromTaskCacheKey(key, 'tasks:lib:v7:');
      if (!folderTitle) continue;
      const folder = this.findFolderByNormalizedName(folderTitle);
      for (const item of data) {
        addMapped(item, folder, folderTitle);
      }
    }

    for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix('tasks:hr-folder:')) {
      if (!Array.isArray(data) || data.length === 0) continue;
      const folderTitle = this.folderTitleFromTaskCacheKey(key, 'tasks:hr-folder:');
      if (!folderTitle) continue;
      const folder = this.findFolderByNormalizedName(folderTitle);
      for (const item of data) {
        addMapped(item, folder, folderTitle);
      }
    }

    // Also search the Comments panel items for the currently open folder.
    for (const item of this.commentItems ?? []) {
      const folder =
        this.findFolderByNormalizedName(
          this.normalizeCacheKeyPart(String((item as any)?.eFormDetails?.folderName ?? '')),
        ) ??
        this.files.find((f) => f.id === this.selectedFileId) ??
        this.allLibrariesFiles.find((f) => f.id === this.selectedFileId) ??
        null;
      if (!folder) {
        // Build a minimal folder from the selected file id if needed.
        const selected = this.allLibrariesFiles.find((f) => f.id === this.selectedFileId);
        if (selected) {
          addMapped(item, selected, selected.name);
        } else if (this.mappedCommentMatchesTerm(item, query)) {
          // Skip if we cannot resolve a folder to open.
        }
        continue;
      }
      addMapped(item, folder, folder.name);
    }

    // Keyword matches from the task-list crawl — covers folders that were never opened.
    let crawlHits = 0;
    for (const task of this.allLibraryTasks) {
      if (crawlHits >= COMMENT_SEARCH_CRAWL_HITS_MAX) break;
      const taskId = String(task.id ?? '').trim();
      if (!taskId || seen.has(taskId)) continue;
      if (!this.taskCommentMatchesTerm(task.fields, query)) continue;
      const folder = this.resolveFolderForCrawledTask(task);
      if (!folder) continue;
      seen.add(taskId);
      crawlHits++;
      hits.push({
        taskId,
        libraryName: String(folder.libraryName ?? task.libraryName ?? '').trim(),
        title: folder.name,
        author: this.extractCommentHitAuthor(task.fields),
        date: this.extractCommentHitDate(task.fields),
        body: this.extractCommentHitBody(task.fields),
        folder,
      });
    }

    return hits.sort((a, b) => {
      const dateA = new Date(a.date || 0).getTime();
      const dateB = new Date(b.date || 0).getTime();
      if (dateA !== dateB) return dateB - dateA;
      return (a.title ?? '').localeCompare(b.title ?? '');
    });
  }

  /**
   * Map a raw task-list row to its All Files folder: eForm id first, then folder/path
   * columns, then the person the folder is named after. Null when nothing matches.
   */
  private resolveFolderForCrawledTask(task: LibraryTaskItem): SpFile | null {
    this.ensureTaskFolderIndex();
    const f = task.fields;
    const pick = (candidates: SpFile[] | undefined): SpFile | null => {
      if (!candidates?.length) return null;
      return candidates.find((c) => c.libraryName === task.libraryName) ?? candidates[0];
    };

    const field11 = String(f['field_11'] ?? '').trim();
    const eFormId = String(
      f['eFormListId'] ?? f['ListId'] ?? (/^\d+$/.test(field11) ? field11 : '')
    ).trim() || extractTrailingFolderId(String(f['Title'] ?? f['Name'] ?? ''));
    const byId = eFormId ? pick(this.taskFolderIndexById.get(eFormId)) : null;
    if (byId) return byId;

    const folderRefs = ['Folder', 'FolderName', 'DocumentFolder', 'RelatedFolder', 'FileDirRef', 'FileLeafRef', 'Path', 'Title', 'Name'];
    for (const key of folderRefs) {
      const text = this.stringifyTaskFieldValue(f[key]);
      if (!text) continue;
      for (const segment of text.split(/[/\\]/)) {
        const hit = pick(this.taskFolderIndexByName.get(normalizeName(segment)));
        if (hit) return hit;
      }
    }

    const people = ['EmployeeName', 'Employee', 'Requestor', 'SubmittedBy', 'Submitter', 'CustomCreatedBy'];
    for (const key of people) {
      const name = this.extractPersonName(f[key]);
      const hit = name ? pick(this.taskFolderIndexByName.get(normalizeName(name))) : null;
      if (hit) return hit;
    }
    return null;
  }

  private ensureTaskFolderIndex(): void {
    const size = this.allLibrariesFiles.length + this.allFiles.length;
    if (
      size === this.taskFolderIndexSize &&
      this.allLibrariesFiles === this.taskFolderIndexLibs &&
      this.allFiles === this.taskFolderIndexFiles
    ) return;
    this.taskFolderIndexSize = size;
    this.taskFolderIndexLibs = this.allLibrariesFiles;
    this.taskFolderIndexFiles = this.allFiles;
    this.taskFolderIndexByName.clear();
    this.taskFolderIndexById.clear();
    this.folderIndexByKey.clear();

    const add = (map: Map<string, SpFile[]>, key: string, folder: SpFile) => {
      if (!key) return;
      const list = map.get(key);
      if (!list) map.set(key, [folder]);
      else if (!list.some((f) => f.id === folder.id)) list.push(folder);
    };
    for (const folder of [...this.allLibrariesFiles, ...this.allFiles]) {
      if (!folder.isfolder) continue;
      const key = this.normalizeCacheKeyPart(folder.name);
      if (key && !this.folderIndexByKey.has(key)) this.folderIndexByKey.set(key, folder);
      if (!this.hasFolderDisplayName(folder)) continue;
      add(this.taskFolderIndexByName, normalizeName(folder.name), folder);
      add(this.taskFolderIndexById, extractTrailingFolderId(folder.name), folder);
    }
  }

  private folderTitleFromTaskCacheKey(key: string, prefix: string): string {
    if (!key.startsWith(prefix)) return '';
    const rest = key.slice(prefix.length);
    // lib keys: list:folder — hr keys: v6:folder or folder
    const parts = rest.split(':').filter(Boolean);
    return parts.length ? parts[parts.length - 1] : '';
  }

  private findFolderByNormalizedName(normalizedFolderName: string): SpFile | null {
    const target = this.normalizeCacheKeyPart(normalizedFolderName);
    if (!target) return null;
    this.ensureTaskFolderIndex();
    return this.folderIndexByKey.get(target) ?? null;
  }

  /** Build search text once per object; later searches reuse it. */
  private cachedHaystack(source: unknown, build: () => string): string {
    if (!source || typeof source !== 'object') return build();
    let haystack = this.commentHaystackCache.get(source);
    if (haystack === undefined) {
      haystack = build();
      this.commentHaystackCache.set(source, haystack);
    }
    return haystack;
  }

  private mappedCommentMatchesTerm(item: any, term: string): boolean {
    const haystack = this.cachedHaystack(item, () => this.buildMappedCommentHaystack(item));
    if (!haystack) return false;
    if (haystack.includes(term)) return true;
    const tokens = term.split(/[\s._@+/\\-]+/).filter((t) => t.length > 0);
    if (tokens.length <= 1) return false;
    return tokens.every((token) => haystack.includes(token));
  }

  private buildMappedCommentHaystack(item: any): string {
    const details = item?.eFormDetails ?? {};
    return [
      item?.name,
      item?.description,
      item?.submittedBy,
      item?.status,
      details?.type,
      details?.status,
      details?.submitter,
      details?.listName,
      details?.comment,
      details?.commentHtml,
      details?.body,
      details?.reason,
      details?.employeeName,
      details?.approver1,
      details?.approver1Comment,
      details?.approver2Comment,
      details?.approver3Comment,
      details?.category,
      details?.eFormListId,
    ]
      .map((v) => this.stripHtmlToText(String(v ?? '')).toLowerCase())
      .filter(Boolean)
      .join(' ');
  }

  private toCommentSearchHit(item: any, folder: SpFile): CommentSearchHit {
    const details = item?.eFormDetails ?? {};
    const body = this.stripHtmlToText(
      String(
        details.comment ??
        details.commentHtml ??
        details.body ??
        item?.description ??
        item?.name ??
        '',
      ),
    );
    const author = String(
      item?.submittedBy ?? details.submitter ?? details.customModifiedBy ?? '',
    ).trim();
    const date = String(
      details.customCreatedDate ??
      item?.submittedDate ??
      details.submittedDate ??
      item?.lastModifiedDateTime ??
      '',
    ).trim();

    return {
      taskId: String(item?.id ?? '').trim(),
      libraryName: String(folder.libraryName ?? details.listName ?? '').trim(),
      title: folder.name,
      author,
      date,
      body,
      folder,
    };
  }

  private folderMatchesTerm(folder: SpFile, term: string): boolean {
    const query = term.trim().toLowerCase();
    if (!query) return true;

    const libraryDisplayName =
      this.documentLibraries.find((lib: { name: string }) => lib.name === folder.libraryName)
        ?.displayName ?? '';

    const haystackParts = [
      folder.name,
      folder.modifiedby,
      folder.assignedTo,
      folder.createdBy,
      libraryDisplayName,
      folder.libraryName,
    ]
      .map((v) => String(v ?? '').toLowerCase().trim())
      .filter(Boolean);

    const haystack = haystackParts.join(' ');
    const normalizedHaystack = this.normalizeMatchText(haystack);
    const normalizedTerm = this.normalizeMatchText(query);

    // Exact substring / normalized substring — search text must appear in the folder.
    if (haystack.includes(query)) return true;
    if (normalizedTerm.length >= 2 && normalizedHaystack.includes(normalizedTerm)) return true;

    // Multi-word: every search token must appear in the folder (AND).
    // Never use term.includes(token) — that made "8063" match folders with "3" / "2".
    const tokens = query.split(/[\s._@+/\\-]+/).filter((t) => t.length > 0);
    if (tokens.length <= 1) {
      const nameTokens = (folder.name ?? '')
        .toLowerCase()
        .split(/[\s._@+/\\-]+/)
        .filter((t) => t.length > 0);
      return nameTokens.some((token) => {
        if (token.includes(query)) return true;
        const normToken = this.normalizeMatchText(token);
        return normalizedTerm.length >= 2 && normToken.includes(normalizedTerm);
      });
    }

    return tokens.every((token) => {
      if (haystack.includes(token)) return true;
      const normToken = this.normalizeMatchText(token);
      return normToken.length >= 2 && normalizedHaystack.includes(normToken);
    });
  }

  private fileMatchesTerm(file: SpFile, term: string): boolean {
    const name = file.name?.toLowerCase() ?? '';
    const modifiedBy = file.modifiedby?.toLowerCase() ?? '';
    return name.includes(term) || modifiedBy.includes(term) || this.fileCreatorMatchesTerm(file, term);
  }

  private fileCreatorMatchesTerm(file: SpFile, term: string): boolean {
    return (file.assignedTo?.toLowerCase() ?? '').includes(term)
      || (file.createdBy?.toLowerCase() ?? '').includes(term);
  }

  private sortByLibraryOrder(folders: SpFile[]): SpFile[] {
    const orderMap = new Map(this.documentLibraries.map((lib: { name: string }, index: number) => [lib.name, index]));
    return [...folders].sort((a, b) => {
      const orderA = orderMap.get(a.libraryName ?? '') ?? 999;
      const orderB = orderMap.get(b.libraryName ?? '') ?? 999;
      if (orderA !== orderB) return orderA - orderB;
      return (a.name ?? '').localeCompare(b.name ?? '');
    });
  }

  private taskCommentMatchesTerm(fields: Record<string, unknown>, term: string): boolean {
    const haystack = this.cachedHaystack(fields, () => this.buildTaskCommentHaystack(fields));
    if (!haystack) return false;

    const query = term.trim().toLowerCase();
    if (!query) return true;
    if (haystack.includes(query)) return true;

    const tokens = query.split(/[\s._@+/\\-]+/).filter((t) => t.length > 0);
    if (tokens.length <= 1) return false;
    return tokens.every((token) => haystack.includes(token));
  }

  private buildTaskCommentHaystack(fields: Record<string, unknown>): string {
    // Prefer comment/body + people fields. Title/Name are included so task titles still match,
    // but folder cards are no longer what we display or open.
    const haystacks = [
      fields['EmployeeName'],
      fields['Employee'],
      fields['Requestor'],
      fields['SubmittedBy'],
      fields['Submitter'],
      fields['Author'],
      fields['CreatedBy'],
      fields['CustomCreatedBy'],
      fields['Comment'],
      fields['Comments'],
      fields['Notes'],
      fields['Description'],
      fields['Body'],
      fields['RequestorComment'],
      fields['ApproverComment'],
      fields['Approver1Comment'],
      fields['Approver2Comment'],
      fields['Approver3Comment'],
      fields['field_10'],
      fields['Title'],
      fields['Name'],
    ];
    return haystacks
      .map((value) => this.stringifyTaskFieldValue(value).toLowerCase())
      .filter(Boolean)
      .join(' ');
  }

  private stripHtmlToText(value: string): string {
    return String(value ?? '')
      .replace(/&nbsp;|&#160;/gi, ' ')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]*>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private extractCommentHitAuthor(fields: Record<string, unknown>): string {
    const candidates = [
      fields['CustomCreatedBy'],
      fields['Requestor'],
      fields['SubmittedBy'],
      fields['Submitter'],
      fields['EmployeeName'],
      fields['Employee'],
      fields['Author'],
      fields['CreatedBy'],
      fields['AssignedTo'],
    ];
    for (const candidate of candidates) {
      const name = this.extractPersonName(candidate);
      if (name) return name;
    }
    return '';
  }

  private extractCommentHitDate(fields: Record<string, unknown>): string {
    return String(
      fields['CustomCreatedDate'] ??
      fields['Created'] ??
      fields['SubmittedDate'] ??
      fields['CustomModifiedDate'] ??
      fields['Modified'] ??
      ''
    ).trim();
  }

  private extractCommentHitBody(fields: Record<string, unknown>): string {
    const raw = String(
      fields['Comment'] ??
      fields['Comments'] ??
      fields['Notes'] ??
      fields['field_10'] ??
      fields['Description'] ??
      fields['Body'] ??
      fields['Title'] ??
      fields['Name'] ??
      ''
    ).trim();
    return this.stripHtmlToText(raw);
  }

  private stringifyTaskFieldValue(value: unknown): string {
    if (value == null) return '';
    if (Array.isArray(value)) {
      return value.map((v) => this.stringifyTaskFieldValue(v)).filter(Boolean).join(' ');
    }
    if (typeof value === 'object') {
      const record = value as Record<string, unknown>;
      return String(
        record['LookupValue'] ??
        record['Title'] ??
        record['displayName'] ??
        record['email'] ??
        record['name'] ??
        record['value'] ??
        ''
      );
    }
    return String(value);
  }

  private normalizeMatchText(value: string): string {
    return value.toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  private getFolderMatchTokens(folderName: string): string[] {
    const raw = String(folderName ?? '').trim().toLowerCase();
    if (!raw) return [];

    const withoutLeadingNumber = raw.replace(/^\d+\s+/, '').trim();
    const employeeId = raw.match(/^\d+/)?.[0] ?? '';
    const emailMatch = raw.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i)?.[0] ?? '';
    const emailLocal = emailMatch ? emailMatch.split('@')[0] : '';
    const displayFromEmail = emailLocal.replace(/[._-]+/g, ' ');

    return [...new Set([
      withoutLeadingNumber,
      emailMatch,
      emailLocal,
      displayFromEmail,
      employeeId,
      raw,
    ]
      .map((token) => this.normalizeMatchText(token))
      .filter((token) => token.length >= 2))];
  }

  private extractFolderCandidatesFromTask(fields: Record<string, unknown>): string[] {
    const rawValues = [
      fields['Folder'],
      fields['FolderName'],
      fields['DocumentFolder'],
      fields['RelatedFolder'],
      fields['Title'],
      fields['Name'],
      fields['FileLeafRef'],
      fields['FileDirRef'],
      fields['Path'],
      fields['Requestor'],
      fields['SubmittedBy'],
      fields['EmployeeName'],
      fields['Employee'],
      fields['eFormListId'],
      fields['ListId'],
      fields['field_11'],
    ];

    const names: string[] = [];
    for (const val of rawValues) {
      if (val == null || val === '') continue;
      const text = this.stringifyTaskFieldValue(val);
      names.push(text);
      for (const segment of text.split(/[/\\]/)) {
        const trimmed = segment.trim();
        if (trimmed) names.push(trimmed);
      }
    }
    return names;
  }

  private resolveFolderForTask(libraryName: string, fields: Record<string, unknown>): SpFile | null {
    const libraryFolders = this.allLibrariesFiles.filter(
      (f) => f.isfolder && f.libraryName === libraryName
    );
    if (libraryFolders.length === 0) return null;

    const candidates = this.extractFolderCandidatesFromTask(fields);
    for (const folder of libraryFolders) {
      const folderTokens = this.getFolderMatchTokens(folder.name);
      if (folderTokens.length === 0) continue;

      for (const candidate of candidates) {
        const normalizedCandidate = this.normalizeMatchText(candidate);
        if (!normalizedCandidate) continue;
        if (folderTokens.some((token) =>
          normalizedCandidate.includes(token) || token.includes(normalizedCandidate)
        )) {
          return folder;
        }
      }
    }

    const title = String(fields['Title'] ?? fields['Name'] ?? '').trim().toLowerCase();
    if (title) {
      const match = libraryFolders.find((folder) => {
        const folderName = folder.name.toLowerCase();
        return folderName.includes(title) || title.includes(folderName);
      });
      if (match) return match;
    }

    return null;
  }

  /** Display only: "Procurement eForm" -> "Procurement". The real library name is left untouched. */
  private removeEformFromLibraryName(libraryName: string): string {
    const label = libraryName.replace(/\be-?forms?\b/gi, '').replace(/\s{2,}/g, ' ').trim();
    return label || libraryName;
  }

  onLibraryChange(libraryName: string): void {
    this.selectedLibrary = libraryName;
    this.resetRenderedPage();
    this.graphSearchSeq++;
    this.graphSearchFolders = [];
    this.isGraphSearching = false;
    void this.loadFiles().then(() => {
      if (this.filterType === 'files' && this.searchTerm.trim()) {
        void this.searchFoldersViaGraph(this.searchTerm.trim(), 'files');
      }
    });
  }

  formatFileSize(bytes: number): string {
    if (!bytes) return '';
    const units = ['B', 'KB', 'MB', 'GB'];
    const size = Math.floor(Math.log(bytes) / Math.log(1024));
    return `${(bytes / Math.pow(1024, size)).toFixed(1)} ${units[size]}`;
  }
}