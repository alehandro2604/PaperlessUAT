// ============================================================
// IMPORTS & DEPENDENCIES
// ============================================================
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ChangeDetectorRef, Component, OnInit, OnDestroy, ViewChild, HostListener } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { firstValueFrom, Observable, Subscription } from 'rxjs';
import { sharePointConfig } from './sharepoint.config';
import { AppConstants } from './app.constants';
import {
  HR_TASK_LIST_NAMES, HR_SOURCE_EFORM_LIST_NAMES, TODO_EXTRA_TASK_LIST_NAMES,
  isTodoProcurementTaskList, isHrSourceEFormList,
  isHrTitleMatchedTaskList,
} from './hr-task-lists.config';
import { getFileExtension, getFileCategory, getFileIcon, normalizeSharePointFileUrl } from './file-utils';
import { graphGet, graphGetWithRetry, clearGraphThrottleCooldown, toGraphPath, normalizeName } from './microsoft-graph';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { AuthService } from './services/auth.service';
import { FormConfigurationService } from './services/form-configuration.service';
import { EForm } from './models/form-configuration.model';
import {
  MobileCardView, MobileNavigationComponent, MOBILE_NAV_INDEX, getAllFilesTabMobileState, getAttachmentsTabMobileState, getCommentsTabMobileState, getTaskTabMobileState, getTodoTabMobileState, resolveMobileNavClick,
} from './components/mobile navigation/mobile-navigation';
import {
  resolveMobileSearchChange, syncMobileSearchTerm,
} from './components/mobile navigation/mobile-search.utils';
import { TodoListComponent, SubmitterSelection } from './components/todo-list/todo-list.component';
import { AllFilesComponent, AttachmentSearchHit, CommentSearchHit } from './components/All-files/all-files';
import { TodoService } from './services/todo.service';
import { UserService } from './services/user.service';
import { DelegateService } from './services/delegate.service';
import { CommentService } from './services/comment.service';
import { TaskService } from './services/task.service';
import { SubordinaryTaskService } from './services/subordinaryTask.service';
import { FileCrawlCacheService } from './services/file-crawl.service';
import { SiteMetadataService } from './services/site-metadata.service';
import { BackendApiService } from './services/backend-api.service';
import { HrTaskMapperService } from './services/hr-task-mapper.service';
import { HrTaskFolderMatchService } from './services/hr-task-folder-match.service';
import { SharePointTaskQueryService } from './services/sharepoint-task-query.service';

// ============================================================
// ANGULAR COMPONENT DECORATOR
// ============================================================
import { PowerappsModalComponent } from './components/powerApps/powerapps-modal.component';
import { NewCommentComponent } from './components/new-comment/new-comment';
import { LoadingScreenComponent } from './components/loading screen/loading-screen';
import { DelegateComponent } from './components/delegate-component/delegate';
import { ManualRefreshComponent } from './components/manual-refresh/refresh';
import { AppDropdownComponent, AppDropdownOption } from './components/app-dropdown/app-dropdown.component';
import { HrFilesComponent } from './components/hr-files/hr-files.component';
import { NewCommentSavedEvent, CommentListItem } from './models/comment.models';
import { buildCommentCardFromSavedEvent } from './utils/comment-card-builders';
import { formatFormTitle } from './form-utils';
import {
  isSyntheticCommentCard, matchesCommentLabel, isCommentEntryTitle, isStatuslessCommentTitle,
  getCommentEntryKind, shouldHideCommentStatus, isCommentTypeItem, isSharePointCommentEntry,
  extractCommentCardTaskId, matchesApprovalFilter, escapeHtml, buildCommentHtmlWithAttachment,
  extractAttachmentFromCommentHtml, openAttachmentLink, onCommentLinkClick, buildQuickCommentItemFromSearchHit,
} from './utils/comment-utils';
import {
  isItemSubmittedBy, getHrTaskPersonLookupIds, doesHrTaskMatchPersonLookupId, normalizeHrTaskChainName,
  getMappedHrTaskChainName, getRawHrTaskChainName, hrTaskChainNamesMatch, doesRawItemReferenceEFormKeys,
  doesRawItemMatchSeedEFormIdAndTaskName, extractEFormKeysFromMappedItem, doesMappedItemMatchSeedEFormIdAndTaskName,
  collectEFormKeyToTaskNames, buildPersonLookupFilterExpr, isDocumentLibraryTaskList, isChildSubjectColumn,
  getDateSearchTokens, getTaskEFormTitle, extractTrailingFolderId,
} from './utils/hr-task-matching';
import {
  cleanParsedPersonName, extractPersonName, getEmailFromHrFolderName, normalizeTaskMatchText,
  fieldMatchesFolderToken, extractPersonNameParts, getHrPersonalFolderMatchTokens, candidatesMatchFolderPerson,
  isItemSubmittedByFolderPerson, collectTaskAssigneeValues, stringifyTaskFieldValue,
} from './utils/person-name-matching';

type HrTaskDetailTone = 'success' | 'warning' | 'danger' | 'neutral';

/** Sectioned view of an HR task, rendered by the details panel. */
interface HrTaskDetailView {
  title: string;
  subtitle: string;
  status: { value: string; tone: HrTaskDetailTone } | null;
  statuses: { label: string; value: string; tone: HrTaskDetailTone }[];
  references: { label: string; value: string }[];
  people: { role: string; name: string; initials: string; isMe: boolean }[];
  details: { label: string; value: string }[];
  approvals: { label: string; name: string; date: string; comment: string }[];
  timeline: { label: string; value: string }[];
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, TodoListComponent, AllFilesComponent, PowerappsModalComponent, NewCommentComponent, LoadingScreenComponent, MobileNavigationComponent, DelegateComponent, ManualRefreshComponent, AppDropdownComponent, HrFilesComponent],
  templateUrl: './app.html',
  styleUrls: ['./app.css', './components/attachments/attachments.css', './form-list.css', './components/mobile navigation/mobile-navigation.css', './components/task-details/task-details.css', './components/comments/comment-card.css'],
})
export class AppComponent implements OnInit, OnDestroy {
  @ViewChild(AllFilesComponent) private allFilesComponent?: AllFilesComponent;
  @ViewChild(TodoListComponent) private todoListComponent?: TodoListComponent;

  /** Initials for the signed-in user's avatar in the top bar. */
  protected getUserInitials(name: string | null | undefined): string {
    const base = String(name ?? '').split('@')[0].replace(/[._-]+/g, ' ');
    const initials = base.split(/\s+/).filter(Boolean).map(w => w[0]).join('');
    return (initials || '?').slice(0, 3).toUpperCase();
  }

  protected isUserMenuOpen = false;

  protected toggleUserMenu(event: Event): void {
    event.stopPropagation();
    this.isUserMenuOpen = !this.isUserMenuOpen;
  }

  /** Close the menu when the user clicks anywhere else on the page or presses Escape. */
  @HostListener('document:click')
  @HostListener('document:keydown.escape')
  protected closeUserMenu(): void {
    this.isUserMenuOpen = false;
  }



  constructor(
    private readonly http: HttpClient,
    private readonly cdr: ChangeDetectorRef,
    private readonly todoService: TodoService,
    private readonly userService: UserService,
    private readonly authService: AuthService,
    private readonly delegateService: DelegateService,
    private readonly commentService: CommentService,
    private readonly formConfigService: FormConfigurationService,
    private readonly taskService: TaskService,
    private readonly subordinaryTaskService: SubordinaryTaskService,
    private readonly fileCrawlCache: FileCrawlCacheService,
    private readonly siteMetadataService: SiteMetadataService,
    private readonly backendApi: BackendApiService,
    private readonly hrTaskMapper: HrTaskMapperService,
    private readonly folderMatch: HrTaskFolderMatchService,
    private readonly taskQuery: SharePointTaskQueryService,
  ) { }

  private static readonly HR_USER_TASKS_CACHE_KEY = 'tasks:hr-user:v5';
  /** Persisted HR Files people-folder list so the first open paints from cache. */
  private static readonly HR_ROOT_FOLDERS_CACHE_KEY = 'files:hr-root:v1';
  /** Pause between batches of HR task-list fetches so SharePoint throttling (429) stays rare. */
  private static readonly HR_TASK_LIST_BATCH_GAP_MS = 400;
  /**
   * Background poll for To Do watermarks. Folder Comments uses a longer cadence
   * (FOLDER_COMMENTS_POLL_MS) so a person with many lists is not re-fanned every tick.
   */
  private static readonly SHAREPOINT_CACHE_POLL_MS = 60_000;
  /** Idle top-up of the open HR/All Files Comments folder. Keep ≥3 min to avoid 429s. */
  private static readonly FOLDER_COMMENTS_POLL_MS = 3 * 60_000;

  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private normalizeCacheKeyPart(value: string): string {
    return String(value ?? '').trim().toLowerCase();
  }

  private hrFolderTaskCacheKey(folderName: string): string {
    // v9: only finished crawls are persisted — drops older snapshots that may be partial.
    return `tasks:hr-folder:v9:${this.normalizeCacheKeyPart(folderName)}`;
  }

  /** Shared-Redis `{ email, lookupId }` for an HR person folder (under the shared hr-folder prefix). */
  private hrPersonIdsCacheKey(folderName: string): string {
    return `tasks:hr-folder:ids:${this.normalizeCacheKeyPart(folderName)}`;
  }

  private libFolderTaskCacheKey(listName: string, folderName: string): string {
    return `tasks:lib:v7:${this.normalizeCacheKeyPart(listName)}:${this.normalizeCacheKeyPart(folderName)}`;
  }

  private driveFolderCacheKey(driveId: string, folderId: string): string {
    return `files:folder:${driveId}:${folderId}`;
  }

  private invalidateFolderTaskCache(libraryName: string, folderName: string): void {
    const folder = String(folderName || '').trim();
    if (!folder) return;
    if (normalizeName(libraryName) === normalizeName(this.targetLibraryName)) {
      const key = this.hrFolderTaskCacheKey(folder);
      this.fileCrawlCache.invalidate(key);
      this.hrFolderCompleteAt.delete(key);
      this.hrFolderPartialKeys.delete(key);
      return;
    }
    const mappedTaskList = this.getTaskListForLibrary(libraryName);
    if (mappedTaskList) {
      this.fileCrawlCache.invalidate(this.libFolderTaskCacheKey(mappedTaskList, folder));
    }
  }

  private applyCachedFolderTaskItems(cached: any[], folderName: string, pendingRefresh = false): void {
    if (pendingRefresh) {
      this.isLoadingMoreComments = true;
      this.commentsMessage = '';
    } else {
      this.isLoadingMoreComments = false;
    }
    this.publishFolderTaskProgress(
      this.sortCommentItemsByDateDesc(this.applyAllFilesFolderSubmitter(cached)),
      {
        done: !pendingRefresh,
        emptyMessage: 'No tasks found.',
      },
    );
  }





  // this is to check if the item is a user comment card
  protected isUserCommentCard(item: { id?: unknown }): boolean {
    return String(item?.id ?? '').startsWith('comment:');
  }

  // ============================================================
  // UI STATE
  // ============================================================
  protected isConnecting = false;
  protected statusMessage = 'Checking Microsoft 365 sign-in...';
  protected commentsMessage = 'Loading comments...';
  protected errorMessage = '';
  protected showUserFile = false;
  protected isLoadingUserFiles = false;
  /** True while the HR Files folder list in the first section is loading - separate from attachments. */
  protected isLoadingHrFilesList = false;
  protected isLoadingComments = false;
  /** True while a Comments network fetch is still running after the first page is shown. */
  protected isLoadingMoreComments = false;
  protected showToDoSection = true;
  protected isLoadingTodoTasks = false;
  /** Prevent stale To Do loads when the tab or subordinate toggle is used rapidly. */
  private todoTaskLoadSeq = 0;
  /**
   * True after the first To Do Graph pass (assignee + blocked-list fallback) finishes.
   * False while still loading or after a mid-load cancel - used to decide resume vs full reload.
   */
  private todoFirstPassComplete = false;
  /** True while Progress=Pending procurement pages are draining in the background. */
  protected todoProcurementDrainActive = false;
  /** Prevent stale Comments/HR task loads when task groups are clicked rapidly. */
  private hrCommentsLoadSeq = 0;
  /** Count of in-flight HR / All Files folder Comments Graph crawls. */
  private folderTaskGraphInFlight = 0;
  private userHrTasksLoaded = false;
  protected showAllFilesSection = false;
  /** Once true, All Files stays in the DOM (hidden) so listings/search caches survive tab switches. */
  protected allFilesMounted = false;
  protected searchTerm = '';
  protected hideComments = true;
  protected userFileProgressMessage = '';
  protected userFileError = '';
  protected userFileWarning = '';
  protected mobileCardView: MobileCardView = 'todo';
  protected currentFolderId: string | null = null;
  protected currentFolderWebUrl = '';
  protected currentFolderName = '';
  protected folderStack: Array<{ id: string; name: string; webUrl: string }> = [];
  /** Top-level folder opened from HR Files / All Files - back never goes above this. */
  private attachmentBrowsingRoot: { id: string; name: string; webUrl: string } | null = null;
  protected currentLibraryDriveId: string | null = null;
  protected currentLibraryName: string | null = null;
  protected selectedFolderName = '';
  private selectedAllFilesFolderId: string | null = null;
  /** Parent folder name when a level-2 ChildSubject folder is selected (for task-cache reuse). */
  private selectedAllFilesParentFolderName = '';
  private childSubjectSourceFetchedForFolderId: string | null = null;
  /** ChildSubject / ChildSubject2 names keyed by library + trailing eForm id. */
  private readonly childSubjectNamesByLibraryEFormId = new Map<string, string[]>();
  /** SharePoint lastModifiedBy for the folder selected in All Files (shown as Submitted by). */
  protected selectedFolderModifiedBy = '';
  /** After Comments-search click: scroll/highlight this task id in the Comments panel. */
  private pendingCommentFocusId: string | null = null;
  protected highlightedCommentId: string | null = null;
  /** After Attachments-search click: highlight this file id in the Attachments panel. */
  private pendingAttachmentFocusId: string | null = null;
  protected highlightedAttachmentId: string | null = null;

  /** Library name ? task list name, sourced from the eForms list via FormConfigurationService. */
  private documentLibraryTaskMap: Record<string, string> = {};
  private documentLibraryTaskMapSub?: Subscription;

  // ============================================================
  // HR TASK LIST SELECTION
  // ============================================================
  protected allowedHrTaskLists: string[] = [];
  protected selectedHrTaskList: string = '';
  protected get isUsingHrTaskList(): boolean { return this.selectedHrTaskList !== ''; }

  // ============================================================
  // ATTACHMENTS / DRIVE FILES
  // ============================================================
  protected userFiles: Array<{
    id: string;
    name: string;
    webUrl: string;
    parentId?: string;
    lastModifiedDateTime?: string;
    size?: number;
    isFolder: boolean;
    modifiedBy?: string;
    fileExtension?: string;
    mimeType?: string;
    fileCategory?: string;
    fileIcon?: string;
  }> = [];

  /** Prevent stale folder results when users click folders quickly. */
  private driveFolderLoadSeq = 0;
  /** In-flight silent prefetches keyed by `files:folder:{driveId}:{folderId}`. */
  private readonly folderPrefetchInFlight = new Set<string>();
  private folderPrefetchTimer: ReturnType<typeof setTimeout> | null = null;
  /** Keep low - parallel prefetches were a primary Graph 429 source with All Files. */
  private readonly folderPrefetchConcurrency = 2;
  private readonly folderPrefetchMaxBatch = 8;
  /** When true, idle prefetch is suspended so user clicks / task loads get Graph capacity. */
  private folderPrefetchPaused = false;
  /** In-flight silent HR person task prefetches keyed by `tasks:hr-folder:{name}`. */
  private readonly hrTaskPrefetchInFlight = new Set<string>();
  private hrTaskPrefetchTimer: ReturnType<typeof setTimeout> | null = null;
  /** Debounce timer for hover/pointerdown single-person Comments warm. */
  private hrPersonHoverPrefetchTimer: ReturnType<typeof setTimeout> | null = null;
  /** True while the one-at-a-time "warm every HR person" loop is running. */
  private hrWarmAllRunning = false;
  /**
   * When each HR folder last finished a FULL crawl this session (not a partial seed or a
   * Redis snapshot). Reopening within `hrFilesSkipSoftRefreshMs` skips the Graph refresh.
   */
  private readonly hrFolderCompleteAt = new Map<string, number>();
  /**
   * HR folder keys whose in-memory snapshot came from a crawl that has not finished
   * (first paints, prefetch seeds, cancelled loads). These stay out of shared Redis
   * and are resumed, never presented as the final list.
   */
  private readonly hrFolderPartialKeys = new Set<string>();
  /** Prevent stale All Files task loads when folders are clicked rapidly. */
  private allFilesFolderTaskLoadSeq = 0;

  protected hrPersonalRootFolders: Array<{
    id: string;
    name: string;
    webUrl: string;
    isFolder: boolean;
  }> = [];

  private userRootFolderId: string | null = null;
  protected hrPersonalFolderName: string | null = null;

  // ============================================================
  // CURRENT USER
  // ============================================================
  protected currentUser: {
    email: string;
    username: string;
    userPrincipalName: string;
    employeeId?: string;
  } | null = null;

  protected get isUserLoggedIn(): boolean { return !!this.currentUser; }

  /** Item count shown in the attachments header - matches what filteredAttachments displays. */
  protected get loadedHrPersonalItemCount(): number {
    if (this.currentLibraryDriveId || this.isUsingHrTaskList) {
      return this.filteredAttachments.length;
    }
    return this.visibleFiles.length;
  }

  /** True when the attachments panel is showing the user's HR Personal drive tree (not All Files). */
  private isBrowsingPersonalDriveAttachments(): boolean {
    return !!this.userRootFolderId && !this.currentLibraryDriveId && !this.showAllFilesSection;
  }


  // ============================================================
  // FILTER USERS
  // ============================================================
  protected readonly approvalFilterOptions: AppDropdownOption[] = [
    { value: '', label: 'All statuses', tone: 'neutral' },
    { value: 'approved', label: 'Complete', tone: 'complete' },
    { value: 'pending', label: 'Pending', tone: 'pending' },
    { value: 'general-comments', label: 'General Comments', tone: 'comment' },
    { value: 'new-attachments', label: 'New Attachments', tone: 'attachment' },
    { value: 'rfa', label: 'Request For Action', tone: 'action' },
  ];
  protected readonly hrTaskListOptions: AppDropdownOption[] = [
    { value: '', label: '-- Select a list --' },
    { value: 'ProcEForm', label: 'ProcEForm' },
  ];

  protected getTaskFilterOptions(forms: EForm[] | null | undefined): AppDropdownOption[] {
    return [
      { value: '', label: 'All Tasks' },
      ...(forms ?? []).map(form => ({
        value: form.taskListName || form.title,
        label: this.formatFormTitle(form.title),
      })),
    ];
  }

  protected getAssigneeOptions(assignedTo: string, eform?: { status?: string; completedBy?: string; eFormDetails?: any }): AppDropdownOption[] {
    return this.getAssignees(assignedTo).map(person => ({
      value: person,
      label: this.formatAssigneeLabel(person, eform),
    }));
  }

  /**
   * Multi-assignee dropdown only while the task is still open.
   */
  protected showAssigneeDropdown(
    assignedTo: string,
    eform?: { status?: string; completedBy?: string; eFormDetails?: any },
  ): boolean {
    return this.getAssignees(assignedTo).length >= 2 && !this.isEformComplete(eform);
  }

  /** True when the comment/task card should use Completed By instead of Assigned To. */
  protected isEformComplete(eform?: { status?: string; eFormDetails?: any }): boolean {
    const details = eform?.eFormDetails ?? {};
    const progress = String(details.progress ?? details.eFormProgress ?? '').toLowerCase().trim();
    if (progress === 'complete' || progress === 'completed') return true;
    if (progress === 'pending') return false;

    const taskOutcome = String(details.taskOutcome ?? details.taskOutcomeField ?? '').toLowerCase().trim();
    if (taskOutcome === 'complete' || taskOutcome === 'completed') return true;

    const status = this.getEformStatus(eform);
    if (/\b(pending|awaiting|waiting|needs?)\b.*\bapprov/.test(status)) return false;
    return status.includes('approv') || status.includes('complet');
  }

  /** Colour of the status pill on a Comments card. */
  protected getCommentStatusTone(eform?: { status?: string; eFormDetails?: any }): HrTaskDetailTone {
    const status = String(eform?.eFormDetails?.status ?? '').toLowerCase();
    if (/reject|denied|cancel/.test(status)) return 'danger';
    if (this.isEformComplete(eform)) return 'success';
    if (/pending|progress|waiting|awaiting/.test(status)) return 'warning';
    return 'neutral';
  }

  /** Completed By when approved; falls back to last editor only when no completer is stored. */
  protected getLastModifiedByName(
    eform?: { status?: string; completedBy?: string; eFormDetails?: any } | null,
  ): string {
    const details = eform?.eFormDetails ?? {};
    const rawFields = details.rawFields;

    const completedByName =
      extractPersonName(eform?.completedBy) ||
      extractPersonName(details.completedBy) ||
      (rawFields && typeof rawFields === 'object'
        ? extractPersonName((rawFields as Record<string, unknown>)['CompletedBy'])
        : '');
    if (completedByName) return completedByName;

    const completed = this.getCompletedByNames(eform ?? {}).find(Boolean);
    if (completed) return completed;

    const custom = String(details.customModifiedBy ?? '').trim();
    if (custom) return custom;

    if (rawFields && typeof rawFields === 'object') {
      const fromFields =
        extractPersonName((rawFields as Record<string, unknown>)['CustomModifiedBy']) ||
        extractPersonName((rawFields as Record<string, unknown>)['Editor']) ||
        extractPersonName((rawFields as Record<string, unknown>)['ModifiedBy']);
      if (fromFields) return fromFields;
    }

    const modifiedBy = String(details.modifiedBy ?? '').trim();
    // Ignore bare SharePoint lookup ids.
    if (modifiedBy && !/^\d+$/.test(modifiedBy)) return modifiedBy;

    return '';
  }

  /** Appends ✓ when this assignee has completed; pending names stay plain. */
  protected formatAssigneeLabel(person: string, eform?: { status?: string; completedBy?: string; eFormDetails?: any }): string {
    const name = String(person ?? '').trim();
    if (!name) return '';
    if (this.hasAssigneeCompleted(name, eform)) {
      return `${name} ✓`;
    }
    return name;
  }

  protected hasAssigneeCompleted(person: string, eform?: { status?: string; completedBy?: string; eFormDetails?: any }): boolean {
    const name = String(person ?? '').trim();
    if (!name || !eform) return false;

    const matchedCompleter = this.getCompletedByNames(eform).some(completer =>
      this.userService.matchesAssigneeField(name, completer, completer)
    );
    if (matchedCompleter) return true;

    if (this.isEformComplete(eform)) {
      const lastModifier = this.getLastModifiedByName(eform);
      if (lastModifier && this.userService.matchesAssigneeField(name, lastModifier, lastModifier)) {
        return true;
      }
      return true;
    }

    return false;
  }

  private getEformStatus(eform?: { status?: string; eFormDetails?: any }): string {
    return String(eform?.eFormDetails?.status ?? eform?.status ?? '').toLowerCase();
  }

  private isEformPending(eform?: { status?: string; eFormDetails?: any }): boolean {
    const status = this.getEformStatus(eform);
    if (!status) return false;
    return !this.isEformComplete(eform);
  }

  private getCompletedByNames(eform: { completedBy?: string; eFormDetails?: any }): string[] {
    const details = eform.eFormDetails ?? {};
    const names = new Set<string>();

    const push = (value: unknown) => {
      const text = String(value ?? '').trim();
      if (!text || text.toLowerCase() === 'enter value here') return;
      for (const part of this.userService.parseAssignees(text)) {
        names.add(part);
      }
    };

    push(eform.completedBy);
    push(details.completedBy);

    const rawFields = details.rawFields;
    if (rawFields && typeof rawFields === 'object') {
      push((rawFields as Record<string, unknown>)['CompletedBy']);
    }

    const comment = String(details.comment ?? details.commentHtml ?? '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, '');
    const completedByMatches = comment.matchAll(/\[Completed by\s+([^\]]+?)(?:\s+on\s+[^\]]+)?\]/gi);
    for (const match of completedByMatches) {
      const raw = String(match[1] ?? '').replace(/:\s*$/, '').trim();
      push(raw);
    }

    // Approvers who already acted (have a date), or the primary completer when the task is done.
    if (String(details.approver1Date ?? '').trim()) push(details.approver1);
    if (String(details.approver2Date ?? '').trim()) push(details.approver2);
    if (String(details.approver3Date ?? '').trim()) push(details.approver3);
    if (this.isEformComplete(eform)) {
      push(details.approver1 || details.approver2 || details.approver3);
    }

    return Array.from(names);
  }

  // ============================================================
  // COMMENTS / HR TASKS DATA
  // ============================================================
  protected commentItems: Array<{
    id: string;
    name: string;
    webUrl: string;
    lastModifiedDateTime?: string;
    size?: number;
    isFolder: boolean;
    status?: string;
    submittedBy?: string;
    submittedDate?: string;
    completedBy?: string;
    description?: string;
    contentType?: string;
    downloadUrl?: string;
    extractedContent?: string;
    eFormDetails?: any;
    isContentLoaded?: boolean;
    showFullContent?: boolean;
  }> = [];

  // ============================================================
  // FILTER STATE - cached, only recomputed when dirty
  // ============================================================
  protected selectedTaskFilter = '';
  protected selectedApprovalFilter = '';
  protected searchQuery = '';
  protected selectedSubmitter: string | null = null;
  protected selectedEFormListId: string | null = null;
  protected selectedEFormTitle: string | null = null;
  protected selectedGroupTaskIds: string[] = [];
  protected superiorMode = false;

  // eForms loaded from SharePoint, used to populate the task-type filter dropdown.
  protected taskFilterForms$!: Observable<EForm[]>;

  private _filteredCommentItems: Array<typeof this.commentItems[0]> = [];
  private _filtersDirty = true;

  // ============================================================
  // FILTERED COMMENT ITEMS - latest-first pages + Load more
  // ============================================================
  protected get filteredCommentItems(): Array<typeof this.commentItems[0]> {
    if (!this._filtersDirty) return this._filteredCommentItems;
    this._filteredCommentItems = this.filterCommentItems();
    this._filtersDirty = false;
    return this._filteredCommentItems;
  }

  /** Placeholder card count for the Comments skeleton while tasks load. */
  protected readonly commentsSkeletonRows = [1, 2, 3, 4, 5, 6];

  /** Short, user-facing Comments loader text (no raw folder/email dump). */
  protected get commentsLoadingLabel(): string {
    if (this.isLoadingMoreComments) {
      return 'Loading more tasks...';
    }
    const custom = this.commentsMessage?.trim();
    if (this.isLoadingComments) {
      // Prefer progress text (e.g. "Scanning ProcTaskArchive...") over a generic label.
      if (custom) return custom;
      if (this.selectedHrPersonalTaskFolder || this.selectedFolderName) {
        return 'Loading tasks...';
      }
      return 'Loading comments...';
    }
    return custom || 'Loading comments...';
  }

  /** First paint shows the newest page; "Load more" reveals the next page from the loaded/cached set. */
  protected visibleItemCount = 20;
  protected readonly commentsInitialPageSize = 20;
  protected readonly itemPageSize = 20;

  protected get pagedCommentItems(): Array<typeof this.commentItems[0]> {
    return this.filteredCommentItems.slice(0, this.visibleItemCount);
  }

  protected get hasMoreItems(): boolean {
    return this.filteredCommentItems.length > this.visibleItemCount;
  }

  protected get remainingItemCount(): number {
    return Math.max(this.filteredCommentItems.length - this.visibleItemCount, 0);
  }

  protected loadMoreItems(): void {
    this.visibleItemCount += this.itemPageSize;
    this.refreshView();
  }

  private invalidateCommentFilters(resetPagination = true): void {
    this._filtersDirty = true;
    if (resetPagination) {
      this.visibleItemCount = this.commentsInitialPageSize;
    }
  }

  /** Newest submittedDate first - ignore submitter for ordering. */
  private sortCommentItemsByDateDesc<T extends { submittedDate?: string; lastModifiedDateTime?: string }>(
    items: T[],
  ): T[] {
    return [...items].sort(
      (a, b) =>
        new Date(b.submittedDate || b.lastModifiedDateTime || 0).getTime() -
        new Date(a.submittedDate || a.lastModifiedDateTime || 0).getTime(),
    );
  }

  /**
   * Publish folder/HR task results: always date-sorted (latest first).
   * Prefer painting when a load is complete so newer rows do not keep jumping to the top.
   * "Load more" pages the already-loaded set.
   * `background: true` = silent merge after the interactive spinner is already gone.
   */
  private publishFolderTaskProgress(
    items: any[],
    options: {
      done?: boolean;
      emptyMessage?: string;
      background?: boolean;
      /** While true, show "Loading more tasks…" during silent background Graph top-up. */
      keepLoadingMore?: boolean;
    } = {},
  ): void {
    const sorted = this.sortCommentItemsByDateDesc(
      this.filterTasksForSelectedAllFilesFolder(items)
    );
    this.commentItems = sorted;
    void this.publishSelectedFolderChildSubjects(sorted, !!options.done && !options.background);

    // Silent top-up after first paint — keep Load more position.
    // `keepLoadingMore` stays true while deeper Graph pages are still coming in.
    if (options.background) {
      this.isLoadingComments = false;
      this.isLoadingMoreComments = !!options.keepLoadingMore;
      if (sorted.length > 0) {
        this.commentsMessage = '';
      }
      this.invalidateCommentFilters(false);
      this.refreshView();
      this.focusPendingCommentIfNeeded();
      return;
    }

    if (options.done) {
      const wasLoading = this.isLoadingComments;
      this.isLoadingComments = false;
      this.isLoadingMoreComments = false;
      this.commentsMessage = sorted.length === 0 ? (options.emptyMessage || 'No tasks found.') : '';
      // Reset page window only on the first completed paint - keep Load more position
      // when a background related-steps refresh replaces the list.
      if (wasLoading || this.visibleItemCount < this.commentsInitialPageSize) {
        this.visibleItemCount = this.commentsInitialPageSize;
        this.invalidateCommentFilters(true);
      } else {
        this.invalidateCommentFilters(false);
      }
      this.refreshView();
      this.focusPendingCommentIfNeeded();
      return;
    }

    // Incomplete batches: paint first hits, then keep "Loading more tasks..." until
    // done:true — otherwise assignee/supplement top-up looks stuck mid-count.
    if (this.isLoadingComments) {
      if (sorted.length === 0) {
        this.commentsMessage = 'Loading tasks...';
        this.refreshView();
        return;
      }
      this.isLoadingComments = false;
      this.visibleItemCount = this.commentsInitialPageSize;
      this.invalidateCommentFilters(true);
      this.commentsMessage = '';
    } else {
      this.invalidateCommentFilters(false);
      if (sorted.length > 0) {
        this.commentsMessage = '';
      }
    }
    this.isLoadingMoreComments = true;
    this.refreshView();
    this.focusPendingCommentIfNeeded();
  }

  private focusPendingCommentIfNeeded(): void {
    const id = String(this.pendingCommentFocusId ?? '').trim();
    if (!id) return;

    const idx = this.filteredCommentItems.findIndex((item) => String(item?.id ?? '') === id);
    if (idx < 0) return;

    this.visibleItemCount = Math.max(this.visibleItemCount, idx + 1);
    this.highlightedCommentId = id;
    this.pendingCommentFocusId = null;
    this.refreshView();

    setTimeout(() => {
      const el = document.getElementById(`comment-item-${id}`);
      el?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 80);
  }

  private childSubjectCacheKey(libraryName: string, folderEFormId: string): string {
    return `${this.normalizeCacheKeyPart(libraryName)}:${folderEFormId}`;
  }

  /**
   * Unblock the ChildSubject accordion without waiting for the full folder task scan.
   * Procurement ChildSubject columns live on the source eForm item (trailing folder id).
   */
  private async primeChildSubjectForFolder(
    libraryName: string,
    folderName: string,
    folderId: string,
  ): Promise<void> {
    if (!folderId || !this.showAllFilesSection) return;

    const eFormId = extractTrailingFolderId(folderName);
    if (!eFormId) return;

    const cacheKey = this.childSubjectCacheKey(libraryName, eFormId);
    const cached = this.childSubjectNamesByLibraryEFormId.get(cacheKey);
    if (cached && cached.length > 0) {
      if (this.selectedAllFilesFolderId === folderId) {
        this.childSubjectSourceFetchedForFolderId = folderId;
        this.allFilesComponent?.setChildSubjectFolderNames(folderId, cached);
      }
      return;
    }

    const names = await this.loadChildSubjectNamesFromSourceEFormDirect(libraryName, folderName, eFormId);
    if (this.selectedAllFilesFolderId !== folderId) return;

    if (names.length > 0) {
      this.childSubjectSourceFetchedForFolderId = folderId;
      this.childSubjectNamesByLibraryEFormId.set(cacheKey, names);
      this.allFilesComponent?.setChildSubjectFolderNames(folderId, names);
    }
    // Empty / failed: leave pending for publishSelectedFolderChildSubjects (task path).
  }

  /** Populate All Files level 2 from ChildSubject + ChildSubject2 on the source eForm row. */
  private async publishSelectedFolderChildSubjects(items: any[], done: boolean): Promise<void> {
    const folderId = this.selectedAllFilesFolderId;
    if (!folderId || !this.showAllFilesSection) return;

    let names = this.extractChildSubjectNamesFromTaskItems(items);
    if (names.length === 0 && this.childSubjectSourceFetchedForFolderId !== folderId && items.length > 0) {
      names = await this.loadChildSubjectNamesFromSourceEForm(items);
      this.childSubjectSourceFetchedForFolderId = folderId;
      const libraryName = this.currentLibraryName ?? '';
      const eFormId = extractTrailingFolderId(this.selectedFolderName);
      if (names.length > 0 && libraryName && eFormId) {
        this.childSubjectNamesByLibraryEFormId.set(this.childSubjectCacheKey(libraryName, eFormId), names);
      }
    }

    if (names.length > 0) {
      this.allFilesComponent?.setChildSubjectFolderNames(folderId, names);
      return;
    }

    // Don't wipe an early primeChildSubjectForFolder result when tasks have no ChildSubject columns.
    if (done && this.childSubjectSourceFetchedForFolderId !== folderId) {
      this.allFilesComponent?.setChildSubjectFolderNames(folderId, names);
    }
  }

  private extractChildSubjectNamesFromTaskItems(items: any[]): string[] {
    const names: string[] = [];
    const seen = new Set<string>();
    const selectedName = normalizeName(this.selectedFolderName);

    for (const item of items) {
      const fields = item?.eFormDetails?.rawFields ?? {};
      names.push(...this.extractChildSubjectNamesFromFields(fields, selectedName, seen));
    }
    return names;
  }

  private extractChildSubjectNamesFromFields(
    fields: Record<string, unknown>,
    selectedName: string,
    seen: Set<string>,
  ): string[] {
    const names: string[] = [];
    for (const [key, rawValue] of Object.entries(fields)) {
      if (!isChildSubjectColumn(key)) continue;

      const text = stringifyTaskFieldValue(rawValue);
      for (const candidate of text.split(/[\r\n;]+/)) {
        const name = candidate.replace(/^#?\d+;#/, '').trim();
        const normalizedName = normalizeName(name);
        if (!name || !normalizedName || normalizedName === selectedName || seen.has(normalizedName)) {
          continue;
        }
        seen.add(normalizedName);
        names.push(name);
      }
    }
    return names;
  }

  /** ChildSubject columns live on the source eForm list item, not always on the task row. */
  private async loadChildSubjectNamesFromSourceEForm(items: any[]): Promise<string[]> {
    const folderEFormId = extractTrailingFolderId(this.selectedFolderName);
    const seedTask =
      items.find((item) => String(item?.eFormDetails?.eFormListId ?? '').trim() === folderEFormId) ??
      items[0];
    if (!seedTask) return [];

    const eFormListId = String(seedTask?.eFormDetails?.eFormListId ?? folderEFormId ?? '').trim();
    if (!eFormListId) return [];

    const libraryName = this.currentLibraryName ?? '';
    const taskListName = String(seedTask?.eFormDetails?.listName ?? '').trim();
    return this.loadChildSubjectNamesFromSourceEFormDirect(libraryName, this.selectedFolderName, eFormListId, taskListName);
  }

  /** Resolve ChildSubject names from the source eForm item by trailing folder / eForm id. */
  private async loadChildSubjectNamesFromSourceEFormDirect(
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
      const token = await this.getSharePointToken();
      await this.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );
      if (!this.cachedSiteId) return [];

      const sourceListName = this.parseSharePointListNameFromUrl(form.url) || String(form.prefix ?? '').trim();
      const sourceList = this.cachedSiteLists.find((list: any) =>
        (list.name ?? '').toLowerCase() === sourceListName.toLowerCase() ||
        (list.displayName ?? '').toLowerCase() === sourceListName.toLowerCase()
      );
      if (!sourceList?.id) return [];

      const response: any = await graphGetWithRetry(
        this.http,
        `/sites/${this.cachedSiteId}/lists/${sourceList.id}/items/${eFormListId}?$expand=fields`,
        token,
        AppConstants.graphFileListingTimeoutMs,
      );
      const fields = (response?.fields ?? {}) as Record<string, unknown>;
      return this.extractChildSubjectNamesFromFields(fields, normalizeName(folderName), new Set());
    } catch (err) {
      console.warn('[All Files] Could not load ChildSubject columns from source eForm item.', err);
      return [];
    }
  }

  private parseSharePointListNameFromUrl(url: string): string {
    return this.formConfigService.parseSharePointListNameFromUrl(url);
  }

  /** Level-2 All Files dropdown uses ChildSubject and ChildSubject2 only. */


  // ============================================================
  // DATE SEARCH HELPER
  // ============================================================

  private filterCommentItems(): Array<typeof this.commentItems[0]> {
    let filtered = this.commentItems;
    if (this.selectedSubmitter && !this.showAllFilesSection) {
      filtered = filtered.filter(item => {
        const matchesEFormList = !this.selectedEFormListId || (
          (String(item.eFormDetails?.eFormListId ?? '').trim() || 'Unknown eForm') === this.selectedEFormListId
        );
        if (!matchesEFormList) {
          return false;
        }

        const isCommentEntry = isCommentTypeItem(item);
        const selectedTaskId = String(this.selectedTask?.Id ?? '').trim();
        const groupTaskIds = new Set(this.selectedGroupTaskIds);

        if (!isCommentEntry) {
          const matchesEFormTitle = !this.selectedEFormTitle || (
            getTaskEFormTitle(item) === this.selectedEFormTitle
          );
          if (!matchesEFormTitle) {
            return false;
          }
        } else if (isSyntheticCommentCard(item)) {
          const cardTaskId = extractCommentCardTaskId(item.id);
          if (!cardTaskId) {
            return true;
          }
          if (groupTaskIds.size > 0) {
            return groupTaskIds.has(cardTaskId);
          }
          if (selectedTaskId && cardTaskId !== selectedTaskId) {
            return false;
          }
        }

        if (this.showToDoSection && !this.superiorMode) {
          if (isCommentEntry) {
            return true;
          }
          if (!isItemSubmittedBy(item, this.selectedSubmitter)) {
            return false;
          }
          // When a submitter is selected in the To Do view, the Comments panel should
          // show the submitter's full eForm history (Pending + Completed), not only
          // items currently assigned to the logged-in user.
          return true;
        }

        if (this.showToDoSection && this.superiorMode && this.selectedSubmitter) {
          if (isCommentEntry) {
            return true;
          }
          return this.isItemAssignedToSubordinate(item, this.selectedSubmitter);
        }

        return true;
      });
    }

    if (this.searchQuery.trim()) {
      const query = this.searchQuery.toLowerCase().trim();
      filtered = filtered.filter(item => {
        const details = item.eFormDetails ?? {};
        const searchableValues = [
          item.name,
          item.description,
          item.submittedBy,
          item.status,
          getDateSearchTokens(item.submittedDate ?? ''),
          details.type,
          details.status,
          details.submitter,
          details.listName,
          details.approver1,
          details.approver2,
          details.approver3,
          details.reason,
          details.stage,
          details.section,
          details.body,
          details.comment,
          details.category,
          details.taskOutcome,
          details.assignedTo,
          details.fromDate,
          details.toDate,
          details.eFormListId,
          details.employeeName,
          details.comment,
          details.commentHtml,
          details.eFormListId,
          details.eFormCategory,
          details.eFormProgress,
          details.eFormListId,
          details.eFormListId,
        ];

        return searchableValues.some(value => String(value ?? '').toLowerCase().includes(query));
      });
    }

    if (!this.showAllFilesSection && this.selectedTaskFilter) {
      const normalize = (v: string) => v.toLowerCase().replace(/\s+/g, '');
      const tf = normalize(this.selectedTaskFilter);
      filtered = filtered.filter(item => {
        const type = normalize(String(item.eFormDetails?.type ?? ''));
        const list = normalize(String(item.eFormDetails?.listName ?? ''));
        return type.includes(tf) || list.includes(tf);
      });
    }

    if (this.selectedApprovalFilter) {
      filtered = filtered.filter(item => matchesApprovalFilter(item, this.selectedApprovalFilter));
    }

    if (this.showAllFilesSection) {
      const folderName = String(this.selectedFolderName ?? '').trim();
      if (folderName && extractTrailingFolderId(folderName)) {
        filtered = filtered.filter(item => this.folderMatch.doesMappedHrTaskMatchFolder(item, folderName));
      }
    }

    if (this.isHrPersonalFilesContext()) {
      filtered = filtered.filter(item => !isCommentTypeItem(item));
    }

    return filtered;
  }

  /** Same title resolution used by todo-list Level-2 grouping. */

  private getPrimaryItemAssignee(item: { eFormDetails?: Record<string, unknown>; assignedTo?: unknown }): string {
    const details = item.eFormDetails ?? {};
    const raw = String(details['assignedTo'] ?? item.assignedTo ?? '').trim();
    return this.userService.parseAssignees(raw)[0] ?? '';
  }

  private getSubordinateItemAssigneeValues(item: { eFormDetails?: Record<string, unknown>; assignedTo?: unknown }): string[] {
    const primary = this.getPrimaryItemAssignee(item);
    return primary ? [primary] : [];
  }


  private isItemAssignedToSubordinate(item: { eFormDetails?: Record<string, unknown>; assignedTo?: unknown }, subordinateName: string): boolean {
    const assigneeValues = this.getSubordinateItemAssigneeValues(item);
    if (!assigneeValues.length) {
      return false;
    }

    const subordinate = this.subordinaryTaskService.getCachedSubordinates().find(employee =>
      employee.FullName === subordinateName ||
      employee.ADmail === subordinateName
    );

    if (!subordinate) {
      return false;
    }

    return assigneeValues.some(value =>
      this.subordinaryTaskService.isTaskAssignedToEmployeeForDomainUser(value, subordinate)
    );
  }

  // Alias used by the template
  protected get attachmentItems() { return this.userFiles; }

  // ============================================================
  // FOLDER MAP - cached, rebuilt only when userFiles changes
  // ============================================================
  private _folderMap = new Map<string, typeof this.userFiles>();
  private _folderMapDirty = true;

  private refreshAttachmentFolderMap(): void {
    if (!this._folderMapDirty) return;
    this._folderMap.clear();
    for (const f of this.userFiles) {
      const key = f.parentId ?? '';
      if (!this._folderMap.has(key)) this._folderMap.set(key, []);
      this._folderMap.get(key)!.push(f);
    }
    this._folderMapDirty = false;
  }

  protected get visibleFiles() {
    this.refreshAttachmentFolderMap();
    const key = this.currentFolderId ?? this.userRootFolderId ?? '';
    return this._folderMap.get(key) ?? [];
  }

  /** Show Back while browsing an opened folder tree. */
  protected get showBackButtonInAttachments(): boolean {
    return this.showUserFile && !!this.attachmentBrowsingRoot;
  }

  /** True when Back can move up to a parent folder within the opened tree. */
  protected get canGoBackInAttachments(): boolean {
    return this.folderStack.length > 0;
  }

  private get hasLoadedPersonalFiles(): boolean {
    return !this.isUsingHrTaskList && this.userRootFolderId !== null;
  }

  // ============================================================
  // GRAPH API SCOPES & LIBRARY CONFIG
  // ============================================================
  private readonly targetLibraryName = sharePointConfig.hrPersonalListDisplayName || 'HRPersonal';

  // ============================================================
  // CACHING
  // ============================================================
  private cachedDriveId: string | null = null;
  private hrPersonalDriveId: string | null = null;
  private cachedSiteId: string | null = null;
  private cachedSiteWebUrl: string = '';
  private cachedSiteLists: any[] = [];
  /**
   * Where each To Do list's paging stopped, keyed by list name. Normally empty after a
   * full load (every nextLink drained). A non-null cursor means a rare partial remaining
   * for the "Load remaining tasks" retry.
   */
  private readonly todoListCursors = new Map<string, string | null>();

  /**
   * Newest `lastModifiedDateTime` already ingested per list. The manual refresh pages
   * newest-first and stops at the first row at or below this mark, so it downloads only
   * what changed since the last look instead of re-collecting and re-filtering the lot.
   */
  private readonly todoListWatermarks = new Map<string, string>();
  protected isLoadingMoreTodoTasks = false;
  /** Interval that picks up external SharePoint task edits into Redis while signed in. */
  private sharePointCachePollTimer: ReturnType<typeof setInterval> | null = null;
  private lastFolderCommentsPollMs = 0;
  private readonly onDocumentVisibilityForCachePoll = (): void => {
    if (document.visibilityState === 'visible') {
      this.startSharePointCachePoll();
      void this.pollSharePointTaskCaches();
    } else {
      this.stopSharePointCachePollTimerOnly();
    }
  };

  /** True while any To Do list still has unfetched older rows. */
  protected get hasMoreTodoTasks(): boolean {
    for (const cursor of this.todoListCursors.values()) {
      if (cursor) return true;
    }
    return false;
  }

  /** True while a procurement list still has Progress=Pending pages left to fetch. */
  private hasMoreProcurementTodoCursors(): boolean {
    for (const [listName, cursor] of this.todoListCursors.entries()) {
      if (cursor && isTodoProcurementTaskList(listName)) return true;
    }
    return false;
  }
  /** True after a full To Do fetch (HR + procurement lists) has populated the task cache. */
  private todoScopeCacheValid = false;
  private readonly folderCache = new Map<string, { id: string; name: string; webUrl?: string }>();


  // ============================================================
  // PRIVATE RUNTIME STATE
  // ============================================================
  private userFileLoadingTimeoutId: ReturnType<typeof setTimeout> | null = null;
  private hrOlderHistoryLoadInProgress = false;
  private selectedHrPersonalTaskFolder = '';
  /** Survives tab switches so HR Files Comments can restore from cache instantly. */
  private lastHrFilesCommentsFolder = '';

  // Batches multiple refreshView() calls into one detectChanges()
  private viewUpdatePending = false;

  // ============================================================
  // ATTACHMENT SEARCH
  // ============================================================
  attachmentSearch = '';

  get filteredAttachments() {
    // For HR Task lists (All Files), show all userFiles directly
    if (this.isUsingHrTaskList) return this.userFiles;

    // Library browsing loads one folder level at a time via Graph - userFiles is already scoped
    const source = this.currentLibraryDriveId ? this.userFiles : this.visibleFiles;
    return this.applyAttachmentSearch(source);
  }

  private applyAttachmentSearch(source: typeof this.userFiles) {
    if (!this.attachmentSearch.trim()) return source;
    const query = this.attachmentSearch.toUpperCase().trim();
    return source.filter(item => {
      if (item.isFolder) return false;
      if (item.name?.toUpperCase().includes(query.toUpperCase())) return true;
      if (item.modifiedBy?.toUpperCase().includes(query.toUpperCase())) return true;
      return false;
    });
  }

  // ============================================================
  // TRACKING
  // ============================================================
  protected getAssignees(assignedTo: string | undefined | null): string[] {
    return this.userService.parseAssignees(assignedTo);
  }

  /** Assigned To text for comment cards - rebuilds from raw SharePoint fields when blank. */
  protected getAssignedToDisplay(eform: { eFormDetails?: any; assignedTo?: string } | null | undefined): string {
    const details = eform?.eFormDetails ?? {};
    const direct = String(details.assignedTo ?? eform?.assignedTo ?? '').trim();
    if (direct) return direct;

    const rawFields = details.rawFields;
    if (rawFields && typeof rawFields === 'object') {
      return collectTaskAssigneeValues(rawFields as Record<string, any>).join(', ');
    }
    return '';
  }

  // ============================================================
  // LIFECYCLE
  // ============================================================
  async ngOnInit(): Promise<void> {
    this.taskFilterForms$ = this.formConfigService.hrForms$;
    this.documentLibraryTaskMapSub = this.formConfigService.documentLibraryTaskMap$
      .subscribe((map) => { this.documentLibraryTaskMap = map ?? {}; });
    await this.checkLoginOnStart();
    try {
      const tasks = await this.taskService.loadTaskLists();
      this.allowedHrTaskLists = tasks;
    } catch (error) {
      console.warn('Could not load SharePoint task lists.', error);
      this.allowedHrTaskLists = [];
    }
  }

  ngOnDestroy(): void {
    this.documentLibraryTaskMapSub?.unsubscribe();
    this.stopSharePointCachePoll();
    if (this.folderPrefetchTimer) {
      clearTimeout(this.folderPrefetchTimer);
      this.folderPrefetchTimer = null;
    }
    if (this.hrTaskPrefetchTimer) {
      clearTimeout(this.hrTaskPrefetchTimer);
      this.hrTaskPrefetchTimer = null;
    }
  }

  // ============================================================
  // BATCHED VIEW UPDATE
  // ============================================================
  protected refreshView(): void {
    if (this.viewUpdatePending) return;
    this.viewUpdatePending = true;
    queueMicrotask(() => {
      this.viewUpdatePending = false;
      this.cdr.detectChanges();
    });
  }

  // ============================================================
  // SSO INITIALIZATION
  // ============================================================
  protected async checkLoginOnStart(): Promise<void> {
    try {
      await this.authService.initialize();
      let response = null;
      try {
        response = await this.authService.handleRedirect();
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
          await this.beginSignInRedirect('login');
          return;
        }
        throw error;
      }
      if (response) {
        this.authService.setActiveAccount(response.account);
        await this.loadCurrentUserProfile(response.accessToken);
        this.statusMessage = `Signed in as ${this.currentUser?.username || 'User'}`;
        this.authService.restorePostLoginPathIfNeeded();
        return;
      }
      await this.trySilentSignIn();
    } catch {
      await this.beginSignInRedirect();
    }
  }

  private async trySilentSignIn(): Promise<void> {
    let account = this.authService.getActiveAccount();
    const cachedAccounts = this.authService.getAllAccounts();
    if (!account && cachedAccounts.length > 0) {
      account = cachedAccounts[0];
      this.authService.setActiveAccount(account);
    }

    if (!account) {
      await this.beginSignInRedirect();
      return;
    }

    try {
      this.authService.setActiveAccount(account);
      const token = await this.authService.acquireGraphToken();
      await this.loadCurrentUserProfile(token);
      this.formConfigService.load().subscribe();
      this.statusMessage = `Signed in as ${this.currentUser?.username || 'User'}`;
    } catch (error) {
      const prompt = error instanceof InteractionRequiredAuthError ? 'login' : 'none';
      await this.beginSignInRedirect(prompt);
    }
  }

  private async beginSignInRedirect(prompt: 'none' | 'login' = 'none'): Promise<void> {
    if (this.isConnecting) return;
    this.errorMessage = '';
    this.isConnecting = true;
    this.statusMessage = 'Redirecting to Microsoft sign-in...';
    try {
      await this.authService.loginRedirect(prompt);
    } catch {
      this.isConnecting = false;
      this.statusMessage = 'Could not start Microsoft sign-in. Please refresh the page.';
    }
  }

  private async loadCurrentUserProfile(token: string): Promise<void> {
    await new Promise<void>(resolve => {
      this.graphGet('/me?$select=userPrincipalName,mail,displayName,employeeId', token).subscribe({
        next: async (user: any) => {
          const previousUserPrincipalName = this.currentUser?.userPrincipalName ?? '';
          const nextUserPrincipalName = user.userPrincipalName || user.mail || '';
          if (previousUserPrincipalName && previousUserPrincipalName !== nextUserPrincipalName) {
            this.clearUserScopedSharePointState();
          }

          this.currentUser = {
            email: user.mail || user.userPrincipalName || 'N/A',
            username: user.displayName || user.userPrincipalName || 'User',
            userPrincipalName: nextUserPrincipalName,
            employeeId: user.employeeId || undefined,
          };

          // Set current user in UserService for other components to use
          this.userService.setCurrentUser(this.currentUser);

          // Hydrate personal + shared Redis crawl cache so All Files / HR Files
          // paint instantly after a tab restart (including other users' folder crawls).
          await this.fileCrawlCache.bindToUser(nextUserPrincipalName);

          // Warm shared site/list metadata once before To Do / folder lookups compete.
          try {
            const spToken = await this.getSharePointToken();
            await this.ensureSiteMetadata(
              sharePointConfig.siteHostName,
              sharePointConfig.sitePath,
              spToken,
            );
          } catch {
            // To Do / HR Files will retry via their own ensureSiteMetadata calls.
          }

          // To Do first - primary Graph consumer after login.
          this.initializeToDoOnLogin();

          // Defer HR warm + delegates so they don't race the first To Do Graph wave.
          // People-list warm fills shared Redis early so the first HR Files open is fast.
          setTimeout(() => {
            void this.loadHrPersonalFolderName();
            void this.delegateService.loadDelegates();
            void this.warmHrPersonalRootFoldersCache();
          }, 750);
          resolve();
        },
        error: () => {
          this.currentUser = null;
          this.hrPersonalFolderName = null;
          this.userService.setCurrentUser(null);
          void this.delegateService.loadDelegates();
          resolve();
        },
      });
    });
  }

  private clearUserScopedSharePointState(): void {
    this.showUserFile = false;
    this.userFiles = [];
    this.hrPersonalRootFolders = [];
    this.userRootFolderId = null;
    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.folderStack = [];
    this.attachmentBrowsingRoot = null;
    this.currentLibraryDriveId = null;
    this.currentLibraryName = null;
    this.hrPersonalDriveId = null;
    this.hrPersonalFolderName = null;
    this.selectedFolderName = '';
    this.selectedAllFilesParentFolderName = '';
    this.selectedAllFilesFolderId = null;
    this.childSubjectSourceFetchedForFolderId = null;
    this.childSubjectNamesByLibraryEFormId.clear();
    this.selectedHrPersonalTaskFolder = '';
    this.commentItems = [];
    this.commentsMessage = '';
    this.userHrTasksLoaded = false;
    this.todoFirstPassComplete = false;
    this.todoProcurementDrainActive = false;
    this.todoScopeCacheValid = false;
    this.fileCrawlCache.invalidate(AppComponent.HR_USER_TASKS_CACHE_KEY);
    this.userFileError = '';
    this.userFileWarning = '';
    this.userFileProgressMessage = '';
    this._folderMapDirty = true;
    this.folderCache.clear();
    this.clearLocalStorageFolderCache();
    this.invalidateCommentFilters();
  }

  protected logout(): void {
    this.hrPersonalFolderName = null;
    this.stopSharePointCachePoll();
    // Wipe memory + IndexedDB so a shared machine keeps nothing after sign-out.
    this.fileCrawlCache.clearAll();
    this.clearLocalStorageFolderCache();
    this.cachedSiteId = null;
    this.cachedSiteWebUrl = '';
    this.cachedSiteLists = [];
    this.cachedDriveId = null;
    this.driveResolveInFlight = null;
    this.taskQuery.reset();
    this.todoFirstPassComplete = false;
    this.todoProcurementDrainActive = false;
    this.todoScopeCacheValid = false;
    this.todoListCursors.clear();
    this.todoListWatermarks.clear();
    this.siteMetadataService.reset();
    this.authService.logoutRedirect();
    this.hrPersonalFolderName = null;
  }

  private async getSharePointToken(): Promise<string> {
    return this.authService.acquireSharePointToken();
  }

  // ============================================================
  // FILTER EVENT HANDLERS
  // ============================================================
  /** Turns a PascalCase/camelCase eForm title (e.g. "SickLeaveByAppointment") into spaced words ("Sick Leave By Appointment"). */
  protected formatFormTitle(title: string | null | undefined): string {
    return formatFormTitle(title);
  }

  protected onTaskFilterChange(value: string): void {
    this.selectedTaskFilter = value;
    this.invalidateCommentFilters();
    this.refreshView();
  }

  protected onApprovalFilterChange(value: string): void {
    this.selectedApprovalFilter = value;
    this.invalidateCommentFilters();
    this.refreshView();
  }

  protected onSearchChange(event: Event): void {
    this.searchQuery = (event.target as HTMLInputElement).value;
    this.invalidateCommentFilters();
    this.refreshView();
  }

  protected onAttachmentSearchChange(event: Event): void {
    this.attachmentSearch = (event.target as HTMLInputElement).value;
    this.refreshView();
  }

  protected toggleUserInfo(): void {
    if (!this.isUserLoggedIn) { return; }
    this.cancelBackgroundTodoLoad();
    this.leaveAllFilesTaskView();
    this.showToDoSection = false;
    this.showAllFilesSection = false;
    this.hideComments = false;
    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    Object.assign(this, getTaskTabMobileState());
    this.invalidateCommentFilters();
    this.syncMobileSearchInput();
    if (this.hrPersonalRootFolders.length === 0 && !this.isLoadingHrFilesList) {
      this.loadHrFilesInFirstSection();
    } else if (this.hrPersonalRootFolders.length > 0) {
      // Warm Comments cache for visible people while the user browses the list.
      this.scheduleHrFolderTasksPrefetch(this.hrPersonalRootFolders);
    }
    // Instantly restore the last HR person Comments from cache after To Do / All Files.
    this.restoreHrFilesCommentsFromCache();
    this.refreshView();
  }

  /** Rehydrate Comments for the last opened HR Files person without a Graph crawl. */
  private restoreHrFilesCommentsFromCache(): void {
    const folder = String(this.lastHrFilesCommentsFolder || '').trim();
    if (!folder) return;
    const cached =
      this.fileCrawlCache.get(this.hrFolderTaskCacheKey(folder)) ??
      this.fileCrawlCache.getStale(this.hrFolderTaskCacheKey(folder));
    if (!Array.isArray(cached)) return;

    // Leaving the view cancels any crawl in flight. Unless a crawl for this person finished
    // recently, start it again instead of presenting the cached list as final.
    const cacheKey = this.hrFolderTaskCacheKey(folder);
    const completeAt = this.hrFolderCompleteAt.get(cacheKey);
    const finishedRecently =
      completeAt !== undefined && Date.now() - completeAt <= AppConstants.hrFilesSkipSoftRefreshMs;
    if (this.hrFolderPartialKeys.has(cacheKey) || !finishedRecently) {
      this.hideComments = false;
      this.loadAllFilesTasksForLibrary(this.targetLibraryName, folder);
      return;
    }

    this.selectedFolderName = folder;
    this.selectedHrPersonalTaskFolder = folder;
    this.hideComments = false;
    this.publishFolderTaskProgress(this.applyAllFilesFolderSubmitter(cached), {
      done: true,
      emptyMessage: 'No tasks found.',
    });
  }

  /** Open To Do as the default section after sign-in instead of preloading HR Files. */
  private initializeToDoOnLogin(): void {
    this.showToDoSection = true;
    this.showAllFilesSection = false;
    this.hideComments = true;
    Object.assign(this, getTodoTabMobileState());
    void this.loadTodoTasksForCurrentUser();
    // Prefetch HRPersonal group membership (SharePoint REST - does not use Graph slots).
    void this.isUserInHrPersonalAllAccessGroup();
    this.startSharePointCachePoll();
    this.refreshView();
  }

  protected toggleToDoSection(): void {
    // Drop HR Files / All Files folder scope so My Tasks comments are not loaded
    // for the last person opened in HR Files.
    this.leaveAllFilesTaskView();
    this.showToDoSection = true;
    this.showAllFilesSection = false;
    this.hideComments = true;
    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    Object.assign(this, getTodoTabMobileState());
    this.invalidateCommentFilters();
    this.syncMobileSearchInput();
    // Instant paint from cache, then finish any interrupted load (never treat a
    // partial/cancelled snapshot as final - that caused intermittent missing tasks).
    this.restoreTodoTasksFromCache();
    if (!this.isLoadingTodoTasks) {
      this.ensureTodoLoadComplete();
    }
    this.refreshView();
  }

  /**
   * Pull the next slice of older rows from every To Do list that still has a cursor.
   * Nothing is re-fetched: each list resumes exactly where its first pass stopped.
   */
  protected async loadMoreTodoTasks(): Promise<void> {
    if (this.isLoadingMoreTodoTasks || !this.hasMoreTodoTasks) return;

    this.isLoadingMoreTodoTasks = true;
    this.refreshView();

    try {
      await this.fetchMoreTodoFromCursors(AppConstants.todoLoadMorePagesPerList);
    } catch {
      // Best effort - the button stays available so the user can retry.
    } finally {
      this.isLoadingMoreTodoTasks = false;
      this.refreshView();
    }
  }

  /** Resume list cursors for `pagesPerList` more Graph pages and merge into To Do. */
  private async fetchMoreTodoFromCursors(
    pagesPerList: number,
    options: { procurementOnly?: boolean } = {},
  ): Promise<void> {
    if (!this.hasMoreTodoTasks) return;

    const token = await this.getSharePointToken();
    const siteId = this.cachedSiteId;
    const siteListsArr = this.cachedSiteLists;
    if (!siteId || siteListsArr.length === 0) return;

    const siteWebUrl = this.cachedSiteWebUrl;
    const userEmail = (this.currentUser?.email ?? '').toLowerCase();
    const userUpn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();

    const pendingLists = [...this.todoListCursors.entries()]
      .filter((entry): entry is [string, string] => !!entry[1])
      .filter(([listName]) =>
        !options.procurementOnly || isTodoProcurementTaskList(listName),
      );

    if (pendingLists.length === 0) return;

    const fetchOne = async ([listName, cursor]: [string, string]): Promise<any[]> => {
      const list = siteListsArr.find(
        (l: any) =>
          (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
          (l.displayName ?? '').toLowerCase() === listName.toLowerCase(),
      );
      if (!list?.id) {
        this.todoListCursors.delete(listName);
        return [];
      }

      try {
        const raw = isTodoProcurementTaskList(listName)
          ? await this.fetchProcurementTodoPendingItems(
              siteId, list.id, token, listName, cursor, pagesPerList,
            )
          : await this.fetchTodoScanPages(
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
        this.todoListCursors.delete(listName);
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
      const existing = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
      const merged = this.mergeHrTasks(existing, fetched);
      this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, merged);
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), { clearLoading: false });
      this.userHrTasksLoaded = true;
      this.refreshView();
    }
  }

  /** Restore To Do from cache (memory or persisted snapshot) - instant when switching back or reopening the tab. */
  private restoreTodoTasksFromCache(): boolean {
    const cached = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY);
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
  private ensureTodoLoadComplete(): void {
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
  private async resumeTodoProcurementDrain(): Promise<void> {
    if (!this.hasMoreProcurementTodoCursors()) {
      this.todoScopeCacheValid = this.todoFirstPassComplete;
      return;
    }

    const loadSeq = ++this.todoTaskLoadSeq;
    // Background only - do not re-show "Loading more tasks..." over an already painted list.
    this.pauseFolderPrefetch();
    this.refreshView();
    try {
      await this.drainProcurementTodoInBackground(loadSeq);
      if (!this.isStaleTodoTaskLoad(loadSeq) && this.todoFirstPassComplete) {
        this.todoScopeCacheValid = !this.hasMoreProcurementTodoCursors();
        void this.warmHrPersonalRootFoldersCache();
      }
    } finally {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.folderPrefetchPaused = false;
        this.refreshView();
      }
    }
  }

  /** Load HR tasks assigned to the logged-in user into the To Do list. */
  protected async loadTodoTasksForCurrentUser(options: { forceRefresh?: boolean } = {}): Promise<void> {
    if (!this.currentUser) return;

    const forceRefresh = options.forceRefresh === true;
    // Stale (even day-old, persisted) tasks paint instantly; the fetch below refreshes them.
    const cached = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY);
    const hasCachedTasks = Array.isArray(cached) && cached.length > 0;

    if (hasCachedTasks && !forceRefresh) {
      this.restoreTodoTasksFromCache();
      this.refreshView();
    }

    const loadSeq = ++this.todoTaskLoadSeq;
    this.isLoadingTodoTasks = true;
    this.todoFirstPassComplete = false;
    this.todoScopeCacheValid = false;
    // Yield Graph capacity to To Do - idle folder children crawls compete for the same budget.
    this.pauseFolderPrefetch();
    // A fresh load re-pages each list from the top, so old continuations are stale.
    this.todoListCursors.clear();

    if (hasCachedTasks) {
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(cached));
      this.userHrTasksLoaded = true;
      this.refreshView();
    } else {
      this.todoService.setLoading(true);
      this.refreshView();
    }

    try {
      await this.ensureFormConfigLoaded();
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
      this.refreshView();

      // Limited fallback for lists that reject AssignedToLookupId (not procurement).
      if (this.taskQuery.blockedAssigneeLookupListIds.size > 0) {
        const fallbackItems = await this.loadHrTasksFromLists(
          false,
          { createdSinceIso: sinceIso },
          {
            updateTodo: false,
            updateCommentItems: false,
            kickoffOlderBackfill: false,
            skipCache: true,
            progressiveTodo: false,
            todoLoadSeq: loadSeq,
            todoAssigneeOnly: false,
            restrictToAssigneeBlockedLists: true,
            fastLoadPageLimit: AppConstants.hrTasksTodoLoadPageLimit,
            todoScope: true,
          }
        );
        if (!this.isStaleTodoTaskLoad(loadSeq) && fallbackItems.length > 0) {
          const cachedItems = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
          const merged = this.mergeHrTasks(cachedItems, fallbackItems);
          this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, merged);
          this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), { clearLoading: false });
          this.refreshView();
        }
      }

      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        const folderAssigned = this.collectAssignedProcTasksFromFolderCaches();
        if (folderAssigned.length > 0) {
          const cachedItems = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
          const merged = this.mergeHrTasks(cachedItems, folderAssigned);
          this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, merged);
          this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged), { clearLoading: false });
          this.refreshView();
        }
      }

      if (this.isStaleTodoTaskLoad(loadSeq)) return;

      // First Graph pass done - paint without holding "Loading more tasks...".
      // Remaining Proc* pages continue in the background drain below.
      this.todoFirstPassComplete = true;
      this.isLoadingTodoTasks = false;
      this.todoService.setLoading(false);
      this.refreshView();

      await this.drainProcurementTodoInBackground(loadSeq);
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.todoScopeCacheValid = !this.hasMoreProcurementTodoCursors();
        this.userHrTasksLoaded = true;
        // Warm HR Files people list only after To Do Graph work settles.
        void this.warmHrPersonalRootFoldersCache();
      }
    } catch {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.todoService.setLoading(false);
      }
    } finally {
      if (!this.isStaleTodoTaskLoad(loadSeq)) {
        this.isLoadingTodoTasks = false;
        this.todoService.setLoading(false);
        this.folderPrefetchPaused = false;
        this.refreshView();
      }
    }
  }

  /** Continue Progress=Pending procurement cursors without holding the To Do spinner. */
  private async drainProcurementTodoInBackground(loadSeq: number): Promise<void> {
    this.todoProcurementDrainActive = true;
    try {
      while (!this.isStaleTodoTaskLoad(loadSeq) && this.hasMoreProcurementTodoCursors()) {
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

  private isStaleTodoTaskLoad(loadSeq: number): boolean {
    return loadSeq !== this.todoTaskLoadSeq;
  }

  /** Stop in-flight To Do fetches so HR Files / All Files get Graph quota. */
  private cancelBackgroundTodoLoad(): void {
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
      this.refreshView();
    }
  }

  protected toggleAllFilesSection(): void {
    this.cancelBackgroundTodoLoad();
    this.cancelPendingFolderTaskLoads();
    this.selectedHrPersonalTaskFolder = '';
    this.allFilesMounted = true;
    this.showAllFilesSection = true;
    this.showToDoSection = false;
    this.hideComments = true;
    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    this.selectedFolderName = '';
    this.showUserFile = false;
    this.userFiles = [];
    this._folderMapDirty = true;
    this.folderStack = [];
    this.attachmentBrowsingRoot = null;
    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.userRootFolderId = null;
    Object.assign(this, getAllFilesTabMobileState());
    this.invalidateCommentFilters();
    this.syncMobileSearchInput();
    this.refreshView();
  }

  protected onMobileSearchChange(term: string): void {
    const result = resolveMobileSearchChange(
      term,
      this.showToDoSection,
      this.showAllFilesSection
    );
    this.searchTerm = result.searchTerm;

    if (result.section === 'allFiles') {
      this.allFilesComponent?.searchFiles(this.searchTerm);
      return;
    }

    if (result.section === 'todo') {
      this.todoListComponent?.searchTasks(this.searchTerm);
      return;
    }

    this.searchQuery = result.searchQuery ?? this.searchTerm;
    this.invalidateCommentFilters();
    if (result.navigateToComments) {
      this.hideComments = false;
      Object.assign(this, getCommentsTabMobileState());
    }
    this.refreshView();
  }

  private syncMobileSearchInput(): void {
    this.searchTerm = syncMobileSearchTerm({
      showToDoSection: this.showToDoSection,
      showAllFilesSection: this.showAllFilesSection,
      searchQuery: this.searchQuery,
      todoSearchTerm: this.todoListComponent?.searchTerm,
      allFilesSearchTerm: this.allFilesComponent?.searchTerm,
    });
  }

  protected onSuperiorModeChange(checked: boolean): void {
    this.superiorMode = checked;
    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    this.invalidateCommentFilters();
    if (checked) {
      void this.reloadTodoTasksWithSubordinates();
    } else {
      void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
    }
    this.refreshView();
  }

  private async reloadTodoTasksWithSubordinates(): Promise<void> {
    if (!this.currentUser) return;

    const identifiers = [
      this.currentUser.email,
      this.currentUser.userPrincipalName,
      this.currentUser.username,
    ].map(value => (value ?? '').trim()).filter(Boolean);

    if (!identifiers.length) return;

    const loadSeq = ++this.todoTaskLoadSeq;
    this.isLoadingTodoTasks = true;
    this.pauseFolderPrefetch();
    this.todoService.setLoading(true);
    this.refreshView();

    const since = new Date();
    since.setUTCDate(since.getUTCDate() - AppConstants.hrTasksRecentFetchDays);

    try {
      await this.subordinaryTaskService.ensureSubordinatesForManager(identifiers, false);
      if (this.isStaleTodoTaskLoad(loadSeq)) return;

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
        this.folderPrefetchPaused = false;
        this.refreshView();
      }
    }
  }

  // Receives submitter + eFormListId + eFormTitle from todo-list after a grouped row is clicked.
  protected onSubmitterSelected(selection: SubmitterSelection | null): void {
    this.selectedSubmitter = selection?.name ?? null;
    this.selectedEFormListId = selection?.eFormListId ?? null;
    this.selectedEFormTitle = selection?.eFormTitle ?? null;
    this.selectedGroupTaskIds = (selection?.taskIds ?? [])
      .map(id => String(id ?? '').trim())
      .filter(Boolean);

    const rawTaskId = String(selection?.taskId ?? '').trim();
    const numericTaskId = Number(rawTaskId);
    this.selectedTask = rawTaskId
      ? { Id: !Number.isNaN(numericTaskId) ? numericTaskId : rawTaskId }
      : null;

    const task = this.todoService.getTaskById(rawTaskId);
    const commentItem = this.commentItems.find(item => String(item.id ?? '').trim() === rawTaskId);
    const groupedItem = this.commentItems.find(
      item =>
        String(item.eFormDetails?.eFormListId ?? '').trim() ===
        String(selection?.eFormListId ?? '').trim(),
    );
    this.selectedCommentListName =
      String((commentItem as { listName?: string })?.listName ?? '').trim() ||
      String(commentItem?.eFormDetails?.listName ?? '').trim() ||
      String(groupedItem?.eFormDetails?.listName ?? (groupedItem as { listName?: string })?.listName ?? '').trim() ||
      String(task?.eFormDetails?.listName ?? '').trim();

    this.invalidateCommentFilters();
    if (selection) {
      if (this.showToDoSection) {
        this.selectedHrPersonalTaskFolder = '';
        this.selectedFolderName = '';
      }
      this.hideComments = false;
      Object.assign(this, getCommentsTabMobileState());
      this.closeFormViewer();
      // Comments panel filters locally. Only hit Graph when we have nothing cached
      // (every accordion click used to call loadHrTasks and re-fan-out all lists).
      const hasLocalComments = this.commentItems.length > 0;
      const cachedTasks = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY);
      if (!hasLocalComments && Array.isArray(cachedTasks) && cachedTasks.length > 0) {
        this.commentItems = this.sortCommentItemsByDateDesc(cachedTasks);
        this.visibleItemCount = this.commentsInitialPageSize;
        this.invalidateCommentFilters();
        this.isLoadingComments = false;
        this.commentsMessage = '';
        this.userHrTasksLoaded = true;
      } else if (!hasLocalComments && !this.userHrTasksLoaded && !this.isLoadingComments) {
        this.loadHrTasks();
      }
    } else {
      this.closeFormViewer();
    }
    this.refreshView();
  }

  /** Resolve the SharePoint tasks list name for a document library. */
  private getTaskListForLibrary(libraryName: string): string | undefined {
    const map = this.documentLibraryTaskMap;
    if (!libraryName || !map) return undefined;

    if (map[libraryName]) return map[libraryName];

    // Case-insensitive fallback: DocLibraryName casing in the eForms list may differ
    // from the SharePoint drive (library) name.
    const target = libraryName.toLowerCase();
    const matchKey = Object.keys(map).find((key) => key.toLowerCase() === target);
    return matchKey ? map[matchKey] : undefined;
  }

  /**
    this is for the proctasks in archive to load the tasks from the ProcTasksArchive list
   */
  private getArchiveCompanionTaskLists(listName: string): string[] {
    const key = String(listName ?? '').trim().toLowerCase();
    if (key === 'proctasks' || key === 'proctasks1') {
      return ['ProcTasksArchive'];
    }
    return [];
  }

  private findCachedSiteListByName(listName: string): { id: string; name?: string; displayName?: string; webUrl?: string } | null {
    const target = String(listName ?? '').trim().toLowerCase();
    if (!target) return null;
    return this.cachedSiteLists.find((l: any) =>
      (l.name ?? '').toLowerCase() === target ||
      (l.displayName ?? '').toLowerCase() === target
    ) ?? null;
  }

  /** Show the Comments panel and load tasks for the clicked folder. */
  private loadAllFilesTasksForLibrary(libraryName: string, folderLabel: string): void {
    this.cancelPendingFolderTaskLoads();
    const loadSeq = ++this.allFilesFolderTaskLoadSeq;
    const loadFolder = String(folderLabel || '').trim();

    // Switching folders: drop the previous folder's rows so a slow cold load shows a
    // spinner instead of the old folder's tasks under the new folder's badge.
    // Same folder (search-hit seed, refresh) keeps what is already painted.
    if (String(this.selectedFolderName ?? '').trim() !== loadFolder) {
      this.commentItems = [];
      this.visibleItemCount = this.commentsInitialPageSize;
      this.invalidateCommentFilters(true);
    }

    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    this.hideComments = false;
    this.selectedFolderName = loadFolder;
    if (normalizeName(libraryName) === normalizeName(this.targetLibraryName)) {
      this.selectedHrPersonalTaskFolder = loadFolder;
    } else {
      this.selectedHrPersonalTaskFolder = '';
    }
    // Do not force Comments tab here - folder clicks open Attachments first so
    // files appear immediately while tasks load in the background.

    // Serve from cache when the same folder was crawled recently.
    if (normalizeName(libraryName) === normalizeName(this.targetLibraryName)) {
      this.lastHrFilesCommentsFolder = loadFolder;
      const cacheKey = this.hrFolderTaskCacheKey(loadFolder);
      const fresh = this.fileCrawlCache.get(cacheKey);
      const cached = fresh ?? this.fileCrawlCache.getStale(cacheKey);
      // Paint cache instantly, then soft-refresh from Graph so new tasks appear —
      // unless a full crawl for this person just finished (rapid re-clicks).
      if (Array.isArray(cached) && cached.length > 0) {
        const completeAt = this.hrFolderCompleteAt.get(cacheKey);
        const needsRefresh =
          completeAt === undefined ||
          Date.now() - completeAt > AppConstants.hrFilesSkipSoftRefreshMs ||
          this.hrFolderPartialKeys.has(cacheKey);
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
      const mappedTaskList = this.getTaskListForLibrary(libraryName);
      if (mappedTaskList) {
        const cacheKey = this.libFolderTaskCacheKey(mappedTaskList, loadFolder);
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

    const mappedTaskList = this.getTaskListForLibrary(libraryName);
    if (mappedTaskList) {
      void this.loadDocumentLibraryTasksForFolder(mappedTaskList, loadFolder, loadSeq);
      return;
    }

    if (loadSeq !== this.allFilesFolderTaskLoadSeq) return;
    this.isLoadingComments = false;
    this.commentsMessage = `No task list configured for ${libraryName}.`;
    if (this.selectedAllFilesFolderId) {
      this.allFilesComponent?.setChildSubjectFolderNames(this.selectedAllFilesFolderId, []);
    }
    this.refreshView();
  }

  private beginFolderTaskLoadingUi(): void {
    // Keep already-painted Comments visible (seeded search hit or prior cache) —
    // only full-screen spinner when the panel is empty.
    if (this.commentItems.length === 0) {
      this.commentsMessage = 'Loading tasks...';
      this.isLoadingComments = true;
      this.isLoadingMoreComments = false;
    } else {
      this.isLoadingComments = false;
      this.isLoadingMoreComments = true;
    }
    this.refreshView();
  }

  /**
   * Doc-lib folder tasks: try shared Redis (written by any signed-in user) before Graph.
   * Spinner only after Redis miss so cached folders feel instant.
   */
  private async loadLibFolderTasksWithSharedCache(
    listName: string,
    folderName: string,
    loadSeq: number,
  ): Promise<void> {
    const cacheKey = this.libFolderTaskCacheKey(listName, folderName);
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
  private async loadHrFolderTasksWithSharedCache(
    folderName: string,
    loadSeq: number,
  ): Promise<void> {
    const cacheKey = this.hrFolderTaskCacheKey(folderName);
    const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
    if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

    // Array hit from shared Redis — paint then soft-refresh for newest tasks.
    if (hydrated && Array.isArray(hydrated.data)) {
      this.applyCachedFolderTaskItems(hydrated.data, folderName);
      void this.loadHrPersonalTasksForFolder(folderName, loadSeq, { softRefresh: true });
      return;
    }

    this.beginFolderTaskLoadingUi();
    await this.loadHrPersonalTasksForFolder(folderName, loadSeq);
  }

  /**
   * When opening a ChildSubject folder, parent Comments may already be cached.
   * Filter those rows by the child's trailing eForm id instead of waiting on a full list scan.
   */
  private tryReuseParentLibFolderTaskCache(listName: string, childFolderName: string): any[] | null {
    const parentName = String(this.selectedAllFilesParentFolderName ?? '').trim();
    if (!parentName || normalizeName(parentName) === normalizeName(childFolderName)) {
      return null;
    }
    const childEFormId = extractTrailingFolderId(childFolderName);
    if (!childEFormId) return null;

    const parentCached = this.fileCrawlCache.getStale(this.libFolderTaskCacheKey(listName, parentName));
    if (!Array.isArray(parentCached) || parentCached.length === 0) return null;

    const matched = parentCached.filter((item) => this.folderMatch.doesMappedHrTaskMatchFolder(item, childFolderName));
    if (matched.length === 0) return null;

    const sorted = this.sortCommentItemsByDateDesc(matched);
    this.fileCrawlCache.set(this.libFolderTaskCacheKey(listName, childFolderName), sorted);
    return sorted;
  }

  private isStaleAllFilesFolderTaskLoad(loadSeq: number): boolean {
    return loadSeq !== this.allFilesFolderTaskLoadSeq;
  }

  /** Called when a file/folder is clicked in the All Files view. */
  protected onAllFileSelected(file: any): void {
    if (!file) return;
    this.pendingCommentFocusId = null;
    this.highlightedCommentId = null;
    this.pendingAttachmentFocusId = null;
    this.highlightedAttachmentId = null;

    // Folder click: load that folder's children into Attachments (not the whole library).
    if (file.isfolder) {
      const driveId = file.driveId ?? this.currentLibraryDriveId;
      if (driveId) {
        this.selectedAllFilesFolderId = String(file.id ?? '');
        this.selectedAllFilesParentFolderName = String(file.parentFolderName ?? '').trim();
        this.childSubjectSourceFetchedForFolderId = null;
        this.allFilesComponent?.beginChildSubjectLoad(this.selectedAllFilesFolderId);
        this.pauseFolderPrefetch();
        this.currentLibraryDriveId = driveId;
        this.attachmentBrowsingRoot = { id: file.id, name: file.name, webUrl: file.webUrl ?? '' };
        this.folderStack = [];
        this.currentFolderName = file.name;
        // Show Attachments immediately - tasks load in the background for Comments.
        this.hideComments = false;
        Object.assign(this, getAttachmentsTabMobileState());
        void this.loadDriveFolderContents(driveId, file.id, file.webUrl, file.name);

        const libName = file.libraryName ?? this.currentLibraryName;
        if (libName) {
          this.selectedFolderModifiedBy = String(file.modifiedby ?? '').trim();
          // ChildSubject accordion: fetch source eForm in parallel (don't wait for task scan).
          void this.primeChildSubjectForFolder(libName, String(file.name || ''), this.selectedAllFilesFolderId);
          this.loadAllFilesTasksForLibrary(libName, String(file.name || ''));
        }

        return;
      }
      this.userFiles = [{
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
      this._folderMapDirty = true;
      this.showUserFile = true;
      this.refreshView();
      return;
    }

    // Non-folder: show single file entry in attachments panel
    this.userFiles = [{
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
    this._folderMapDirty = true;
    this.showUserFile = true;
    this.refreshView();
  }

  /**
   * All Files → Comments search row: open the related folder and focus that
   * comment/task in the Comments panel (Live behaviour).
   */
  protected onAllFilesCommentSelected(hit: CommentSearchHit): void {
    if (!hit?.folder) return;

    const file = hit.folder;
    const taskId = String(hit.taskId ?? '').trim();
    const driveId = file.driveId ?? this.currentLibraryDriveId;
    if (!driveId) return;

    const libName = String(file.libraryName ?? this.currentLibraryName ?? '').trim();
    const folderId = String(file.id ?? '');
    const folderName = String(file.name || '');
    const folderWebUrl = file.webUrl ?? '';

    // Stop background Comments crawling so the click isn't competing with Graph.
    this.allFilesComponent?.pauseCommentTasksWarm();

    this.pendingCommentFocusId = taskId || null;
    this.highlightedCommentId = taskId || null;
    this.pendingAttachmentFocusId = null;
    this.highlightedAttachmentId = null;
    this.selectedAllFilesFolderId = folderId;
    this.selectedAllFilesParentFolderName = String(file.parentFolderName ?? '').trim();
    this.childSubjectSourceFetchedForFolderId = null;
    this.currentLibraryDriveId = driveId;
    this.attachmentBrowsingRoot = { id: file.id, name: file.name, webUrl: folderWebUrl };
    this.folderStack = [];
    this.currentFolderName = folderName;
    this.selectedFolderModifiedBy = String(file.modifiedby ?? '').trim();
    this.hideComments = false;
    Object.assign(this, getCommentsTabMobileState());

    // Paint only the clicked comment immediately — never scan the full search index here.
    this.seedCommentsFromSearchHit(hit);
    this.highlightedCommentId = taskId || null;
    this.pendingCommentFocusId = taskId || null;
    this.focusPendingCommentIfNeeded();
    // Keep focus armed for when the full folder Comments list replaces the seed.
    this.pendingCommentFocusId = taskId || null;

    // Defer Attachments + full folder Comments so the UI stays responsive on click.
    setTimeout(() => {
      this.pauseFolderPrefetch();
      this.allFilesComponent?.beginChildSubjectLoad(folderId);
      void this.loadDriveFolderContents(driveId, folderId, folderWebUrl, folderName);
      if (libName) {
        void this.primeChildSubjectForFolder(libName, folderName, folderId);
        this.loadAllFilesTasksForLibrary(libName, folderName);
      }
      // Resume indexing after the click-driven Graph work has had time to start.
      setTimeout(() => this.allFilesComponent?.resumeCommentTasksWarm(), 2500);
    }, 0);
  }

  /**
   * All Files → Attachments search row: open the related folder and highlight
   * that file in the Attachments panel (Live behaviour).
   */
  protected onAllFilesAttachmentSelected(hit: AttachmentSearchHit): void {
    if (!hit?.folder) return;

    const file = hit.folder;
    const fileId = String(hit.fileId ?? '').trim();
    const driveId = file.driveId ?? this.currentLibraryDriveId;
    if (!driveId) return;

    const libName = String(file.libraryName ?? this.currentLibraryName ?? '').trim();
    const folderId = String(file.id ?? '');
    const folderName = String(file.name || '');
    const folderWebUrl = file.webUrl ?? '';

    this.pendingCommentFocusId = null;
    this.highlightedCommentId = null;
    this.pendingAttachmentFocusId = fileId || null;
    this.highlightedAttachmentId = fileId || null;
    this.selectedAllFilesFolderId = folderId;
    this.selectedAllFilesParentFolderName = String(file.parentFolderName ?? '').trim();
    this.childSubjectSourceFetchedForFolderId = null;
    this.currentLibraryDriveId = driveId;
    this.attachmentBrowsingRoot = { id: file.id, name: file.name, webUrl: folderWebUrl };
    this.folderStack = [];
    this.currentFolderName = folderName;
    this.selectedFolderModifiedBy = String(file.modifiedby ?? hit.author ?? '').trim();
    this.hideComments = false;
    Object.assign(this, getAttachmentsTabMobileState());

    // Instant paint: show the clicked file while the full folder listing loads.
    this.userFiles = [{
      id: hit.fileId,
      name: hit.fileName,
      webUrl: hit.webUrl ?? '',
      lastModifiedDateTime: hit.date || undefined,
      isFolder: false,
      modifiedBy: hit.author || undefined,
      fileIcon: hit.fileIcon || undefined,
    }];
    this._folderMapDirty = true;
    this.showUserFile = true;
    this.isLoadingUserFiles = false;
    this.refreshView();

    setTimeout(() => {
      this.pauseFolderPrefetch();
      this.allFilesComponent?.beginChildSubjectLoad(folderId);
      void this.loadDriveFolderContents(driveId, folderId, folderWebUrl, folderName).then(() => {
        this.focusPendingAttachmentIfNeeded();
      });
      if (libName) {
        void this.primeChildSubjectForFolder(libName, folderName, folderId);
        this.loadAllFilesTasksForLibrary(libName, folderName);
      }
    }, 0);
  }

  private focusPendingAttachmentIfNeeded(): void {
    const id = String(this.pendingAttachmentFocusId ?? '').trim();
    if (!id) return;
    this.highlightedAttachmentId = id;
    this.pendingAttachmentFocusId = null;
    this.refreshView();
    setTimeout(() => {
      const el = document.getElementById(`attachment-item-${id}`);
      el?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 50);
  }

  /** Instant Comments paint: only the clicked search hit (full folder load follows async). */
  private seedCommentsFromSearchHit(hit: CommentSearchHit): void {
    const folderName = String(hit.folder?.name ?? '').trim();
    if (!folderName) return;

    this.selectedSubmitter = null;
    this.selectedEFormListId = null;
    this.selectedEFormTitle = null;
    this.selectedFolderName = folderName;
    this.hideComments = false;

    const libName = String(hit.folder?.libraryName ?? hit.libraryName ?? '').trim();
    const listName = (libName && this.getTaskListForLibrary(libName)) || libName || 'Tasks';
    const item = buildQuickCommentItemFromSearchHit(hit, listName);

    this.isLoadingComments = false;
    this.isLoadingMoreComments = false;
    this.commentsMessage = '';
    this.commentItems = [item];
    this.visibleItemCount = this.commentsInitialPageSize;
    this.invalidateCommentFilters(true);
    this.refreshView();
  }


  /**
   * Quietly cache Comments for an All Files folder during Comments search.
   * Does not change the open Comments panel — only warms tasks:lib/hr caches.
   */
  protected onEnsureFolderComments(req: { libraryName: string; folderName: string }): void {
    const libraryName = String(req?.libraryName ?? '').trim();
    const folderName = String(req?.folderName ?? '').trim();
    if (!libraryName || !folderName) return;
    void this.silentWarmFolderCommentsForSearch(libraryName, folderName);
  }

  /**
   * Quietly cache Attachments for an All Files folder during Attachments search.
   * Does not change the open Attachments panel — only warms files:folder caches.
   */
  protected onEnsureFolderAttachments(req: { driveId: string; folderId: string }): void {
    const driveId = String(req?.driveId ?? '').trim();
    const folderId = String(req?.folderId ?? '').trim();
    if (!driveId || !folderId) return;
    void this.silentWarmFolderAttachmentsForSearch(driveId, folderId);
  }

  private async silentWarmFolderAttachmentsForSearch(
    driveId: string,
    folderId: string,
  ): Promise<void> {
    const cacheKey = this.driveFolderCacheKey(driveId, folderId);
    try {
      const existing = this.fileCrawlCache.getStale(cacheKey);
      if (Array.isArray(existing)) {
        this.allFilesComponent?.refreshAttachmentSearchFromCaches();
        return;
      }

      const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
      if (hydrated && Array.isArray(hydrated.data)) {
        this.fileCrawlCache.set(cacheKey, hydrated.data);
        this.allFilesComponent?.refreshAttachmentSearchFromCaches();
        if (hydrated.fresh) return;
      }

      if (this.folderPrefetchInFlight.has(cacheKey)) return;
      this.folderPrefetchInFlight.add(cacheKey);
      try {
        const token = await this.getSharePointToken();
        const mapped = await this.fetchDriveFolderChildrenMapped(driveId, folderId, token);
        this.fileCrawlCache.set(cacheKey, mapped);
        this.allFilesComponent?.refreshAttachmentSearchFromCaches();
      } finally {
        this.folderPrefetchInFlight.delete(cacheKey);
      }
    } catch (err) {
      console.warn(`[Attachments search] Failed to warm folder ${folderId}`, err);
    }
  }

  private async silentWarmFolderCommentsForSearch(
    libraryName: string,
    folderName: string,
  ): Promise<void> {
    try {
      if (normalizeName(libraryName) === normalizeName(this.targetLibraryName)) {
        const cacheKey = this.hrFolderTaskCacheKey(folderName);
        const existing = this.fileCrawlCache.getStale(cacheKey);
        if (Array.isArray(existing) && existing.length > 0) {
          this.allFilesComponent?.refreshCommentSearchFromCaches();
          return;
        }
        const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
        if (hydrated && Array.isArray(hydrated.data)) {
          this.fileCrawlCache.set(cacheKey, hydrated.data);
          this.allFilesComponent?.refreshCommentSearchFromCaches();
        }
        // Skip Graph for HR during search warm — opening the folder loads Comments normally.
        return;
      }

      const mappedTaskList = this.getTaskListForLibrary(libraryName);
      if (!mappedTaskList) return;

      const cacheKey = this.libFolderTaskCacheKey(mappedTaskList, folderName);
      const existing = this.fileCrawlCache.getStale(cacheKey);
      if (Array.isArray(existing) && existing.length > 0) {
        this.allFilesComponent?.refreshCommentSearchFromCaches();
        return;
      }

      const hydrated = await this.fileCrawlCache.hydrateFromPersistent<any[]>(cacheKey);
      if (hydrated && Array.isArray(hydrated.data)) {
        this.fileCrawlCache.set(cacheKey, hydrated.data);
        this.allFilesComponent?.refreshCommentSearchFromCaches();
        if (hydrated.fresh) return;
      }

      await this.loadDocumentLibraryTasksForFolder(
        mappedTaskList,
        folderName,
        this.allFilesFolderTaskLoadSeq,
        { silent: true },
      );
    } catch (err) {
      console.warn(`[Comments search] Failed to warm "${folderName}"`, err);
    }
  }

  /** Remember which library is active in All Files - do not preload attachments. */
  protected onLibrarySelected(event: { name: string; driveId: string } | null): void {
    if (!event) return;
    const { name: libraryName, driveId } = event;

    this.currentLibraryName = libraryName;
    this.currentLibraryDriveId = driveId;
    this.folderStack = [];
    this.attachmentBrowsingRoot = null;
    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.userRootFolderId = null;
    this.showUserFile = false;
    this.userFiles = [];
    this._folderMapDirty = true;

    // Tasks and attachments load only when a folder is clicked
    this.selectedFolderName = '';
    this.selectedFolderModifiedBy = '';
    this.commentItems = [];
    this.commentsMessage = '';
    this.hideComments = true;
    this.invalidateCommentFilters();
    this.refreshView();
  }


  /** Load one-level children of a drive folder and show them in Attachments. */
  protected async loadDriveFolderContents(
    driveId: string,
    folderId: string,
    folderWebUrl?: string,
    folderName?: string,
    options: { skipCache?: boolean } = {},
  ): Promise<void> {
    if (!driveId || !folderId) return;
    const seq = ++this.driveFolderLoadSeq;
    const cacheKey = this.driveFolderCacheKey(driveId, folderId);

    // Document-library browsing - do not mix with the HR Personal drive tree.
    this.userRootFolderId = null;

    let paintedFromStaleCache = false;
    if (!options.skipCache) {
      const cached = this.fileCrawlCache.get(cacheKey);
      if (cached) {
        if (seq !== this.driveFolderLoadSeq) return;
        this.applyDriveFolderContentsToUi(cached, folderId, folderWebUrl, folderName);
        this.schedulePrefetchChildFolders(driveId, cached);
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
          if (hydrated.fresh) {
            this.schedulePrefetchChildFolders(driveId, hydrated.data);
            return;
          }
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
    this.refreshView();
    try {
      const token = await this.getSharePointToken();
      const mapped = await this.fetchDriveFolderChildrenMapped(driveId, folderId, token);
      if (seq !== this.driveFolderLoadSeq) return;

      this.fileCrawlCache.set(cacheKey, mapped);
      this.applyDriveFolderContentsToUi(mapped, folderId, folderWebUrl, folderName);
      this.schedulePrefetchChildFolders(driveId, mapped);
    } catch (err: any) {
      if (seq !== this.driveFolderLoadSeq) return;
      this.userFileError = err?.message || 'Failed to load folder contents';
    } finally {
      if (seq !== this.driveFolderLoadSeq) return;
      this.isLoadingUserFiles = false;
      this.refreshView();
    }
  }

  private applyDriveFolderContentsToUi(
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
    if (folderId === this.selectedAllFilesFolderId) {
      this.allFilesComponent?.refreshChildFoldersForFolder(folderId);
    }
    this.allFilesComponent?.refreshAttachmentSearchFromCaches();
    this.focusPendingAttachmentIfNeeded();
    this.refreshView();
  }

  private async fetchDriveFolderChildrenMapped(
    driveId: string,
    folderId: string,
    token: string,
  ): Promise<any[]> {
    const path =
      `/drives/${driveId}/items/${folderId}/children` +
      `?$select=id,name,webUrl,lastModifiedDateTime,lastModifiedBy,size,file,folder,parentReference&$top=500`;
    const resp: any = await graphGetWithRetry(
      this.http,
      path,
      token,
      AppConstants.graphFileListingTimeoutMs,
    );
    const items = resp?.value ?? [];
    return items
      .filter((item: any) => !!item.file || !!item.folder)
      .map((item: any) => {
        const ext = item.file ? getFileExtension(item.name) : '';
        const category = item.file ? getFileCategory(ext) : '';
        const icon = item.file ? getFileIcon(category, ext) : '';
        return {
          id: item.id,
          name: item.name,
          webUrl: item.webUrl ?? '',
          parentId: item.parentReference?.id ?? folderId,
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
  }

  /** Idle-prefetch visible All Files folders so the first click is often already cached. */
  protected onFoldersForPrefetch(
    folders: Array<{ driveId: string; folderId: string }>,
  ): void {
    // Disabled: idle prefetch was a major Graph spam source on every folder paint / click.
    return;
  }

  private pauseFolderPrefetch(): void {
    this.folderPrefetchPaused = true;
    if (this.folderPrefetchTimer) {
      clearTimeout(this.folderPrefetchTimer);
      this.folderPrefetchTimer = null;
    }
    if (this.hrTaskPrefetchTimer) {
      clearTimeout(this.hrTaskPrefetchTimer);
      this.hrTaskPrefetchTimer = null;
    }
    if (this.hrPersonHoverPrefetchTimer) {
      clearTimeout(this.hrPersonHoverPrefetchTimer);
      this.hrPersonHoverPrefetchTimer = null;
    }
    // Resume after click-path work; stay paused while To Do still owns Graph quota.
    setTimeout(() => {
      if (this.isLoadingTodoTasks) return;
      this.folderPrefetchPaused = false;
    }, 8000);
  }

  /**
   * Idle-prefetch Comments for visible HR Files people so the first click is often cached
   * (same idea as All Files folder content prefetch).
   * Capped batch — full-list prefetch flooded Graph and caused 429s.
   */
  private scheduleHrFolderTasksPrefetch(
    folders: Array<{ name?: string }>,
  ): void {
    const names = folders
      .map(f => String(f?.name ?? '').trim())
      .filter(Boolean)
      .filter(name => !this.fileCrawlCache.get(this.hrFolderTaskCacheKey(name)))
      .filter(name => !this.hrTaskPrefetchInFlight.has(this.hrFolderTaskCacheKey(name)))
      .slice(0, AppConstants.hrFilesTaskPrefetchMaxBatch);
    if (names.length === 0) return;

    if (this.hrTaskPrefetchTimer) {
      clearTimeout(this.hrTaskPrefetchTimer);
    }
    this.hrTaskPrefetchTimer = setTimeout(() => {
      this.hrTaskPrefetchTimer = null;
      if (this.folderPrefetchPaused) return;
      if (this.isLoadingComments || this.isLoadingMoreComments) return;
      void this.prefetchHrFolderTasks(names);
    }, AppConstants.hrFilesTaskPrefetchDelayMs);
  }

  /** Hover / pointerdown: start warming this person's Comments before the click. */
  protected prefetchHrPersonComments(item: { name?: string } | null | undefined): void {
    const name = String(item?.name ?? '').trim();
    if (!name) return;
    const key = this.hrFolderTaskCacheKey(name);
    if (this.fileCrawlCache.get(key) || this.hrTaskPrefetchInFlight.has(key)) return;
    if (this.folderPrefetchPaused && this.isLoadingTodoTasks && !this.todoFirstPassComplete) {
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
  private async prefetchHrFolderTasks(folderNames: string[]): Promise<void> {
    if (this.folderPrefetchPaused || folderNames.length === 0) return;
    if (this.isLoadingHrFilesList) return;
    // Allow during To Do background drain — only skip while the first To Do pass
    // still owns the spinner (pauseFolderPrefetch covers that window).
    if (this.isLoadingTodoTasks && !this.todoFirstPassComplete) return;

    const pending = folderNames.filter(name => {
      const key = this.hrFolderTaskCacheKey(name);
      if (this.fileCrawlCache.get(key)) return false;
      if (this.hrTaskPrefetchInFlight.has(key)) return false;
      return true;
    });
    if (pending.length === 0) return;

    let token: string;
    try {
      token = await this.getSharePointToken();
    } catch {
      return;
    }

    try {
      await this.ensureFormConfigLoaded();
      await this.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );
      if (!this.cachedSiteId) return;
    } catch {
      return;
    }

    const listsToQuery = this.getTaskListsToQuery(this.cachedSiteLists);
    if (listsToQuery.length === 0) return;

    const concurrency = AppConstants.hrFilesTaskPrefetchConcurrency;
    for (let i = 0; i < pending.length; i += concurrency) {
      if (this.folderPrefetchPaused) return;
      if (this.isLoadingComments || this.isLoadingMoreComments) return;

      const batch = pending.slice(i, i + concurrency);
      await Promise.all(
        batch.map(async folderName => {
          const key = this.hrFolderTaskCacheKey(folderName);
          if (this.fileCrawlCache.get(key) || this.hrTaskPrefetchInFlight.has(key)) return;

          // Snapshot active load seq - abort if the user starts a real folder task load.
          const loadSeq = this.allFilesFolderTaskLoadSeq;
          this.hrTaskPrefetchInFlight.add(key);
          let seedSnapshot: any[] | null = null;
          try {
            const mapped = await this.collectHrPersonalTasksFromLists(
              listsToQuery,
              folderName,
              token,
              loadSeq,
              AppConstants.hrFilesTaskListPageLimit,
              undefined,
              {
                onSeedComplete: (seedItems) => {
                  if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
                  if (this.folderPrefetchPaused) return;
                  // Cache seed early so a click during prefetch is already instant — memory
                  // only, and never over an existing snapshot (it may be a complete one).
                  if (this.fileCrawlCache.getStale(key) == null) {
                    seedSnapshot = this.sortCommentItemsByDateDesc(seedItems);
                    this.fileCrawlCache.setMemoryOnly(key, seedSnapshot);
                    this.hrFolderPartialKeys.add(key);
                  }
                },
              },
            );
            if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
            if (this.folderPrefetchPaused) return;
            // Skip write if a click already cached this folder while we were fetching —
            // but do replace our own partial seed with the full result.
            const current = this.fileCrawlCache.get(key);
            if (current && current !== seedSnapshot) return;
            // Never shrink an older snapshot: a silently failed list must not drop its rows.
            const prior = this.fileCrawlCache.getStale<any[]>(key);
            const full = Array.isArray(prior) && prior !== seedSnapshot
              ? this.mergeHrTasks(prior, mapped)
              : mapped;
            this.fileCrawlCache.set(key, this.sortCommentItemsByDateDesc(full));
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

  private schedulePrefetchChildFolders(
    driveId: string,
    items: Array<{ id?: string; isFolder?: boolean }>,
  ): void {
    // Disabled - child-folder children crawls compounded click/nav Graph traffic.
    return;
  }

  /** Silently warm the per-folder cache without touching Attachments UI. */
  private async prefetchDriveFolders(driveId: string, folderIds: string[]): Promise<void> {
    if (this.folderPrefetchPaused || this.isLoadingTodoTasks) return;
    const pending = folderIds.filter(folderId => {
      const key = this.driveFolderCacheKey(driveId, folderId);
      if (this.fileCrawlCache.get(key)) return false;
      if (this.folderPrefetchInFlight.has(key)) return false;
      return true;
    });
    if (pending.length === 0) return;

    let token: string;
    try {
      token = await this.getSharePointToken();
    } catch {
      return;
    }

    for (let i = 0; i < pending.length; i += this.folderPrefetchConcurrency) {
      if (this.folderPrefetchPaused || this.isLoadingTodoTasks) return;
      const batch = pending.slice(i, i + this.folderPrefetchConcurrency);
      await Promise.all(
        batch.map(async folderId => {
          const key = this.driveFolderCacheKey(driveId, folderId);
          if (this.fileCrawlCache.get(key) || this.folderPrefetchInFlight.has(key)) return;
          this.folderPrefetchInFlight.add(key);
          try {
            const mapped = await this.fetchDriveFolderChildrenMapped(driveId, folderId, token);
            this.fileCrawlCache.set(key, mapped);
          } catch {
            // Prefetch failures are silent - click path will retry.
          } finally {
            this.folderPrefetchInFlight.delete(key);
          }
        }),
      );
      // Pace batches so prefetch doesn't starve click / search traffic.
      if (i + this.folderPrefetchConcurrency < pending.length) {
        await new Promise(resolve => setTimeout(resolve, 400));
      }
    }
  }

  // ============================================================
  // LOAD LOGGED-IN USER FILES (main entry point)
  // ============================================================
  protected openMyFiles(): void {
    if (!this.currentUser) {
      alert('You must log in first to access your profile.');
      this.userFileError = 'Please log in first.';
      this.showUserFile = false;
      return;
    }
    if (this.isLoadingUserFiles) return;

    // Attachments-only entry: do not reload Comments/tasks.
    this.showUserFile = true;
    this.isLoadingUserFiles = true;
    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.folderStack = [];
    this.attachmentBrowsingRoot = null;
    this.currentLibraryDriveId = null;
    this.currentLibraryName = null;
    this.userRootFolderId = null;
    this.userFileProgressMessage = 'Connecting to SharePoint...';
    this.userFileError = '';
    this.userFileWarning = '';
    this.userFiles = [];
    this._folderMapDirty = true;
    this.startUserFileLoadingWatchdog();
    this.refreshView();

    void this.loadMyFilesFromDrive();
  }

  protected loadHrFilesInFirstSection(): void {
    if (!this.currentUser) {
      this.userFileError = 'Please log in first.';
      return;
    }
    if (this.isLoadingHrFilesList) return;

    this.cancelBackgroundTodoLoad();
    clearGraphThrottleCooldown();
    if (this.hrTaskPrefetchTimer) {
      clearTimeout(this.hrTaskPrefetchTimer);
      this.hrTaskPrefetchTimer = null;
    }

    this.userFileError = '';
    this.userFileWarning = '';

    // Instant paint from cache (memory / previous session) so the first open is not blank.
    const cachedRoots =
      this.fileCrawlCache.get<typeof this.hrPersonalRootFolders>(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY) ??
      this.fileCrawlCache.getStale<typeof this.hrPersonalRootFolders>(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY);
    if (Array.isArray(cachedRoots) && cachedRoots.length > 0) {
      this.hrPersonalRootFolders = cachedRoots;
      this.isLoadingHrFilesList = false;
      this.userFileProgressMessage = '';
      this.clearUserFileLoadingWatchdog();
      this.refreshView();
      // Top up from SharePoint without wiping the painted list.
      void this.loadHrPersonalFilesForFirstSection({ softRefresh: true });
      return;
    }

    this.isLoadingHrFilesList = true;
    this.userFileProgressMessage = 'Connecting to SharePoint...';
    this.hrPersonalRootFolders = [];
    this.startUserFileLoadingWatchdog();
    this.refreshView();

    void this.loadHrPersonalFilesForFirstSection();
  }

  private showMyFilesPanel(): void {
    this.showUserFile = true;
    this.refreshView();
  }

  // ============================================================
  // MANUAL REFRESH (refresh button) - reloads only the active section.
  // ============================================================
  protected onManualRefresh(): void {
    if (!this.currentUser) {
      this.showModal('Login Required', 'You must log in first to refresh.', 'error');
      return;
    }

    if (this.showAllFilesSection) {
      this.refreshAllFilesSection();
      return;
    }

    if (this.showToDoSection) {
      this.refreshToDoSection();
      return;
    }

    this.refreshHrSection();
  }

  private refreshAllFilesSection(): void {
    if (this.selectedFolderName && this.currentLibraryName) {
      this.invalidateFolderTaskCache(this.currentLibraryName, this.selectedFolderName);
    }
    this.fileCrawlCache.invalidatePrefix('files:folder:');
    void this.allFilesComponent?.refreshAllFilesData(true);

    if (this.selectedFolderName && this.currentLibraryName) {
      this.loadAllFilesTasksForLibrary(this.currentLibraryName, this.selectedFolderName);
    }

    const stackFolder = this.folderStack[this.folderStack.length - 1];
    const folderId =
      this.currentFolderId ||
      this.attachmentBrowsingRoot?.id ||
      stackFolder?.id ||
      '';
    const folderWebUrl =
      this.currentFolderWebUrl ||
      this.attachmentBrowsingRoot?.webUrl ||
      stackFolder?.webUrl ||
      '';
    const folderName =
      this.currentFolderName ||
      this.attachmentBrowsingRoot?.name ||
      stackFolder?.name ||
      '';
    if (this.currentLibraryDriveId && folderId) {
      void this.loadDriveFolderContents(
        this.currentLibraryDriveId,
        folderId,
        folderWebUrl,
        folderName,
        { skipCache: true },
      );
    }

    this.refreshView();
  }

  private refreshToDoSection(): void {
    // Subordinate view has its own load path and no per-list watermarks.
    if (this.superiorMode) {
      this.fileCrawlCache.invalidate(AppComponent.HR_USER_TASKS_CACHE_KEY);
      void this.reloadTodoTasksWithSubordinates();
      this.refreshView();
      return;
    }

    // Nothing cached yet - nothing to top up, so do the normal first load.
    if (this.todoListWatermarks.size === 0) {
      void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
      this.refreshView();
      return;
    }

    void this.refreshTodoTasksIncrementally();

    if (this.selectedSubmitter) {
      this.loadHrTasks();
    }

    this.refreshView();
  }

  /**
   * Top up To Do with whatever changed since the last look, keeping everything already
   * collected. Each list is read newest-first and stops at the first row already held,
   * so no previously fetched row is downloaded, re-mapped or re-filtered a second time.
   */
  private async refreshTodoTasksIncrementally(): Promise<boolean> {
    if (this.isLoadingMoreTodoTasks || this.isLoadingTodoTasks) return false;

    this.isLoadingMoreTodoTasks = true;
    this.refreshView();

    try {
      const token = await this.getSharePointToken();
      const siteId = this.cachedSiteId;
      const siteListsArr = this.cachedSiteLists;
      if (!siteId || siteListsArr.length === 0) {
        void this.loadTodoTasksForCurrentUser({ forceRefresh: true });
        return false;
      }

      const siteWebUrl = this.cachedSiteWebUrl;
      const userEmail = (this.currentUser?.email ?? '').toLowerCase();
      const userUpn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();
      const watermarks = [...this.todoListWatermarks.entries()];

      const refreshOne = async ([listName, watermark]: [string, string]): Promise<any[]> => {
        const list = siteListsArr.find(
          (l: any) =>
            (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
            (l.displayName ?? '').toLowerCase() === listName.toLowerCase(),
        );
        if (!list?.id) return [];

        const raw = await this.fetchTodoRowsNewerThanWatermark(
          siteId, list.id, token, listName, watermark,
        );
        if (raw.length === 0) return [];

        this.noteTodoWatermark(listName, raw);

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
      const existing = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
      const merged = fresh.length > 0 ? this.mergeHrTasks(existing, fresh) : existing;
      if (fresh.length > 0) {
        this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, merged);
      }
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(merged));
      this.userHrTasksLoaded = true;
      return fresh.length > 0;
    } catch {
      // Refresh is best effort - the existing view stays as-is.
      return false;
    } finally {
      this.isLoadingMoreTodoTasks = false;
      this.refreshView();
    }
  }

  /**
   * While signed in, periodically top up To Do Redis from SharePoint watermarks and
   * soft-refresh the open folder Comments cache so external SharePoint edits appear.
   */
  private startSharePointCachePoll(): void {
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', this.onDocumentVisibilityForCachePoll);
      document.addEventListener('visibilitychange', this.onDocumentVisibilityForCachePoll);
    }
    if (typeof document !== 'undefined' && document.visibilityState === 'hidden') {
      return;
    }
    this.stopSharePointCachePollTimerOnly();
    this.sharePointCachePollTimer = setInterval(() => {
      void this.pollSharePointTaskCaches();
    }, AppComponent.SHAREPOINT_CACHE_POLL_MS);
  }

  private stopSharePointCachePollTimerOnly(): void {
    if (this.sharePointCachePollTimer) {
      clearInterval(this.sharePointCachePollTimer);
      this.sharePointCachePollTimer = null;
    }
  }

  private stopSharePointCachePoll(): void {
    this.stopSharePointCachePollTimerOnly();
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', this.onDocumentVisibilityForCachePoll);
    }
  }

  private async pollSharePointTaskCaches(): Promise<void> {
    if (!this.currentUser?.email) return;
    if (typeof document !== 'undefined' && document.visibilityState === 'hidden') return;
    if (this.superiorMode) return;
    // Do not compete with an active To Do / Comments Graph fan-out.
    if (
      this.isLoadingTodoTasks ||
      this.todoProcurementDrainActive ||
      this.isLoadingComments ||
      this.isLoadingMoreComments ||
      this.folderTaskGraphInFlight
    ) {
      return;
    }

    if (this.todoListWatermarks.size > 0) {
      await this.refreshTodoTasksIncrementally();
    }

    const now = Date.now();
    if (now - this.lastFolderCommentsPollMs < AppComponent.FOLDER_COMMENTS_POLL_MS) {
      return;
    }
    this.lastFolderCommentsPollMs = now;
    this.softRefreshOpenFolderTaskCache();
  }

  /** Soft-refresh the open folder Comments cache from Graph without wiping the painted UI. */
  private softRefreshOpenFolderTaskCache(): void {
    const folder = String(this.selectedFolderName || this.selectedHrPersonalTaskFolder || '').trim();
    if (!folder) return;

    const loadSeq = ++this.allFilesFolderTaskLoadSeq;

    if (this.showAllFilesSection && this.currentLibraryName) {
      if (normalizeName(this.currentLibraryName) === normalizeName(this.targetLibraryName)) {
        void this.loadHrPersonalTasksForFolder(folder, loadSeq, { softRefresh: true });
        return;
      }
      const mappedTaskList = this.getTaskListForLibrary(this.currentLibraryName);
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

  private refreshHrSection(): void {
    if (this.selectedFolderName) {
      this.invalidateFolderTaskCache(this.targetLibraryName, this.selectedFolderName);
    }

    this.fileCrawlCache.invalidate(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY);
    this.hrPersonalRootFolders = [];
    this.loadHrFilesInFirstSection();

    const stackFolder = this.folderStack[this.folderStack.length - 1];
    const folderId =
      this.currentFolderId ||
      this.attachmentBrowsingRoot?.id ||
      stackFolder?.id ||
      '';
    const folderWebUrl =
      this.currentFolderWebUrl ||
      this.attachmentBrowsingRoot?.webUrl ||
      stackFolder?.webUrl ||
      '';
    if (this.hrPersonalDriveId && this.currentLibraryDriveId === this.hrPersonalDriveId && folderId) {
      this.fileCrawlCache.invalidate(this.driveFolderCacheKey(this.hrPersonalDriveId, folderId));
      if (this.selectedFolderName) {
        this.loadAllFilesTasksForLibrary(this.targetLibraryName, this.selectedFolderName);
      }
      void this.loadDriveFolderContents(
        this.hrPersonalDriveId,
        folderId,
        folderWebUrl,
        this.currentFolderName || this.attachmentBrowsingRoot?.name || stackFolder?.name,
        { skipCache: true },
      );
    } else if (
      this.showUserFile &&
      this.userRootFolderId &&
      !this.currentLibraryDriveId &&
      !this.isLoadingUserFiles
    ) {
      this.isLoadingUserFiles = true;
      this.userFileProgressMessage = 'Refreshing attachments...';
      this.userFileError = '';
      this.userFileWarning = '';
      this.startUserFileLoadingWatchdog();
      this.refreshView();
      void this.loadMyFilesFromDrive();
    }

    this.refreshView();
  }

  protected get isManualRefreshLoading(): boolean {
    return (
      this.isLoadingComments ||
      this.isLoadingUserFiles ||
      this.isLoadingTodoTasks ||
      this.isLoadingHrFilesList ||
      !!this.allFilesComponent?.isLoading ||
      !!this.allFilesComponent?.isLoadingAllLibraries ||
      !!this.allFilesComponent?.isLoadingCommentTasks ||
      !!this.allFilesComponent?.isSearchingMoreAttachments
    );
  }

  private async loadMyFilesFromDrive(): Promise<void> {
    try {
      const token = await this.getSharePointToken();
      const targetDriveId = await this.getTargetDriveId(token);

      if (!targetDriveId) {
        const site = await this.siteMetadataService.resolve(token);
        const drives: any = await graphGetWithRetry(
          this.http,
          `/sites/${site.siteId}/drives?$select=id,name,webUrl`,
          token,
        );
        const available = drives?.value?.map((x: any) => x.name).join(', ') || 'none';
        this.userFileError = `Could not find "${this.targetLibraryName}". Available libraries: ${available}`;
        this.refreshView();
        return;
      }

      await this.loadCurrentUserPersonalFolderFiles(targetDriveId, token);
    } catch (error: unknown) {
      this.handleUserFilesError(error);
    } finally {
      if (this.isLoadingUserFiles) {
        this.isLoadingUserFiles = false;
        this.clearUserFileLoadingWatchdog();
        this.refreshView();
      }
    }
  }

  private async loadHrPersonalFilesForFirstSection(
    options: { softRefresh?: boolean } = {},
  ): Promise<void> {
    const softRefresh = !!options.softRefresh;
    try {
      const token = await this.getSharePointToken();
      if (!softRefresh) {
        this.userFileProgressMessage = 'Loading HR Files...';
        this.refreshView();
      }

      // Drive resolve (Graph) and group membership (SharePoint REST) run in parallel -
      // the old serial path waited for both before the first children page could start.
      const [targetDriveId, allAccess] = await Promise.all([
        this.getTargetDriveId(token),
        this.isUserInHrPersonalAllAccessGroup(),
      ]);

      if (!targetDriveId) {
        if (!softRefresh) {
          this.userFileError = `Could not find "${this.targetLibraryName}".`;
          this.refreshView();
        }
        return;
      }
      this.hrPersonalDriveId = targetDriveId;

      // All-access: list root immediately. Skip the separate canList probe - it cost an
      // extra Graph RTT and the real children fetch already proves access (403 -> personal).
      if (allAccess) {
        try {
          this.hrPersonalRootFolders = await this.fetchHrPersonalRootFolders(
            targetDriveId,
            token,
            (partial: Array<{ id: string; name: string; webUrl: string; isFolder: boolean }>) => {
              this.hrPersonalRootFolders = partial;
              this.fileCrawlCache.set(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY, partial);
              this.isLoadingHrFilesList = false;
              this.userFileProgressMessage = '';
              this.clearUserFileLoadingWatchdog();
              this.refreshView();
              this.scheduleHrFolderTasksPrefetch(partial);
            },
          );
          this.fileCrawlCache.set(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY, this.hrPersonalRootFolders);
          this.scheduleHrFolderTasksPrefetch(this.hrPersonalRootFolders);
          return;
        } catch (err: any) {
          const status = err?.status ?? err?.error?.status;
          if (status !== 403 && status !== 401) throw err;
          // Not actually allowed to list root - fall through to personal folder.
        }
      }

      const userFolder = await this.findUserFolder(targetDriveId, token);
      if (!userFolder?.id) {
        if (!softRefresh) {
          this.userFileError =
            `Could not find a personal folder for ${this.currentUser?.username || 'the logged-in user'} in "${this.targetLibraryName}".`;
          this.refreshView();
        }
        return;
      }

      this.hrPersonalFolderName = userFolder.name;
      this.hrPersonalRootFolders = [{ ...userFolder, webUrl: userFolder.webUrl ?? '', isFolder: true }];
      this.fileCrawlCache.set(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY, this.hrPersonalRootFolders);
      this.scheduleHrFolderTasksPrefetch(this.hrPersonalRootFolders);
    } catch (error: unknown) {
      if (!softRefresh) {
        this.handleUserFilesError(error);
      }
    } finally {
      if (this.isLoadingHrFilesList) {
        this.isLoadingHrFilesList = false;
        this.clearUserFileLoadingWatchdog();
        this.refreshView();
      }
    }
  }

  // Intentionally removed: hardcoded allow-list for HRPersonal root browsing.(updated)

  private hrPersonalGroupChecked = false;
  private hrPersonalGroupAllowsAllAccess = false;
  private readonly hrPersonalAllAccessGroupId = 3;

  private async isUserInHrPersonalAllAccessGroup(): Promise<boolean> {
    if (this.hrPersonalGroupChecked) return this.hrPersonalGroupAllowsAllAccess;
    this.hrPersonalGroupChecked = true;
    this.hrPersonalGroupAllowsAllAccess = false;

    const email = (this.currentUser?.email ?? '').trim().toLowerCase();
    const upn = (this.currentUser?.userPrincipalName ?? '').trim().toLowerCase();
    if (!email && !upn) return false;

    try {
      const token = await this.authService.acquireSharePointRestWriteToken();
      const url =
        `https://${sharePointConfig.siteHostName}/${sharePointConfig.sitePath}` +
        `/_api/web/sitegroups/getbyid(${this.hrPersonalAllAccessGroupId})/users?$select=Email,UserPrincipalName,LoginName,Title`;

      const resp: any = await firstValueFrom(
        this.http.get(url, {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json;odata=nometadata',
          },
        })
      );

      const users: any[] = resp?.value ?? resp?.d?.results ?? [];
      this.hrPersonalGroupAllowsAllAccess = users.some(u => {
        const uEmail = String(u?.Email ?? '').trim().toLowerCase();
        const uUpn = String(u?.UserPrincipalName ?? '').trim().toLowerCase();
        const login = String(u?.LoginName ?? '').trim().toLowerCase();
        return (
          (email && (uEmail === email || uUpn === email)) ||
          (upn && (uEmail === upn || uUpn === upn)) ||
          (email && login.includes(email)) ||
          (upn && login.includes(upn))
        );
      });
    } catch {
      this.hrPersonalGroupAllowsAllAccess = false;
    }

    return this.hrPersonalGroupAllowsAllAccess;
  }


  /** Background-only: resolve HRPersonal root folders into cache without touching the HR Files UI. */
  private async warmHrPersonalRootFoldersCache(): Promise<void> {
    if (!this.currentUser) return;
    if (this.isLoadingHrFilesList) return;

    try {
      if (!this.fileCrawlCache.get(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY)) {
        const token = await this.getSharePointToken();
        const targetDriveId = await this.getTargetDriveId(token);
        if (!targetDriveId) return;
        this.hrPersonalDriveId = targetDriveId;

        if (!(await this.isUserInHrPersonalAllAccessGroup())) {
          const userFolder = await this.findUserFolder(targetDriveId, token);
          if (!userFolder?.id) return;
          this.hrPersonalFolderName = userFolder.name;
          const single = [{ ...userFolder, webUrl: userFolder.webUrl ?? '', isFolder: true as const }];
          this.fileCrawlCache.set(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY, single);
        } else {
          const folders = await this.fetchHrPersonalRootFolders(targetDriveId, token, (partial) => {
            this.fileCrawlCache.set(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY, partial);
          });
          this.fileCrawlCache.set(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY, folders);
        }
      }
    } catch {
      // Warm is best-effort - first HR Files open will fetch normally.
    }

    // Light Comments warm for the first few people so early clicks hit shared Redis.
    this.scheduleLoginHrTaskWarm();
  }

  /**
   * After people-list warm (or Redis hydrate), prefetch Comments for a small batch
   * of HR folders into shared Redis. Skips when Graph is paused for clicks / To Do.
   */
  private scheduleLoginHrTaskWarm(): void {
    const roots =
      this.fileCrawlCache.get<typeof this.hrPersonalRootFolders>(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY) ??
      this.fileCrawlCache.getStale<typeof this.hrPersonalRootFolders>(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY);
    if (!Array.isArray(roots) || roots.length === 0) return;

    const names = roots
      .map(f => String(f?.name ?? '').trim())
      .filter(Boolean)
      .slice(0, AppConstants.hrFilesTaskPrefetchMaxBatch);
    if (names.length === 0) return;

    if (this.hrTaskPrefetchTimer) {
      clearTimeout(this.hrTaskPrefetchTimer);
    }
    this.hrTaskPrefetchTimer = setTimeout(() => {
      this.hrTaskPrefetchTimer = null;
      if (this.folderPrefetchPaused) {
        // To Do still owns Graph — retry once after the usual pause window.
        this.hrTaskPrefetchTimer = setTimeout(() => {
          this.hrTaskPrefetchTimer = null;
          if (this.folderPrefetchPaused) return;
          if (this.isLoadingComments || this.isLoadingMoreComments) return;
          void this.prefetchHrFolderTasks(names);
        }, 8000);
        return;
      }
      if (this.isLoadingComments || this.isLoadingMoreComments) return;
      void this.prefetchHrFolderTasks(names);
    }, AppConstants.hrFilesTaskPrefetchDelayMs);

    void this.warmAllHrPeopleSlowly();
  }

  /**
   * HR all-access users: trickle-warm Comments for every person into shared Redis,
   * one at a time, so first opens are almost always cache hits for everyone.
   * Yields to any user-driven load and skips people cached within the TTL.
   */
  private async warmAllHrPeopleSlowly(): Promise<void> {
    if (this.hrWarmAllRunning) return;
    this.hrWarmAllRunning = true;
    try {
      if (!(await this.isUserInHrPersonalAllAccessGroup())) return;
      // Let the first-batch prefetch and To Do get Graph first.
      await this.sleep(AppConstants.hrFilesWarmAllStartDelayMs);

      const roots =
        this.fileCrawlCache.getStale<typeof this.hrPersonalRootFolders>(AppComponent.HR_ROOT_FOLDERS_CACHE_KEY) ?? [];
      // Shuffle so several HR browsers spread across different people.
      const names = roots
        .map(f => String(f?.name ?? '').trim())
        .filter(Boolean)
        .sort(() => Math.random() - 0.5);

      for (const name of names) {
        if (!this.currentUser) return;
        if (this.fileCrawlCache.get(this.hrFolderTaskCacheKey(name))) continue;
        while (
          this.folderPrefetchPaused ||
          this.isLoadingComments ||
          this.isLoadingMoreComments ||
          this.isLoadingTodoTasks ||
          this.isLoadingHrFilesList
        ) {
          await this.sleep(5000);
          if (!this.currentUser) return;
        }
        await this.prefetchHrFolderTasks([name]);
        await this.sleep(AppConstants.hrFilesWarmAllGapMs);
      }
    } catch {
      // Best-effort warm — clicks still load normally.
    } finally {
      this.hrWarmAllRunning = false;
    }
  }

  private async fetchHrPersonalRootFolders(
    driveId: string,
    token: string,
    onPage?: (
      folders: Array<{ id: string; name: string; webUrl: string; isFolder: boolean }>
    ) => void,
  ): Promise<Array<{ id: string; name: string; webUrl: string; isFolder: boolean }>> {
    const folders: Array<{ id: string; name: string; webUrl: string; isFolder: boolean }> = [];
    // Smaller pages so the first HR Files paint does not wait for the full library root.
    let nextPath: string | null =
      `/drives/${driveId}/root/children?$select=id,name,webUrl,folder&$top=50`;

    while (nextPath) {
      const page: any = await graphGetWithRetry(
        this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs,
      );
      folders.push(
        ...(page?.value ?? [])
          .filter((item: any) => !!item.folder)
          .map((item: any) => ({
            id: item.id,
            name: item.name,
            webUrl: item.webUrl ?? '',
            isFolder: true,
          }))
      );
      onPage?.(folders.slice());
      nextPath = toGraphPath(page?.['@odata.nextLink']);
    }

    return folders;
  }

  private async loadCurrentUserPersonalFolderFiles(targetDriveId: string, token: string): Promise<void> {
    this.userFileProgressMessage = 'Locating your Personal folder/Files...';
    this.refreshView();

    const userFolder = await this.findUserFolder(targetDriveId, token);
    if (!userFolder?.id) {
      this.userFileError =
        `Could not find a personal folder for ${this.currentUser?.username || 'the logged-in user'} in "${this.targetLibraryName}".`;
      this.refreshView();
      return;
    }

    this.userRootFolderId = userFolder.id;
    this.hrPersonalFolderName = userFolder.name;

    const driveLoadResult = await this.getAllDriveItems(targetDriveId, token, userFolder.id, progress => {
      this.userFileProgressMessage = progress;
      this.refreshView();
    });

    this.applyDriveItems(driveLoadResult.items, userFolder.id);

    if (driveLoadResult.wasTruncated) {
      this.userFileWarning = 'SharePoint scan limit reached; showing partial results.';
    }
  }
  //Folder Name LOOKUP(runs at login,lightweight lookup)
  private async loadHrPersonalFolderName(): Promise<void> {
    if (!this.currentUser) return;
    this.refreshView();
    try {
      const token = await this.getSharePointToken();
      const targetDriveId = await this.getTargetDriveId(token);
      if (!targetDriveId) return;
      const userFolder = await this.findUserFolder(targetDriveId, token);
      if (!userFolder?.id) return;
      this.hrPersonalFolderName = userFolder.name;
    } catch {
      //silent ignore -- card will show fallback text
    } finally {
      this.refreshView();
    }
  }

  private applyDriveItems(items: any[], rootFolderId: string): void {
    this.userFiles = items
      .filter((item: any) => !!item.file || !!item.folder)
      .map((item: any) => {
        const ext = item.file ? getFileExtension(item.name) : '';
        const category = item.file ? getFileCategory(ext) : '';
        const icon = item.file ? getFileIcon(category, ext) : '';
        const parentPath = item.parentReference?.path ?? '';
        const isRootChild = rootFolderId === 'root' && /\/root:?$/i.test(parentPath);
        return {
          id: item.id,
          name: item.name,
          webUrl: item.webUrl ?? '',
          parentId: isRootChild ? rootFolderId : item.parentReference?.id ?? rootFolderId,
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

    this._folderMapDirty = true;
  }

  // ============================================================
  // FOLDER NAVIGATION
  // ============================================================
  protected openFolder(item: { id: string; name: string; webUrl: string }): void {
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
        this.loadAllFilesTasksForLibrary(this.currentLibraryName!, String(item.name || ''));
      }
      return;
    }

    // Personal files: full tree already loaded - filter locally by parentId
    this.pushCurrentFolderOntoStack();
    this.currentFolderId = item.id;
    this.currentFolderWebUrl = item.webUrl;
    this.currentFolderName = item.name;
    this.refreshView();
  }

  /** True when drilling into an Attachments subfolder should also refresh Comments tasks. */
  private shouldReloadTasksForAttachmentFolder(folderName: string): boolean {
    const libName = String(this.currentLibraryName ?? '').trim();
    if (!libName || !String(folderName ?? '').trim()) return false;
    // HRPersonal person tree: Comments already loaded for the root person.
    if (normalizeName(libName) === normalizeName(this.targetLibraryName)) {
      return false;
    }
    // All Files document libraries: folder name is the task Title key.
    return this.showAllFilesSection && !!this.getTaskListForLibrary(libName);
  }

  private pushCurrentFolderOntoStack(): void {
    if (!this.currentFolderId) return;
    this.folderStack.push({
      id: this.currentFolderId,
      name: this.currentFolderName || this.selectedFolderName || '',
      webUrl: this.currentFolderWebUrl,
    });
  }

  protected openHrFileListItem(item: { id: string; name: string; webUrl: string; isFolder: boolean }): void {
    if (item.isFolder) {
      void this.openHrFileListItemAsync(item);
      return;
    }

    if (item.webUrl) {
      window.open(item.webUrl, '_blank', 'noopener,noreferrer');
    }
  }

  private async openHrFileListItemAsync(item: { id: string; name: string; webUrl: string; isFolder: boolean }): Promise<void> {
    if (!(await this.ensureHrPersonalDriveId())) {
      this.userFileError = 'HRPersonal library is not loaded yet. Wait for HR Files to finish loading, then try again.';
      this.refreshView();
      return;
    }

    // Free Graph capacity for this person's click-path load.
    this.cancelBackgroundTodoLoad();
    clearGraphThrottleCooldown();
    this.pauseFolderPrefetch();

    this.showUserFile = true;
    this.hideComments = false;
    this.showToDoSection = false;
    this.showAllFilesSection = false;
    this.currentLibraryDriveId = this.hrPersonalDriveId;
    this.currentLibraryName = this.targetLibraryName;
    this.attachmentBrowsingRoot = { id: item.id, name: item.name, webUrl: item.webUrl };
    this.currentFolderId = item.id;
    this.currentFolderWebUrl = item.webUrl;
    this.currentFolderName = item.name;
    this.folderStack = [];
    this.userRootFolderId = null;
    this.selectedFolderName = item.name;
    this.selectedHrPersonalTaskFolder = item.name;
    this.lastHrFilesCommentsFolder = item.name;

    // Same-session instant return: paint memory/stale BEFORE any async work so
    // navigating away and back never flashes the loading state for cached people.
    const driveId = this.hrPersonalDriveId!;
    const filesKey = this.driveFolderCacheKey(driveId, item.id);
    const cachedFiles =
      this.fileCrawlCache.get<any[]>(filesKey) ??
      this.fileCrawlCache.getStale<any[]>(filesKey);
    if (Array.isArray(cachedFiles)) {
      this.applyDriveFolderContentsToUi(cachedFiles, item.id, item.webUrl, item.name);
    } else {
      this.userFiles = [];
      this._folderMapDirty = true;
      this.isLoadingUserFiles = true;
      this.userFileProgressMessage = 'Loading folder contents...';
    }

    const tasksKey = this.hrFolderTaskCacheKey(item.name);
    const cachedTasks =
      this.fileCrawlCache.get<any[]>(tasksKey) ??
      this.fileCrawlCache.getStale<any[]>(tasksKey);
    // Allow empty arrays — a completed crawl with 0 tasks is still a valid hit.
    if (Array.isArray(cachedTasks)) {
      this.commentItems = this.sortCommentItemsByDateDesc(
        this.applyAllFilesFolderSubmitter(cachedTasks),
      );
      this.commentsMessage = cachedTasks.length ? '' : 'No tasks found.';
      this.isLoadingComments = false;
      // Soft-refresh / deep heal sets Loading more — don't imply the cache is final.
      this.isLoadingMoreComments = false;
    } else {
      this.commentItems = [];
      this.commentsMessage = '';
    }
    this.invalidateCommentFilters();
    // Attachments first — tasks continue in the background for Comments.
    Object.assign(this, getAttachmentsTabMobileState());
    this.refreshView();
    void this.loadDriveFolderContents(driveId, item.id, item.webUrl, item.name);
    this.loadAllFilesTasksForLibrary(this.targetLibraryName, item.name);
  }

  protected goBackToParent(): void {
    this.navigateToParentFolder();
  }

  private navigateToParentFolder(): void {
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
          this.loadAllFilesTasksForLibrary(this.currentLibraryName!, String(parent.name || ''));
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
      this.refreshView();
      return;
    }

    if (this.attachmentBrowsingRoot) {
      this.currentFolderId = this.attachmentBrowsingRoot.id;
      this.currentFolderWebUrl = this.attachmentBrowsingRoot.webUrl;
      this.currentFolderName = this.attachmentBrowsingRoot.name;
      this.refreshView();
      return;
    }

    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.refreshView();
  }

  /** At the entry folder, close attachments; otherwise reload that folder's contents. */
  private resetToAttachmentBrowsingRoot(): void {
    if (!this.attachmentBrowsingRoot || !this.currentLibraryDriveId) {
      this.currentFolderId = null;
      this.currentFolderWebUrl = '';
      this.currentFolderName = '';
      this.refreshView();
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
      this.loadAllFilesTasksForLibrary(
        this.currentLibraryName!,
        String(this.attachmentBrowsingRoot.name || ''),
      );
    }
  }

  // ============================================================
  // HR TASK LIST SELECTION
  // ============================================================
  protected onHrTaskListSelected(listName: string): void {
    if (listName && this.allowedHrTaskLists.includes(listName)) {
      this.selectedHrTaskList = listName;
      // Show the file panel when a list is selected
      this.showUserFile = true;
      this.refreshView();
    }
  }

  protected loadHrTaskListItems(): void {
    if (!this.currentUser || !this.selectedHrTaskList) {
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
    this.refreshView();

    void this.loadHrTaskListItemsInternal();
  }

  private async loadHrTaskListItemsInternal(): Promise<void> {
    try {
      const token = await this.getSharePointToken();
      const siteHost = sharePointConfig.siteHostName;
      const sitePath = sharePointConfig.sitePath;
      const listName = this.selectedHrTaskList;

      if (!listName || !this.allowedHrTaskLists.includes(listName)) {
        this.userFileError = 'Invalid HR task list selected.';
        this.refreshView();
        return;
      }

      const site = await this.siteMetadataService.resolve(token);
      if (!site?.siteId) throw new Error('Could not find site');

      this.userFileProgressMessage = `Finding ${listName} list...`;
      this.refreshView();

      // Match by name (more reliable than $filter) against the shared list metadata.
      const allLists = site.lists;
      const targetList = allLists.find((l: any) =>
        (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
        (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
      );

      if (!targetList?.id) {
        this.userFileError = `Could not find list "${listName}". Available lists: ${allLists.map((l: any) => l.name).join(', ')}`;
        this.refreshView();
        return;
      }

      this.userFileProgressMessage = `Loading items from ${listName}...`;
      this.refreshView();

      // Load items with pagination (hard cap - never drain a multi-thousand-item list here)
      const allItems: any[] = [];
      let nextUrl: string | null = `/sites/${site.siteId}/lists/${targetList.id}/items?$expand=fields&$top=200`;
      let pages = 0;
      const maxPages = 10;

      while (nextUrl && pages < maxPages) {
        this.userFileProgressMessage = `Loading items from ${listName}... (${allItems.length} loaded)`;
        this.refreshView();

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
      this.refreshView();
    } catch (error: unknown) {
      this.handleUserFilesError(error);
    } finally {
      if (this.isLoadingUserFiles) {
        this.isLoadingUserFiles = false;
        this.clearUserFileLoadingWatchdog();
        this.refreshView();
      }
    }
  }

  /** Pull assignee-matched procurement tasks already loaded via All Files folder views. */
  private collectAssignedProcTasksFromFolderCaches(): any[] {
    const matchName = this.currentUser?.username ?? this.currentUser?.userPrincipalName ?? '';
    const matchEmail = (this.currentUser?.email ?? this.currentUser?.userPrincipalName ?? '').toLowerCase();
    if (!matchName && !matchEmail) return [];

    const byKey = new Map<string, any>();
    for (const entry of this.fileCrawlCache.getEntriesByPrefix('tasks:lib:v7:')) {
      if (!/proctasks/i.test(entry.key)) continue;
      const items = entry.data;
      if (!Array.isArray(items)) continue;

      for (const item of items) {
        if (!isTodoProcurementTaskList(String(item?.listName ?? ''))) continue;
        if (!this.isMappedTaskAssignedToLoggedInUser(item, matchName, matchEmail)) continue;
        byKey.set(this.getHrTaskKey(item), {
          ...item,
          isAssignedToCurrentUser: true,
        });
      }
    }
    return [...byKey.values()];
  }

  private isMappedTaskAssignedToLoggedInUser(
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

  /** Task lists from eForms config, hard-coded names, and HRTask* auto-detect. */
  private getTaskListsToQuery(
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

  /**
   * Query high-traffic HR lists first so Comments gets a newest-first hit sooner
   * (avoids a long blank "Loading tasks..." while empty lists run ahead of Sick/Missing Punch).
   */
  private prioritizeHrPersonTaskLists(listNames: string[]): string[] {
    const rank = (name: string): number => {
      const key = String(name ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (key.includes('sickcertificate')) return 0;
      if (key.includes('missingpunch')) return 1;
      if (key.includes('changeofshift')) return 2;
      if (key.includes('telework')) return 3;
      if (key.includes('rest')) return 4;
      if (key.startsWith('hrtask')) return 5;
      if (
        key.includes('probation') ||
        key.includes('increment') ||
        key.includes('performancereview') ||
        key.includes('preformancereview')
      ) {
        return 6;
      }
      return 7;
    };
    return [...listNames].sort((a, b) => {
      const diff = rank(a) - rank(b);
      return diff !== 0 ? diff : a.localeCompare(b);
    });
  }

  /** Add archive companion lists (e.g. ProcTasksArchive for procurement). */
  private expandTaskListsWithArchiveCompanions(listNames: string[]): string[] {
    const expanded = new Set(listNames);
    for (const name of listNames) {
      for (const archive of this.getArchiveCompanionTaskLists(name)) {
        expanded.add(archive);
      }
    }
    return [...expanded];
  }

  private async ensureFormConfigLoaded(): Promise<void> {
    if (this.formConfigService.snapshot.length > 0) {
      return;
    }
    await firstValueFrom(this.formConfigService.load());
  }

  // ============================================================
  // HR TASK HELPERS - merge, older backfill
  // ============================================================
  private getHrTaskKey(item: any): string {
    return `${item.listName ?? 'unknown'}:${item.id}`;
  }

  /** User-created comment cards (not SharePoint HR tasks) - must not appear in the todo list. */


  /** Matches exact labels plus SharePoint variants like "New Attachment/s". */

  /** Hide status badge for general comments, attachments, and request-for-action cards. */
  protected shouldHideCommentStatus(item: {
    id?: unknown;
    name?: unknown;
    eFormDetails?: Record<string, unknown>;
  }): boolean {
    return shouldHideCommentStatus(item);
  }

  /** Returns which comment-entry type an item is, or null for normal eForm tasks. */


  /** Comment rows and user-added comment cards linked by eFormListId. */

  private isHrPersonalFilesContext(): boolean {
    return !!String(this.selectedHrPersonalTaskFolder ?? '').trim();
  }

  /** True when Comments is scoped to a folder (HR Files person or All Files folder). */
  private isFolderTaskCommentsViewActive(): boolean {
    if (this.showAllFilesSection && !!String(this.selectedFolderName ?? '').trim()) {
      return true;
    }
    return this.isHrPersonalFilesContext();
  }

  /** Cancel in-flight folder/global task fetches so a new folder click wins. */
  private cancelPendingFolderTaskLoads(): void {
    this.allFilesFolderTaskLoadSeq++;
    this.hrCommentsLoadSeq++;
  }

  /** Leave HR Files tab context - drop All Files navigation without losing hrPersonalDriveId. */
  private leaveAllFilesTaskView(): void {
    if (
      !this.hrPersonalDriveId &&
      this.currentLibraryDriveId &&
      normalizeName(this.currentLibraryName ?? '') === normalizeName(this.targetLibraryName)
    ) {
      this.hrPersonalDriveId = this.currentLibraryDriveId;
    }

    this.cancelPendingFolderTaskLoads();
    this.currentLibraryDriveId = null;
    this.currentLibraryName = null;
    this.currentFolderId = null;
    this.currentFolderWebUrl = '';
    this.currentFolderName = '';
    this.folderStack = [];
    this.attachmentBrowsingRoot = null;
    this.selectedFolderName = '';
    this.selectedFolderModifiedBy = '';
    this.selectedHrPersonalTaskFolder = '';
    this.showUserFile = false;
    this.userFiles = [];
    this._folderMapDirty = true;
    this.commentItems = [];
    this.commentsMessage = '';
    this.isLoadingComments = false;
    this.isLoadingMoreComments = false;
    this.invalidateCommentFilters();
  }

  private async ensureHrPersonalDriveId(): Promise<boolean> {
    if (this.hrPersonalDriveId) return true;

    if (
      this.currentLibraryDriveId &&
      normalizeName(this.currentLibraryName ?? '') === normalizeName(this.targetLibraryName)
    ) {
      this.hrPersonalDriveId = this.currentLibraryDriveId;
      return true;
    }

    try {
      const token = await this.getSharePointToken();
      const driveId = await this.getTargetDriveId(token);
      if (!driveId) return false;
      this.hrPersonalDriveId = driveId;
      return true;
    } catch {
      return false;
    }
  }

  /** HR Files shows only submitted eForm tasks - not comments, attachments, or action requests. */
  private isHrFilesVisibleItem(
    item: { id?: unknown; name?: unknown; submittedBy?: string; eFormDetails?: Record<string, unknown> },
    folderName: string,
  ): boolean {
    if (isCommentTypeItem(item)) {
      return false;
    }
    return isItemSubmittedByFolderPerson(item, folderName);
  }


  /** True when the mapped HR task was submitted by the person represented by an HR Files folder. */


  private getHrTasksForTodoList<T extends { id?: unknown }>(items: T[]): T[] {
    return items.filter(item => !isSyntheticCommentCard(item));
  }

  /**
   * Remove a completed/delegated task from the To Do Redis snapshot (`tasks:hr-user`)
   * so the next instant paint does not bring it back.
   */
  private removeTaskFromTodoCache(taskId: string): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;

    const cached = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY);
    if (!Array.isArray(cached) || cached.length === 0) return;

    const next = cached.filter((item: any) => String(item?.id ?? '').trim() !== id);
    if (next.length === cached.length) return;

    this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, next);
  }

  /** Merge field updates into the To Do Redis snapshot for one task id. */
  private patchTaskInTodoCache(taskId: string, patch: (item: any) => any): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;

    const cached = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY);
    if (!Array.isArray(cached) || cached.length === 0) return;

    let changed = false;
    const next = cached.map((item: any) => {
      if (String(item?.id ?? '').trim() !== id) return item;
      changed = true;
      return patch(item);
    });
    if (!changed) return;

    this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, next);
  }

  /** Remove a task from every in-memory folder-task snapshot (HR + All Files). */
  private removeTaskFromFolderCaches(taskId: string): void {
    this.mapTaskInFolderCaches(taskId, () => null);
  }

  /** Patch a task inside every in-memory folder-task snapshot that contains it. */
  private patchTaskInFolderCaches(taskId: string, patch: (item: any) => any): void {
    this.mapTaskInFolderCaches(taskId, patch);
  }

  /**
   * Walk HR / lib folder-task caches. Return `null` from `mapFn` to remove the item.
   */
  private mapTaskInFolderCaches(taskId: string, mapFn: (item: any) => any | null): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;

    for (const prefix of ['tasks:hr-folder:', 'tasks:lib:v7:'] as const) {
      for (const { key, data } of this.fileCrawlCache.getStaleEntriesByPrefix(prefix)) {
        if (!Array.isArray(data) || data.length === 0) continue;

        let changed = false;
        const next: any[] = [];
        for (const item of data) {
          if (String(item?.id ?? '').trim() !== id) {
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

  private applyDueDatePatch(item: any, dueDate: string): any {
    const details = { ...(item?.eFormDetails ?? {}) };
    const rawFields = { ...(details.rawFields ?? {}) };
    let wroteRaw = false;
    for (const key of ['DueDate', 'DueDate0', 'Due', 'Due_x0020_Date'] as const) {
      if (Object.prototype.hasOwnProperty.call(rawFields, key)) {
        rawFields[key] = dueDate;
        wroteRaw = true;
        break;
      }
    }
    if (!wroteRaw) rawFields['DueDate'] = dueDate;
    details.dueDate = dueDate;
    details.rawFields = rawFields;
    return { ...item, dueDate, eFormDetails: details };
  }

  private applyAssigneePatch(item: any, assignedToName: string, assignedToEmail = ''): any {
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

  private isAssigneeCurrentUser(assignedToName: string, assignedToEmail = ''): boolean {
    const email = (this.currentUser?.email ?? '').toLowerCase();
    const upn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();
    const name = (this.currentUser?.username ?? '').toLowerCase();
    const candEmail = String(assignedToEmail ?? '').trim().toLowerCase();
    const candName = String(assignedToName ?? '').trim().toLowerCase();
    if (candEmail && (candEmail === email || candEmail === upn)) return true;
    if (candName && name && candName === name) return true;
    return false;
  }

  /** Best-effort folder key invalidation when list/folder can be inferred from the task. */
  private invalidateRelatedFolderTaskCache(task: any): void {
    const folder =
      String(this.selectedFolderName ?? '').trim() ||
      String(this.selectedHrPersonalTaskFolder ?? '').trim() ||
      String(task?.submittedBy ?? task?.eFormDetails?.submittedBy ?? '').trim();
    if (!folder) return;

    const listName = String(
      task?.listName ?? task?.eFormDetails?.listName ?? this.currentLibraryName ?? '',
    ).trim();
    if (listName) {
      this.invalidateFolderTaskCache(listName, folder);
    }
    // HR person folders are keyed under HRPersonal regardless of the task list name.
    this.invalidateFolderTaskCache(this.targetLibraryName, folder);
  }

  /**
   * Keep Redis (and in-memory) task snapshots aligned after an app ? SharePoint write.
   * Also drops the backend `/api/tasks` read-through key when list + email are known.
   */
  private syncTaskCachesAfterMutation(
    task: any,
    kind: 'complete' | 'dueDate' | 'delegate' | 'comment',
    extras?: { dueDate?: string; assignedToName?: string; assignedToEmail?: string },
  ): void {
    const taskId = String(task?.id ?? task?.Id ?? '').trim();
    if (!taskId) return;

    if (kind === 'complete') {
      this.removeTaskFromTodoCache(taskId);
      this.removeTaskFromFolderCaches(taskId);
      this.invalidateRelatedFolderTaskCache(task);
    } else if (kind === 'dueDate' && extras?.dueDate) {
      const dueDate = extras.dueDate;
      this.patchTaskInTodoCache(taskId, (item) => this.applyDueDatePatch(item, dueDate));
      this.patchTaskInFolderCaches(taskId, (item) => this.applyDueDatePatch(item, dueDate));
    } else if (kind === 'delegate') {
      const name = String(extras?.assignedToName ?? '').trim();
      const email = String(extras?.assignedToEmail ?? '').trim();
      if (!this.isAssigneeCurrentUser(name, email)) {
        this.removeTaskFromTodoCache(taskId);
      } else {
        this.patchTaskInTodoCache(taskId, (item) => this.applyAssigneePatch(item, name, email));
      }
      this.patchTaskInFolderCaches(taskId, (item) => this.applyAssigneePatch(item, name, email));
      this.invalidateRelatedFolderTaskCache(task);
    } else if (kind === 'comment') {
      // Action comments may change assignee (handled via delegate); still refresh folder snapshots.
      this.invalidateRelatedFolderTaskCache(task);
    }

    this.invalidateBackendTasksCache(task);
  }

  /** Fire-and-forget invalidate of the Node `/api/tasks` Redis read-through entry. */
  private invalidateBackendTasksCache(task: any): void {
    const list = String(task?.listName ?? task?.eFormDetails?.listName ?? '').trim();
    const email = String(
      task?.userEmail ??
        task?.eFormDetails?.submittedByEmail ??
        this.currentUser?.email ??
        '',
    ).trim();
    if (!list || !email) return;
    void this.backendApi.invalidateTasksForPerson(email, list).catch(() => {});
  }

  private countAssignedHrTasks(): number {
    const cached = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
    return cached.filter((item: any) => item?.isAssignedToCurrentUser === true).length;
  }

  private mergeHrTasks(existing: any[], incoming: any[]): any[] {
    const byKey = new Map<string, any>();
    for (const it of existing) byKey.set(this.getHrTaskKey(it), it);
    for (const it of incoming) byKey.set(this.getHrTaskKey(it), it);
    return [...byKey.values()].sort(
      (a, b) =>
        new Date(b.submittedDate || 0).getTime() - new Date(a.submittedDate || 0).getTime()
    );
  }

  private async loadOlderHrTasks(cutoffIso: string): Promise<void> {
    if (this.hrOlderHistoryLoadInProgress) return;
    this.hrOlderHistoryLoadInProgress = true;
    this.refreshView();
    try {
      const older = await this.loadHrTasksFromLists(false, { createdBeforeIso: cutoffIso });
      if (!older.length || this.isFolderTaskCommentsViewActive()) return;
      this.commentItems = this.mergeHrTasks(this.commentItems, older);
      this.invalidateCommentFilters();
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(this.commentItems));
    } catch { /* ignore */ }
    finally {
      this.hrOlderHistoryLoadInProgress = false;
      this.refreshView();
    }
  }

  // ============================================================
  // LOAD HR TASKS - public entry point
  // ============================================================
  protected loadHrTasks(): void {
    if (!this.currentUser) {
      this.showModal('Login Required', 'You must log in first to view HR Tasks.', 'error');
      this.commentsMessage = 'Please log in first.';
      return;
    }

    if (!this.showToDoSection && this.isFolderTaskCommentsViewActive()) {
      const folder = String(this.selectedHrPersonalTaskFolder || this.selectedFolderName || '').trim();
      const library = this.showAllFilesSection
        ? String(this.currentLibraryName ?? '').trim()
        : this.targetLibraryName;
      if (folder && library) {
        this.loadAllFilesTasksForLibrary(library, folder);
        return;
      }
    }

    const cached = this.fileCrawlCache.get(AppComponent.HR_USER_TASKS_CACHE_KEY);
    if (cached?.length) {
      this.commentItems = this.sortCommentItemsByDateDesc(cached);
      this.visibleItemCount = this.commentsInitialPageSize;
      this.invalidateCommentFilters();
      this.isLoadingComments = false;
      this.isLoadingMoreComments = false;
      this.commentsMessage = '';
      this.userHrTasksLoaded = true;
      this.refreshView();
      return;
    }

    const loadSeq = ++this.hrCommentsLoadSeq;
    this.isLoadingComments = true;
    this.isLoadingMoreComments = false;
    this.commentItems = [];
    this.visibleItemCount = this.commentsInitialPageSize;
    this.invalidateCommentFilters();
    this.refreshView();
    void this.loadHrTasksWithRecentWindow(loadSeq);
  }

  private isStaleHrCommentsLoad(loadSeq: number): boolean {
    return loadSeq !== this.hrCommentsLoadSeq;
  }

  private async loadHrTasksWithRecentWindow(loadSeq: number): Promise<void> {
    try {
      await this.ensureFormConfigLoaded();
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
        this.commentsMessage = 'Failed to load HR Tasks.';
        this.isLoadingComments = false;
        this.refreshView();
      }
    }
  }

  // ============================================================
  // MODAL STATE
  // ============================================================
  modal = {
    visible: false,
    title: '',
    message: '',
    type: 'info' as 'info' | 'success' | 'warning' | 'error',
    action: null,
  };

  showModal(
    title: string,
    message: string,
    type: 'info' | 'success' | 'warning' | 'error' = 'info',
  ): void {
    this.modal = { visible: true, title, message, type, action: null };
    this.refreshView();
  }

  closeModal(): void {
    this.modal.visible = false;
    this.modal.action = null;
    this.refreshView();
  }

  // ============================================================
  // LOAD HR TASKS - core async logic
  // All SharePoint lists queried IN PARALLEL via Promise.all().
  // Optional Graph $filter on list item createdDateTime reduces pages
  // for the "recent" and silent-refresh paths.
  // ============================================================
  private async loadHrTasksFromLists(
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

    const publishPartialTodo = (partialItems: any[]): void => {
      if (!updateTodo || !progressiveTodo || isStaleTodoLoad()) return;
      if (partialItems.length === 0) return;
      const existing = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
      const merged = existing.length > 0
        ? this.mergeHrTasks(existing, partialItems)
        : [...partialItems];
      const sorted = [...merged].sort(
        (a: { submittedDate: any }, b: { submittedDate: any }) =>
          new Date(b.submittedDate || 0).getTime() - new Date(a.submittedDate || 0).getTime()
      );
      this.fileCrawlCache.set(AppComponent.HR_USER_TASKS_CACHE_KEY, sorted);
      this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(sorted), { clearLoading: false });
      this.userHrTasksLoaded = true;
      this.refreshView();
    };

    const isCacheableRecentUserLoad =
      !subordinateTasksOnly &&
      !!range?.createdSinceIso &&
      !range?.createdBeforeIso;

    // Serve recent user tasks from cache unless this is a forced network refresh.
    if (isCacheableRecentUserLoad && !options.skipCache) {
      const cached = this.fileCrawlCache.get(AppComponent.HR_USER_TASKS_CACHE_KEY);
      const cacheOk = !todoScope || this.todoScopeCacheValid;
      if (cached?.length && cacheOk) {
        if (updateCommentItems && !isStaleCommentsLoad()) {
          if (!this.isFolderTaskCommentsViewActive()) {
            this.commentItems = cached;
            this.invalidateCommentFilters();
            this.commentsMessage = '';
            this.isLoadingComments = false;
            this.userHrTasksLoaded = true;
            this.refreshView();
          } else {
            this.isLoadingComments = false;
            this.refreshView();
          }
        }
        if (updateTodo) {
          this.todoService.setTasksFromSharePoint(this.getHrTasksForTodoList(cached));
          this.userHrTasksLoaded = true;
        }
        return cached;
      }
    }

    try {
      const token = await this.getSharePointToken();
      const siteHost = sharePointConfig.siteHostName;
      const sitePath = sharePointConfig.sitePath;

      await this.ensureSiteMetadata(siteHost, sitePath, token);

      const siteId: string = this.cachedSiteId!;
      const siteWebUrl: string = this.cachedSiteWebUrl;
      const siteListsArr: any[] = this.cachedSiteLists;

      const userEmail = (this.currentUser?.email ?? '').toLowerCase();
      const userUpn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();

      if (subordinateTasksOnly) {
        await this.subordinaryTaskService.ensureSubordinatesForManager([
          this.currentUser?.email ?? '',
          this.currentUser?.userPrincipalName ?? '',
          this.currentUser?.username ?? '',
        ].map(value => value.trim()).filter(Boolean), true);
      }

      let listsToQuery = this.getTaskListsToQuery(siteListsArr, {
        includeAllWorkflows: todoScope,
        includeArchiveCompanions: todoScope,
      })
        .filter((listName: string) =>
          siteListsArr.some((l: any) =>
            (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
            (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
          )
        );

      if (restrictToAssigneeBlockedLists) {
        listsToQuery = listsToQuery.filter((listName: string) => {
          // Procurement lists are already covered by the pending/newest-first pass;
          // re-scanning them here would just page through their oldest rows again.
          if (isTodoProcurementTaskList(listName)) return false;
          const list = siteListsArr.find((l: any) =>
            (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
            (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
          );
          return !!list?.id && this.taskQuery.blockedAssigneeLookupListIds.has(list.id);
        });
      } else {
        listsToQuery = listsToQuery.sort((a: string, b: string) => {
          const rank = (name: string): number => {
            const key = name.toLowerCase();
            if (key === 'proctasksarchive') return -1;
            if (isTodoProcurementTaskList(name)) return 0;
            const list = siteListsArr.find((l: any) =>
              (l.name ?? '').toLowerCase() === name.toLowerCase() ||
              (l.displayName ?? '').toLowerCase() === name.toLowerCase()
            );
            if (!list?.id) return 3;
            if (this.taskQuery.getListFilterFieldStatus(list.id, 'AssignedToLookupId') === true) return 1;
            if (this.taskQuery.blockedAssigneeLookupListIds.has(list.id)) return 3;
            return 2;
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

      const fetchOneList = (listName: string) =>
        this.fetchItemsForList(
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
              const items = await fetchOneList(listsToQuery[index]);
              merged.push(...items);
              publishPartialTodo(merged);
            }
          },
        );
        await Promise.all(workers);
        listResults.push(merged);
      } else {
        for (let i = 0; i < listsToQuery.length; i += listConcurrency) {
          const batch = listsToQuery.slice(i, i + listConcurrency);
          const batchResults = await Promise.all(batch.map(listName => fetchOneList(listName)));
          listResults.push(...batchResults);
          publishPartialTodo(listResults.flat());
          if (i + listConcurrency < listsToQuery.length && !progressiveTodo) {
            await this.sleep(AppComponent.HR_TASK_LIST_BATCH_GAP_MS);
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
            const existing = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
            this.fileCrawlCache.set(
              AppComponent.HR_USER_TASKS_CACHE_KEY,
              existing.length > 0 ? this.mergeHrTasks(existing, allItems) : allItems,
            );
          }
          // Do NOT set todoScopeCacheValid here - progressive To Do still has Proc*
          // drain / blocked-list fallback after this pass. Callers mark complete.
        } else if (allItems.length > 0) {
          const existing = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
          const procItems = existing.filter((item: any) =>
            isTodoProcurementTaskList(String(item.listName ?? ''))
          );
          this.fileCrawlCache.set(
            AppComponent.HR_USER_TASKS_CACHE_KEY,
            this.mergeHrTasks(allItems, procItems),
          );
        }
      }

      if (updateCommentItems) {
        if (isStaleCommentsLoad()) return allItems;
        if (this.isFolderTaskCommentsViewActive()) {
          this.isLoadingComments = false;
          this.refreshView();
          return allItems;
        }
        this.commentItems = allItems;
        this.invalidateCommentFilters();
        this.commentsMessage = '';
        this.isLoadingComments = false;
        this.userHrTasksLoaded = true;
        this.refreshView();
        if (kickoffOlderBackfill && range?.createdSinceIso && !isStaleCommentsLoad()) {
          void this.loadOlderHrTasks(range.createdSinceIso);
        }
      }

      if (updateTodo) {
        if (!isStaleTodoLoad()) {
          if (allItems.length > 0) {
            const existing = this.fileCrawlCache.getStale(AppComponent.HR_USER_TASKS_CACHE_KEY) ?? [];
            const merged = existing.length > 0 ? this.mergeHrTasks(existing, allItems) : allItems;
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
          this.commentsMessage = `Failed to load HR Tasks: ${(error as any)?.message || 'Unknown error'}.`;
          this.isLoadingComments = false;
          this.refreshView();
        }
      }
      if (updateTodo) {
        this.todoService.setLoading(false);
      }
      return [];
    }
  }
  // ============================================================
  // FETCH ITEMS FOR ONE LIST (paginated)
  // ============================================================
  private buildHrTaskFilterQuery(
    range: { createdSinceIso?: string; createdBeforeIso?: string } | null
  ): string {
    const filterParts: string[] = [];
    if (range?.createdSinceIso) {
      // Important: HR "recent" visibility should be based on *last change*
      // (Approver actions, uploads, edits) not only SharePoint "createdDateTime".
      // Otherwise tasks completed/modified after the cutoff are missed.
      filterParts.push(`lastModifiedDateTime ge ${range.createdSinceIso}`);
    }
    if (range?.createdBeforeIso) {
      filterParts.push(`lastModifiedDateTime lt ${range.createdBeforeIso}`);
    }
    return filterParts.length > 0
      ? `&$filter=${encodeURIComponent(filterParts.join(' and '))}`
      : '';
  }

  private filterTasksForSelectedAllFilesFolder(items: any[]): any[] {
    if (!this.showAllFilesSection) return items;
    const folderName = String(this.selectedFolderName ?? '').trim();
    if (!folderName || !extractTrailingFolderId(folderName)) return items;
    return items.filter(item => this.folderMatch.doesMappedHrTaskMatchFolder(item, folderName));
  }


  /** Loads HR tasks from SharePoint lists that belong to the person represented by folderName. */
  private async loadHrPersonalTasksForFolder(
    folderName: string,
    loadSeq: number,
    options: { softRefresh?: boolean } = {},
  ): Promise<void> {
    const softRefresh = !!options.softRefresh;
    const folderCacheKey = this.hrFolderTaskCacheKey(folderName);
    const mergeWithCachedFolderTasks = (incoming: any[]): any[] => {
      if (!softRefresh) return incoming;
      const cached = this.fileCrawlCache.getStale<any[]>(folderCacheKey);
      if (!Array.isArray(cached) || cached.length === 0) return incoming;
      return this.mergeHrTasks(cached, incoming);
    };
    try {
      this.folderTaskGraphInFlight += 1;
      // Do not wait for To Do — people with 800+ tasks were blocked for minutes
      // before Comments even started. Folder click already cancelled To Do Graph.
      const tokenPromise = this.getSharePointToken();
      const formPromise = this.ensureFormConfigLoaded();
      const token = await tokenPromise;
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
      await formPromise;
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

      await this.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );

      const listsToQuery = this.prioritizeHrPersonTaskLists(
        this.getTaskListsToQuery(this.cachedSiteLists).filter(listName =>
          this.cachedSiteLists.some((l: any) =>
            (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
            (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
          )
        ),
      );
      if (!softRefresh) {
        this.commentsMessage = 'Loading tasks...';
        this.refreshView();
        // Some people have no hits on the first lists — don't leave the blank
        // full-screen spinner up for the whole fan-out. Switch to inline status.
        window.setTimeout(() => {
          if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
          if (!this.isLoadingComments || this.commentItems.length > 0) return;
          this.isLoadingComments = false;
          this.isLoadingMoreComments = true;
          this.commentsMessage = 'Looking up tasks…';
          this.refreshView();
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
        const items = this.applyAllFilesFolderSubmitter(mergeWithCachedFolderTasks(mapped));
        const sorted = this.sortCommentItemsByDateDesc(items);

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
            this.isLoadingMoreComments = true;
            this.refreshView();
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

      const fastMapped = await this.collectHrPersonalTasksFromLists(
        listsToQuery,
        folderName,
        token,
        loadSeq,
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
        this.isLoadingMoreComments = false;
        this.refreshView();
        return;
      }

      publishMapped(fastMapped, { background: true, complete: true });
      this.hrFolderCompleteAt.set(folderCacheKey, Date.now());
      this.lastHrFilesCommentsFolder = folderName;
    } catch (err: any) {
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
      console.error(`[HR Files] Failed loading tasks for "${folderName}"`, err);
      if (softRefresh) return;
      this.isLoadingMoreComments = false;
      this.isLoadingComments = false;
      // Keep whatever is already painted; only show a message when the panel is empty.
      if (this.commentItems.length === 0) {
        this.invalidateCommentFilters();
        this.commentsMessage = 'Could not load tasks. Press refresh to try again.';
      }
      this.refreshView();
    } finally {
      this.folderTaskGraphInFlight = Math.max(0, this.folderTaskGraphInFlight - 1);
      if (!this.isStaleAllFilesFolderTaskLoad(loadSeq) && (this.isLoadingComments || this.isLoadingMoreComments)) {
        this.isLoadingComments = false;
        this.isLoadingMoreComments = false;
        this.refreshView();
      }
    }
  }

  private mapHrFolderRawItems(
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

  private async collectHrPersonalTasksFromLists(
    listsToQuery: string[],
    folderName: string,
    token: string,
    loadSeq: number,
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
    const userEmail = (this.currentUser?.email ?? '').toLowerCase();
    const userUpn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();
    const concurrency = AppConstants.hrFilesTaskListConcurrency;
    const softRefresh = options?.softRefresh === true;
    const folderPin = String(folderName ?? '').match(/^\d+/)?.[0] ?? '';
    // Email + LookupId never change for a folder — shared Redis skips two Graph
    // round-trips on every cold open after anyone has resolved this person once.
    const idsKey = this.hrPersonIdsCacheKey(folderName);
    const cachedIds = this.fileCrawlCache.getStale<{ email: string; lookupId: string | null }>(idsKey);
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
      const count = this.flattenHrMappedByList(mappedByList).length;
      const seconds = ((Date.now() - crawlStartedAt) / 1000).toFixed(1);
      console.info(`[HR Files] ${crawlMode} "${folderName}" ${phase}: ${count} tasks at ${seconds}s${detail}`);
    };

    const publishListHits = (listName: string, mapped: any[]): void => {
      if (mapped.length === 0) return;
      const existing = mappedByList.get(listName) ?? [];
      mappedByList.set(listName, this.mergeHrTasks(existing, mapped));
      onPartial?.(this.flattenHrMappedByList(mappedByList));
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
                if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
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
                  if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
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
                if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
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
      if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;

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

    if (options?.interactive) {
      // Click path: worker pool with no inter-batch sleep so a slow list never
      // stalls the rest and first paint is not delayed by idle gaps.
      let next = 0;
      const workers = Array.from(
        { length: Math.min(AppConstants.hrFilesInteractiveListConcurrency, listsToQuery.length) },
        async () => {
          while (!this.isStaleAllFilesFolderTaskLoad(loadSeq)) {
            const index = next++;
            if (index >= listsToQuery.length) break;
            await processList(listsToQuery[index]);
            if (onPartial && mappedByList.size > 0) {
              onPartial(this.flattenHrMappedByList(mappedByList));
            }
          }
        },
      );
      await Promise.all(workers);
    } else {
      for (let i = 0; i < listsToQuery.length; i += concurrency) {
        if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) break;
        const batch = listsToQuery.slice(i, i + concurrency);
        await Promise.all(batch.map(listName => processList(listName)));

        if (onPartial && mappedByList.size > 0) {
          onPartial(this.flattenHrMappedByList(mappedByList));
        }
        // Pacing gaps are for background prefetch only; the shared Graph limiter already
        // caps what a click or refresh has in flight.
        if (!crawlMode && i + concurrency < listsToQuery.length) {
          await this.sleep(AppComponent.HR_TASK_LIST_BATCH_GAP_MS);
        }
      }
    }

    // Interactive seed: newest LookupId/Title pages only — paints recent tasks first.
    const seedItems = this.flattenHrMappedByList(mappedByList);
    options?.onSeedComplete?.(seedItems);
    logCrawlPhase('first pass', ` (${listsToQuery.length} lists)`);

    // Deep crawl (cold and soft refresh): full assignee rows and the newest unfiltered pages
    // for every list, plus a LookupId re-read only where the first pass left rows unread.
    if (personLookupId && !this.isStaleAllFilesFolderTaskLoad(loadSeq)) {
      const listsForTopUp = [...listObjByName.entries()];
      const deepLookupPages = AppConstants.hrFilesPersonLookupMaxPages;
      for (let i = 0; i < listsForTopUp.length; i += concurrency) {
        if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) break;
        const batch = listsForTopUp.slice(i, i + concurrency);
        await Promise.all(
          batch.map(async ([listName, listObj]) => {
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
                        if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
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
                    if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
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

            if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) return;
            if (deepLookupItems?.length) {
              publishListHits(listName, mapRaw(deepLookupItems));
            }
            if (assigneeItems?.length) {
              publishListHits(listName, mapRaw(assigneeItems));
            }
            if (supplement.length) {
              publishListHits(listName, mapRaw(supplement));
            }
          }),
        );

        if (onPartial && mappedByList.size > 0) {
          onPartial(this.flattenHrMappedByList(mappedByList));
        }
        if (!crawlMode && i + concurrency < listsForTopUp.length) {
          await this.sleep(AppComponent.HR_TASK_LIST_BATCH_GAP_MS);
        }
      }
    }

    logCrawlPhase('deep crawl', ` (${lookupIncompleteLists.size} lists re-read)`);

    // Second pass: workflow siblings that share the same eForm ID AND the same task name.
    const seedTasksForSiblings = this.preferCurrentUserSeedTasks(
      this.flattenHrMappedByList(mappedByList),
    );
    const eFormKeyToTaskNames = collectEFormKeyToTaskNames(seedTasksForSiblings);
    if (eFormKeyToTaskNames.size > 0 && !this.isStaleAllFilesFolderTaskLoad(loadSeq)) {
      const listsWithHits = [...mappedByList.keys()];
      const backfillPages = AppConstants.hrFilesTaskListBackfillPageLimit;
      for (let i = 0; i < listsWithHits.length; i += concurrency) {
        if (this.isStaleAllFilesFolderTaskLoad(loadSeq)) break;
        const batch = listsWithHits.slice(i, i + concurrency);
        await Promise.all(batch.map(async listName => {
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
          mappedByList.set(listName, this.mergeHrTasks(existing, relatedMapped));
        }));

        if (onPartial) {
          onPartial(this.flattenHrMappedByList(mappedByList));
        }
        if (!crawlMode && i + concurrency < listsWithHits.length) {
          await this.sleep(AppComponent.HR_TASK_LIST_BATCH_GAP_MS);
        }
      }
    }

    logCrawlPhase(
      this.isStaleAllFilesFolderTaskLoad(loadSeq) ? 'cancelled' : 'finished',
      ` (${eFormKeyToTaskNames.size} ids x ${mappedByList.size} lists for related steps)`,
    );
    return this.flattenHrMappedByList(mappedByList);
  }

  private flattenHrMappedByList(mappedByList: Map<string, any[]>): any[] {
    return [...mappedByList.values()]
      .flat()
      .sort((a: any, b: any) =>
        new Date(b.submittedDate || 0).getTime() - new Date(a.submittedDate || 0).getTime(),
      );
  }

  private collectEFormKeysFromMappedTasks(items: any[], folderName?: string): Set<string> {
    const folderEFormId = folderName ? extractTrailingFolderId(folderName) : '';
    const keys = new Set<string>();
    for (const item of items) {
      for (const key of extractEFormKeysFromMappedItem(item)) {
        if (folderEFormId && key !== folderEFormId) continue;
        keys.add(key);
      }
    }
    if (folderEFormId) keys.add(folderEFormId);
    return keys;
  }

  /**
   Prefer the logged-in user's own submissions as Pass-2 seeds.
   If none are present (e.g. browsing another person's folder), fall back to all seed tasks.
   */
  private preferCurrentUserSeedTasks(seedTasks: any[]): any[] {
    const own = seedTasks.filter(task => this.isMappedTaskSubmittedByCurrentUser(task));
    return own.length > 0 ? own : seedTasks;
  }

  private isMappedTaskSubmittedByCurrentUser(item: any): boolean {
    const email = (this.currentUser?.email ?? '').toLowerCase();
    const upn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();
    const display = (this.currentUser?.username ?? '').toLowerCase();
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

  /** Map each eForm id ? normalized task/form names from the seed submissions. */

  /**
   Normalize titles so "Please Approve Missing Punch eForm for - X" and
   "Missing Punch Finalised" resolve to the same chain name.
   */




  /** Related raw item must share an eForm id AND the same task/form name as that seed. */


  /** eFormListId field values + IDs embedded in comment text (ID=9699 / ID: 9699). */

  private async fetchProcurementTodoPendingItems(
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

  /** In All Files, keep each task/comment's real submitter (matches Live). */
  private applyAllFilesFolderSubmitter<T extends { submittedBy?: string; eFormDetails?: Record<string, unknown> }>(
    items: T[],
  ): T[] {
    return items;
  }

  /** OData-safe Title filter for document-library folder tasks (All Files only). */
  private buildDocumentLibraryFolderTitleFilter(folderName: string): string | null {
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
  private async loadDocumentLibraryTasksForFolder(
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
      const token = await this.getSharePointToken();
      if (isStale()) return;

      await this.ensureSiteMetadata(
        sharePointConfig.siteHostName,
        sharePointConfig.sitePath,
        token,
      );

      const siteId = this.cachedSiteId!;
      let listObj = this.cachedSiteLists.find((l: any) =>
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
          listObj = direct.list;
          if (!this.cachedSiteLists.some((l: any) => l?.id === direct.list.id)) {
            this.cachedSiteLists = [...this.cachedSiteLists, direct.list];
          }
        }
      }

      if (!listObj?.id) {
        if (isStale()) return;
        if (softRefresh || silent) return;
        console.warn(`[All Files] Task list "${listName}" unavailable for this user (${lookupStatus}).`);
        this.commentsMessage =
          lookupStatus === 'forbidden'
            ? `You don't have access to the task list "${listName}". Ask a SharePoint admin to grant you read access.`
            : `Could not find task list "${listName}".`;
        this.commentItems = [];
        this.isLoadingComments = false;
        this.refreshView();
        return;
      }

      const userEmail = (this.currentUser?.email ?? '').toLowerCase();
      const userUpn = (this.currentUser?.userPrincipalName ?? '').toLowerCase();
      const collectedMapped: any[] = [];
      const seenIds = new Set<string>();
      const pageSize = AppConstants.docLibraryFolderTaskPageSize;
      const maxPages = AppConstants.docLibraryFolderTaskMaxPages;

      const publishAndCache = (done: boolean): void => {
        if (isStale()) return;
        if (softRefresh && collectedMapped.length === 0) return;
        const items = this.applyAllFilesFolderSubmitter(collectedMapped);
        this.fileCrawlCache.set(
          this.libFolderTaskCacheKey(listName, folderName),
          this.sortCommentItemsByDateDesc(items),
        );
        if (silent) {
          if (done) {
            this.allFilesComponent?.refreshCommentSearchFromCaches();
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
              this.cachedSiteWebUrl,
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

          if (!silent && this.isLoadingComments) {
            this.commentsMessage =
              pagesLoaded === 1
                ? `Scanning ${sourceListName}...`
                : `Scanning ${sourceListName} (page ${pagesLoaded})...`;
            this.refreshView();
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
      const archiveCompanions = this.getArchiveCompanionTaskLists(listName);
      if (collectedMapped.length === 0 && archiveCompanions.length > 0) {
        for (const archiveName of archiveCompanions) {
          const archiveList = this.findCachedSiteListByName(archiveName);
          if (!archiveList?.id) {
            console.warn(`[All Files] Archive list "${archiveName}" not found on site`);
            continue;
          }
          if (!silent && this.isLoadingComments) {
            this.commentsMessage = `Checking ${archiveName}...`;
            this.refreshView();
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
              : this.findCachedSiteListByName(scanName);
          if (!scanList?.id) continue;
          await fullScanList(scanName, scanList);
          if (isStale()) return;
          if (collectedMapped.length > 0) break;
        }
      }

      if (isStale()) return;

      // Related Superior / sibling steps via eForm id - check primary + archive lists.
      const eFormKeys = this.collectEFormKeysFromMappedTasks(collectedMapped, folderName);
      if (eFormKeys.size > 0) {
        const relatedListTargets = [
          { name: listName, list: listObj },
          ...archiveCompanions
            .map((name) => ({ name, list: this.findCachedSiteListByName(name) }))
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
      this.isLoadingMoreComments = false;
      this.commentsMessage = `Failed to load tasks: ${err?.message || 'Unknown error'}`;
      this.isLoadingComments = false;
      this.refreshView();
    } finally {
      this.folderTaskGraphInFlight = Math.max(0, this.folderTaskGraphInFlight - 1);
      if (!silent && !isStale() && (this.isLoadingComments || this.isLoadingMoreComments)) {
        this.isLoadingComments = false;
        this.isLoadingMoreComments = false;
        this.refreshView();
      }
    }
  }

  private async ensureSiteMetadata(
    _siteHost: string,
    _sitePath: string,
    token: string,
  ): Promise<void> {
    if (this.cachedSiteId && this.cachedSiteLists.length > 0) return;

    const metadata = await this.siteMetadataService.resolve(token);
    this.cachedSiteId = metadata.siteId;
    this.cachedSiteWebUrl = metadata.siteWebUrl;
    this.cachedSiteLists = metadata.lists;
    this.taskQuery.setSiteContext(metadata.siteId, metadata.lists);
  }

  /**
   * The one place that decides how a To Do list is queried, so the first load,
   * "Load older tasks" and the manual refresh can never drift apart. Always
   * newest-modified first; procurement lists additionally narrow to pending rows.
   */
  private buildTodoListStartPath(siteId: string, listId: string, listName: string): string {
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

  private todoPreferHeader(listName: string): Record<string, string> | undefined {
    return isTodoProcurementTaskList(listName)
      ? { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }
      : undefined;
  }

  /** Advance a list's high-water mark to the newest row seen. */
  private noteTodoWatermark(listName: string, items: any[]): void {
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
  private async fetchTodoRowsNewerThanWatermark(
    siteId: string,
    listId: string,
    token: string,
    listName: string,
    watermarkIso: string,
  ): Promise<any[]> {
    const prefer = this.todoPreferHeader(listName);
    let nextPath: string | null = this.buildTodoListStartPath(siteId, listId, listName);
    const fresh: any[] = [];
    let pages = 0;

    while (nextPath && pages < AppConstants.todoRefreshMaxPagesPerList) {
      const page: any = await graphGetWithRetry(
        this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs, prefer ?? {},
      );

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

  /**
   * To Do scan for a list that cannot be filtered by assignee: newest-modified first.
   * Hard-capped via `todoInitialPagesPerList`; leftover pages stay on `todoListCursors`
   * for manual "Load older" instead of draining the whole list on login.
   */
  private async fetchTodoScanPages(
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

  private async fetchItemsForList(
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
  ): Promise<any[]> {
    const list = siteListsArr.find(
      (l: any) =>
        (l.name ?? '').toLowerCase() === listName.toLowerCase() ||
        (l.displayName ?? '').toLowerCase() === listName.toLowerCase()
    );
    if (!list?.id) return [];

    // Procurement lists are exempt: if their assignee filter is unavailable they still
    // need the page-scan fallback below, otherwise they contribute nothing at all.
    if (
      todoAssigneeOnly &&
      assigneeLookupId &&
      !isTodoProcurementTaskList(listName) &&
      this.taskQuery.blockedAssigneeLookupListIds.has(list.id)
    ) {
      return [];
    }

    const isOlderBackfill = !!range?.createdBeforeIso;
    const top = isOlderBackfill ? 999 : AppConstants.hrTasksFastLoadPageSize;
    const maxPages = isOlderBackfill
      ? Number.POSITIVE_INFINITY
      : (fastLoadPageLimit ?? AppConstants.hrTasksFastLoadPageLimit);

    let rawItems: any[];
    let skipUserFilter = false;
    try {
      if (todoAssigneeOnly && isTodoProcurementTaskList(listName)) {
        // Graph cannot filter these lists by assignee at all, so narrow on pending +
        // newest-modified instead and match the assignee client-side. Follow every
        // Progress=Pending page (Live parity); unordered fallback stays page-capped.
        const pendingItems = await this.fetchProcurementTodoPendingItems(
          siteId,
          list.id,
          token,
          listName,
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

  /** True for All Files task lists (e.g. DamagesToEnemaltaTasks) where Title is the folder name, not the submitter. */

  /** Strip workflow boilerplate accidentally captured with a parsed person name. */


  private getLoggedInUserDisplayName(): string {
    return this.userService.getCurrentUserName();
  }


  /** Collect every assignee-like value stored on a SharePoint task item. */


  // ============================================================
  // EFORM CONTENT MODAL
  // ============================================================
  protected selectedEFormContent: {
    fileName: string;
    content: string;
    contentType: string;
    uploadDate: string;
    eFormDetails?: any;
    detailRows?: { label: string; value: string; sectionBreak?: boolean }[];
    detailView?: HrTaskDetailView;
  } | null = null;
  protected isLoadingEFormContent = false;

  // Handles the delegate form submission. The SharePoint write is delegated to
  // the service, and the UI is only updated after that write is verified - so an
  // unverified/failed write shows an error instead of a fake success.
  protected async onDelegateTask(payload: { task: any; newAssignee: any }): Promise<void> {
    const task = payload?.task;
    const newAssignee = payload?.newAssignee;

    if (!task || !newAssignee) {
      console.error('Invalid delegation payload');
      return;
    }

    const newAssigneeName: string = newAssignee.Title;
    const newAssigneeEmail: string = newAssignee.Email || newAssignee.UserPrincipalName || '';

    try {
      // Persist to SharePoint first; this resolves only when the change is verified.
      await this.delegateService.updateAssignedTo(task, newAssignee);
      this.onDelegateSuccess(task.id, newAssigneeName, newAssigneeEmail);
    } catch (error: any) {
      console.error('Error delegating task:', error);
      const msg = error?.message || 'Unknown error';
      this.showModal('Delegate Failed', `Could not update SharePoint: ${msg}`, 'error');
    }
  }

  private onDelegateSuccess(itemId: string, newAssigneeName: string, newAssigneeEmail = ''): void {
    // Update the item in-place so the UI reflects the real SharePoint state.
    // Removing it locally was masking failed writes - on refresh the item would
    // reappear because SharePoint still had the old assignee.
    const item = this.commentItems.find(c => c.id === itemId);
    if (item) {
      if (item.eFormDetails) {
        item.eFormDetails = { ...item.eFormDetails, assignedTo: newAssigneeName || newAssigneeEmail };
      }
      (item as any).assignedTo = newAssigneeName || newAssigneeEmail;
    }
    this.syncTaskCachesAfterMutation(item ?? { id: itemId }, 'delegate', {
      assignedToName: newAssigneeName,
      assignedToEmail: newAssigneeEmail,
    });
    this.invalidateCommentFilters();
    this.refreshView();
    this.showModal('Task Delegated', `Task has been delegated to ${newAssigneeName}.`, 'success');
  }

  private buildHrTaskDetailRows(
    d: any,
    task: any,
  ): { label: string; value: string; sectionBreak?: boolean }[] {
    const clean = (v: any) =>
      !v || String(v).trim() === '' || String(v).trim() === 'Enter value here'
        ? '' : String(v).trim();

    const rows: { label: string; value: string; sectionBreak?: boolean }[] = [];
    const push = (label: string, value: string) => {
      if (value) rows.push({ label, value });
    };
    const endSection = () => {
      if (rows.length) rows[rows.length - 1].sectionBreak = true;
    };

    const fmtDate = (val: string) => {
      const dt = new Date(val);
      return isNaN(dt.getTime())
        ? val
        : dt.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
        ' ' + dt.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    };

    const statusLower = clean(d.status).toLowerCase();
    const isApproved = statusLower.includes('approv') || statusLower.includes('complet');
    const isRejected = statusLower.includes('reject') || statusLower.includes('denied');
    const statusSuffix = isApproved ? ' (approved)' : isRejected ? ' (rejected)' : '';

    push('HR Task Type :', d.type || task.name);
    push('eForm List ID :', clean(d.eFormListId));
    push('Task ID :', clean(task.id));
    push('Status :', `${clean(d.status)}${statusSuffix}`);
    endSection();

    push('Requestor :', clean(d.submitter));
    if (clean(d.submittedDate)) push('submitted :', fmtDate(d.submittedDate));
    if (clean(d.section)) push('section :', clean(d.section));
    endSection();

    if (clean(d.missedDate)) push('date :', clean(d.missedDate));
    if (clean(d.timeIn)) push('Time In:', `${clean(d.timeIn)}:00`);
    if (clean(d.timeOut)) push('time out :', `${clean(d.timeOut)}:00`);
    if (clean(d.reason)) push('Reason:', clean(d.reason));
    if (clean(d.fromDate)) push('from date :', clean(d.fromDate));
    if (clean(d.toDate)) push('to date :', clean(d.toDate));
    if (rows.length && rows[rows.length - 1].sectionBreak !== true) endSection();

    if (clean(d.approver1)) {
      push('Approver 1:', clean(d.approver1));
      if (clean(d.approver1Date)) push('Approver 1 Date :', clean(d.approver1Date));
      if (clean(d.approver1Comment)) push('Approver 1 Comment :', clean(d.approver1Comment));
    }
    if (clean(d.approver2)) {
      push('Approver 2:', clean(d.approver2));
      if (clean(d.approver2Date)) push('Approver 2 Date :', clean(d.approver2Date));
      if (clean(d.approver2Comment)) push('Approver 2 Comment :', clean(d.approver2Comment));
    }
    if (clean(d.approver3)) {
      push('Approver 3:', clean(d.approver3));
      if (clean(d.approver3Date)) push('Approver 3 Date :', clean(d.approver3Date));
      if (clean(d.approver3Comment)) push('Approver 3 Comment :', clean(d.approver3Comment));
    }

    if (clean(d.hoursForTransfer)) push('Hours to Transfer :', clean(d.hoursForTransfer));
    if (clean(d.fromYear)) push('From Year :', clean(d.fromYear));
    if (clean(d.toYear)) push('To Year :', clean(d.toYear));
    if (clean(d.needEngineer)) push('Engineer Approval :', clean(d.needEngineer));
    if (clean(d.requestorPin)) push('Requestor PIN :', clean(d.requestorPin));
    if (clean(d.stage)) push('Current Stage :', clean(d.stage));
    if (clean(d.eFormCategory)) push('Category :', clean(d.eFormCategory));
    if (clean(d.eFormProgress)) push('Progress :', clean(d.eFormProgress));
    if (clean(d.taskOutcomeField)) push('Task Outcome :', clean(d.taskOutcomeField));
    if (clean(d.assignedTo2)) push('Assigned To :', clean(d.assignedTo2));
    if (clean(d.assignedToSuperior)) push('Assigned To Superior :', clean(d.assignedToSuperior));
    if (clean(d.assignedToStepNo)) push('Assigned To Step No :', clean(d.assignedToStepNo));
    if (clean(d.assignedToSuperior1)) push('Superior 1 :', clean(d.assignedToSuperior1));
    if (clean(d.assignedToSuperior2field)) push('Superior 2 :', clean(d.assignedToSuperior2field));
    if (clean(d.createdDate)) push('Created :', fmtDate(d.createdDate));
    if (clean(d.modifiedDate)) push('Modified :', fmtDate(d.modifiedDate));

    return rows;
  }

  /** Groups an HR task's fields into the sections shown by the details panel. */
  private buildHrTaskDetailView(d: any, task: any): HrTaskDetailView {
    const clean = (v: any) =>
      !v || String(v).trim() === '' || String(v).trim() === 'Enter value here'
        ? '' : String(v).trim();

    const fmtDate = (val: string) => {
      const dt = new Date(val);
      return isNaN(dt.getTime())
        ? val
        : dt.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
        ' · ' + dt.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    };

    const toneOf = (value: string): HrTaskDetailTone => {
      const v = value.toLowerCase();
      if (v.includes('reject') || v.includes('denied') || v.includes('cancel')) return 'danger';
      if (v.includes('approv') || v.includes('complet') || v.includes('done')) return 'success';
      if (v.includes('pending') || v.includes('progress') || v.includes('waiting')) return 'warning';
      return 'neutral';
    };

    const initialsOf = (name: string) =>
      name.split(/[\s-]+/).filter(Boolean).slice(0, 2).map((p) => p[0].toUpperCase()).join('');

    const me = this.normalizePersonName(this.getLoggedInUserDisplayName());
    const person = (role: string, value: any) => {
      const name = clean(value);
      return name
        ? { role, name, initials: initialsOf(name), isMe: !!me && this.normalizePersonName(name) === me }
        : null;
    };

    const status = clean(d.status);
    const eFormListId = clean(d.eFormListId);
    const taskId = clean(task.id);
    const uploaded = task.submittedDate || task.lastModifiedDateTime;
    const uploadedDate = uploaded ? new Date(uploaded) : null;

    const subtitle = [
      eFormListId && `eForm #${eFormListId}`,
      taskId && `Task #${taskId}`,
      uploadedDate && !isNaN(uploadedDate.getTime()) &&
        `Uploaded ${uploadedDate.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })}`,
    ].filter(Boolean).join(' · ');


    const references = [
      { label: 'eForm List ID', value: eFormListId },
      { label: 'Task ID', value: taskId },
      { label: 'Category', value: clean(d.eFormCategory) || clean(d.category) },
    ].filter((r) => r.value);

    const people = [
      person('Requestor', d.submitter),
      person('Assigned to', d.assignedTo2 || d.assignedTo),
      person('Superior 1', d.assignedToSuperior1),
      person('Superior 2', d.assignedToSuperior2field),
    ].filter((p): p is NonNullable<typeof p> => !!p);

    // Form-specific fields that don't belong in the fixed sections above.
    const details = [
      { label: 'Section', value: clean(d.section) },
      { label: 'Date', value: clean(d.missedDate) },
      { label: 'Time in', value: clean(d.timeIn) && `${clean(d.timeIn)}:00` },
      { label: 'Time out', value: clean(d.timeOut) && `${clean(d.timeOut)}:00` },
      { label: 'Reason', value: clean(d.reason) },
      { label: 'From date', value: clean(d.fromDate) },
      { label: 'To date', value: clean(d.toDate) },
      { label: 'Hours to transfer', value: clean(d.hoursForTransfer) },
      { label: 'From year', value: clean(d.fromYear) },
      { label: 'To year', value: clean(d.toYear) },
      { label: 'Engineer approval', value: clean(d.needEngineer) },
      { label: 'Requestor PIN', value: clean(d.requestorPin) },
      { label: 'Current stage', value: clean(d.stage) },
      { label: 'Assigned to superior', value: clean(d.assignedToSuperior) },
      { label: 'Assigned to step no', value: clean(d.assignedToStepNo) },
    ].filter((r) => r.value);

    const approvals = [1, 2, 3]
      .filter((n) => clean(d[`approver${n}`]))
      .map((n) => ({
        label: `Approver ${n}`,
        name: clean(d[`approver${n}`]),
        date: clean(d[`approver${n}Date`]),
        comment: clean(d[`approver${n}Comment`]),
      }));

    const timeline = [
      { label: 'Submitted', value: clean(d.submittedDate) && fmtDate(d.submittedDate) },
      { label: 'Created', value: clean(d.createdDate) && fmtDate(d.createdDate) },
      { label: 'Last modified', value: clean(d.modifiedDate) && fmtDate(d.modifiedDate) },
    ].filter((t) => t.value);

    return {
      statuses: [{ label: 'Status', value: status, tone: toneOf(status) }],
      title: d.type || task.name,
      subtitle,
      status: status ? { value: status, tone: toneOf(status) } : null,
      references,
      people,
      details,
      approvals,
      timeline,
    };
  }

  private normalizePersonName(value: string): string {
    return String(value || '').toLowerCase().replace(/[^a-z]/g, '');
  }

  protected openHrTaskForm(task: any): void {
    if (!task.id) return;

    try {
      const d = task.eFormDetails;
      const detailRows = this.buildHrTaskDetailRows(d, task);
      const content = detailRows.map((row) => `${row.label}: ${row.value}`).join('\n');

      this.selectedEFormContent = {
        fileName: `${d.type || task.name} - eForm List ID ${d.eFormListId ?? ''}`,
        content,
        contentType: 'text/plain',
        uploadDate: task.submittedDate || task.lastModifiedDateTime,
        eFormDetails: d,
        detailRows,
        detailView: this.buildHrTaskDetailView(d, task),
      };
    } catch (error) {
      this.selectedEFormContent = {
        fileName: task.name,
        content: `Failed to load content: ${(error as any)?.message || 'Unknown error'}`,
        contentType: 'text/plain',
        uploadDate: task.lastModifiedDateTime || task.submittedDate,
      };
    }
    this.isLoadingEFormContent = false;
    this.refreshView();
  }

  protected closeFormViewer(): void { this.selectedEFormContent = null; }

  // ============================================================
  // GRAPH API - thin wrapper used inside this component
  // Services use the standalone graphGet() helper directly.
  // ============================================================
  private graphGet(path: string, token: string, timeoutMs: number = AppConstants.graphDefaultTimeoutMs) {
    return graphGet(this.http, path, token, timeoutMs);
  }

  // ============================================================
  // GET TARGET DRIVE ID
  // ============================================================
  /** Dedupes concurrent drive lookups during login / HR Files open. */
  private driveResolveInFlight: Promise<string | null> | null = null;

  private async getTargetDriveId(token: string): Promise<string | null> {
    if (this.cachedDriveId && this.cachedSiteId) return this.cachedDriveId;
    if (this.driveResolveInFlight) return this.driveResolveInFlight;

    this.driveResolveInFlight = this.resolveTargetDriveId(token);
    try {
      return await this.driveResolveInFlight;
    } finally {
      this.driveResolveInFlight = null;
    }
  }

  private async resolveTargetDriveId(token: string): Promise<string | null> {
    if (this.cachedDriveId && this.cachedSiteId) return this.cachedDriveId;

    const site = await this.siteMetadataService.resolve(token);
    this.cachedSiteId = site.siteId;
    this.cachedSiteWebUrl = site.siteWebUrl;
    if (this.cachedSiteLists.length === 0) this.cachedSiteLists = site.lists;

    const drives: any = await graphGetWithRetry(
      this.http,
      `/sites/${site.siteId}/drives?$select=id,name,webUrl`,
      token,
    );
    const driveList: any[] = drives?.value ?? [];

    const targetDrive = driveList.find(
      (d: any) => normalizeName(d.name) === normalizeName(this.targetLibraryName)
    );
    if (targetDrive?.id) { this.cachedDriveId = targetDrive.id; return targetDrive.id; }

    const documentsDrive = driveList.find((d: any) => {
      const n = normalizeName(d.name);
      return n === 'documents' || n === 'shareddocuments';
    });
    if (documentsDrive?.id) {
      const rootChildren: any = await graphGetWithRetry(
        this.http,
        `/drives/${documentsDrive.id}/root/children?$select=id,name,webUrl,folder&$top=500`,
        token,
        AppConstants.graphFileListingTimeoutMs,
      );
      const targetFolder = (rootChildren?.value ?? []).find(
        (item: any) =>
          !!item?.folder && normalizeName(item.name) === normalizeName(this.targetLibraryName)
      );
      if (targetFolder?.id) { this.cachedDriveId = documentsDrive.id; return documentsDrive.id; }
    }
    return null;
  }

  // ============================================================
  // RESOLVE HR PERSONAL FOLDER NAME (for user card display)
  // ============================================================
  private async resolveHrPersonalFolderName(): Promise<void> {
    if (!this.currentUser) {
      this.hrPersonalFolderName = null;
      return;
    }
    if (this.hrPersonalFolderName) return;

    this.refreshView();
    try {
      const token = await this.getSharePointToken();
      const driveId = await this.getTargetDriveId(token);
      if (!driveId) return;

      const userFolder = await this.findUserFolder(driveId, token);
      if (userFolder?.name) {
        this.hrPersonalFolderName = userFolder.name;
        if (!this.userRootFolderId) {
          this.userRootFolderId = userFolder.id;
        }
      }
    } catch { /* non-blocking - card falls back to placeholder text */ }
    finally {
      this.refreshView();
    }
  }

  // ============================================================
  // FIND USER FOLDER (4-level cache hierarchy)
  // ============================================================
  private async findUserFolder(
    driveId: string, token: string
  ): Promise<{ id: string; name: string; webUrl?: string } | null> {
    const candidates = this.getUserFolderCandidates();
    if (!candidates.length) return null;

    const memKey = `${driveId}-root`;
    const lsKey = `${AppConstants.folderCacheLsPrefix}${driveId}_${this.currentUser?.userPrincipalName ?? ''}`;

    try {
      const stored = localStorage.getItem(lsKey);
      if (stored) {
        const parsed = JSON.parse(stored) as { id: string; name: string; webUrl: string; cachedAt: number };
        if (parsed?.id && (Date.now() - parsed.cachedAt) < AppConstants.folderCacheTtlMs) {
          this.folderCache.set(memKey, parsed);
          return { id: parsed.id, name: parsed.name, webUrl: parsed.webUrl };
        }
      }
    } catch { /* ignore corrupt cache */ }

    const memCached = this.folderCache.get(memKey);
    if (memCached) return memCached;

    const directResult = await this.findUserFolderDirect(driveId, token, candidates);
    if (directResult) {
      this.folderCache.set(memKey, directResult);
      try {
        localStorage.setItem(lsKey, JSON.stringify({
          ...directResult, webUrl: directResult.webUrl ?? '', cachedAt: Date.now(),
        }));
      } catch { /* ignore corrupt cache */ }
      return directResult;
    }

    // Do NOT deep-walk the drive (was up to ~1500 children Graph calls). Direct path +
    // root listing above is enough for the card; My Files can resolve later on demand.
    return null;
  }

  private clearLocalStorageFolderCache(): void {
    try {
      Object.keys(localStorage)
        .filter(k => k.startsWith(AppConstants.folderCacheLsPrefix))
        .forEach(k => localStorage.removeItem(k));
    } catch { /* ignore */ }
  }

  private async findUserFolderDirect(
    driveId: string,
    token: string,
    candidates: string[]
  ): Promise<{ id: string; name: string; webUrl?: string } | null> {
    const user = this.currentUser;
    if (!user) return null;

    const email = (user.email || user.userPrincipalName || '').trim().toLowerCase();
    const upn = (user.userPrincipalName || '').trim().toLowerCase();
    const employeeId = (user.employeeId || '').trim();

    const exactNames: string[] = [];
    if (employeeId) {
      if (email) exactNames.push(`${employeeId} ${email}`);
      if (upn && upn !== email) exactNames.push(`${employeeId} ${upn}`);
    }
    if (email) exactNames.push(email);
    if (upn && upn !== email) exactNames.push(upn);

    for (const exactName of exactNames) {
      try {
        const folder: any = await graphGetWithRetry(
          this.http,
          `/drives/${driveId}/root:/${encodeURIComponent(exactName)}?$select=id,name,webUrl,folder`,
          token,
          AppConstants.graphFileListingTimeoutMs,
        );
        if (folder?.id && folder?.folder) {
          return { id: folder.id, name: folder.name, webUrl: folder.webUrl };
        }
      } catch { /* try the next candidate */ }
    }

    try {
      const rootChildren: any = await graphGetWithRetry(
        this.http,
        `/drives/${driveId}/root/children?$select=id,name,webUrl,folder&$top=999`,
        token,
        AppConstants.graphFileListingTimeoutMs,
      );
      const folders = (rootChildren?.value ?? []).filter((item: any) => !!item?.folder);

      for (const exactName of exactNames) {
        const match = folders.find(
          (f: any) => (f.name ?? '').toLowerCase() === exactName.toLowerCase()
        );
        if (match) return { id: match.id, name: match.name, webUrl: match.webUrl };
      }

      let bestMatch: { id: string; name: string; webUrl?: string } | null = null;
      let bestScore = 0;
      for (const folder of folders) {
        const score = this.scoreFolderMatch(folder.name ?? '', candidates);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = { id: folder.id, name: folder.name, webUrl: folder.webUrl };
        }
      }
      if (bestMatch && bestScore >= 55) return bestMatch;
    } catch { /* skip */ }
    return null;
  }

  // ============================================================
  // GET ALL DRIVE ITEMS
  // ============================================================
  private async getAllDriveItems(
    driveId: string,
    token: string,
    startFolderId: string = 'root',
    onProgress?: (message: string) => void
  ): Promise<{ items: any[]; hadFailures: boolean; wasTruncated: boolean }> {
    const collected: any[] = [];
    const foldersToVisit: string[] = [startFolderId];
    let hadFailures = false;
    let wasTruncated = false;
    let visitedFolders = 0;
    const startedAtMs = Date.now();

    while (foldersToVisit.length > 0) {
      if (
        visitedFolders >= AppConstants.maxDriveFoldersToScan ||
        collected.length >= AppConstants.maxDriveItemsToCollect ||
        Date.now() - startedAtMs > AppConstants.maxDriveTraversalMs
      ) {
        wasTruncated = hadFailures = true;
        break;
      }

      const batch = foldersToVisit.splice(0, 5);
      this.startUserFileLoadingWatchdog(); // reset once per batch

      const batchPromises = batch.map(async folderId => {
        visitedFolders += 1;
        let nextPath: string | null = folderId === 'root'
          ? `/drives/${driveId}/root/children?$select=id,name,webUrl,lastModifiedDateTime,lastModifiedBy,size,file,folder,parentReference&$top=200`
          : `/drives/${driveId}/items/${folderId}/children?$select=id,name,webUrl,lastModifiedDateTime,lastModifiedBy,size,file,folder,parentReference&$top=200`;
        const folderResults: any[] = [];
        while (nextPath) {
          try {
            const page: any = await graphGetWithRetry(
              this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs,
            );
            const pageItems: any[] = page?.value ?? [];
            folderResults.push(...pageItems);
            for (const item of pageItems) {
              if ((item?.folder?.childCount ?? 0) > 0) foldersToVisit.push(item.id);
            }
            nextPath = toGraphPath(page?.['@odata.nextLink']);
          } catch {
            hadFailures = true;
            break;
          }
        }
        return folderResults;
      });

      try {
        const batchResults = await Promise.all(batchPromises);
        for (const results of batchResults) collected.push(...results);
        onProgress?.(`Loading files... ${collected.length} items found`);
      } catch {
        hadFailures = true;
      }
    }

    return { items: collected, hadFailures, wasTruncated };
  }

  // ============================================================
  // USER FOLDER CANDIDATE NAMES & SCORING
  // ============================================================
  private getUserFolderCandidates(): string[] {
    const user = this.currentUser;
    if (!user) return [];

    const email = (user.email || '').trim().toLowerCase();
    const upn = (user.userPrincipalName || '').trim().toLowerCase();
    const displayName = (user.username || '').trim().toLowerCase();
    const employeeId = (user.employeeId || '').trim().toLowerCase();

    const localFromEmail = email.includes('@') ? email.split('@')[0] : email;
    const localFromUpn = upn.includes('@') ? upn.split('@')[0] : upn;

    const displayParts = displayName.split(/\s+/).filter(Boolean);
    const reversedDisplayName = displayParts.length >= 2
      ? `${displayParts[displayParts.length - 1]} ${displayParts.slice(0, -1).join(' ')}`
      : '';

    const unique = new Set<string>();
    for (const c of [email, upn, localFromEmail, localFromUpn, displayName, reversedDisplayName, employeeId]) {
      if (c) { unique.add(c); unique.add(normalizeName(c)); }
    }
    return Array.from(unique).filter(Boolean);
  }

  private scoreFolderMatch(folderName: string, candidates: string[]): number {
    const nf = normalizeName(folderName);
    const rfl = folderName.toLowerCase();
    let bestScore = 0;

    for (const candidate of candidates) {
      const nc = normalizeName(candidate);
      const rc = candidate.toLowerCase();
      if (!nc && !rc) continue;
      if (nf && nc && nf === nc) bestScore = Math.max(bestScore, 100);
      else if (rfl === rc) bestScore = Math.max(bestScore, 95);
      else if (nf && nc && (nf.startsWith(nc) || nc.startsWith(nf))) bestScore = Math.max(bestScore, 80);
      else if (nf && nc && (nf.includes(nc) || nc.includes(nf))) bestScore = Math.max(bestScore, 70);
      else if (rc && rfl.includes(rc)) bestScore = Math.max(bestScore, 60);
    }
    return bestScore;
  }

  // ============================================================
  // ERROR HANDLER
  // ============================================================
  // the error handler is used to handle the errors that occur when loading the user files
  private handleUserFilesError(error: unknown): void {
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
    this.refreshView();
  }

  // ============================================================
  // WATCHDOG TIMER
  // ============================================================
  //what this does is it will timeout the user file loading if it takes too long
  private startUserFileLoadingWatchdog(): void {
    this.clearUserFileLoadingWatchdog();
    this.userFileLoadingTimeoutId = setTimeout(() => {
      if (this.isLoadingHrFilesList) {
        this.userFileError = 'SharePoint request timed out. Please try again.';
        this.isLoadingHrFilesList = false;
        this.refreshView();
        return;
      }
      if (this.isLoadingUserFiles) {
        this.userFileError = 'SharePoint request timed out. Please try again.';
        this.isLoadingUserFiles = false;
        this.refreshView();
      }
    }, AppConstants.userFileLoadingWatchdogMs);
  }

  private clearUserFileLoadingWatchdog(): void {
    if (this.userFileLoadingTimeoutId !== null) {
      clearTimeout(this.userFileLoadingTimeoutId);
      this.userFileLoadingTimeoutId = null;
    }
  }

  // ============================================================
  // POWERAPPS MODAL
  // ============================================================
  protected showPowerAppsModal = false;

  protected openPowerAppsModal(): void {
    this.showPowerAppsModal = true;
  }

  protected closePowerAppsModal(): void {
    this.showPowerAppsModal = false;
  }

  // ============================================================
  // NEW COMMENT MODAL
  // ============================================================
  protected isCommentPanelOpen = false;
  protected selectedTask: { Id: number | string } | null = null;
  protected selectedCommentListName = '';

  protected get showRequestForActionCategory(): boolean {
    return this.formConfigService.allowsRequestForAction(
      this.selectedEFormListId ?? '',
      this.selectedEFormTitle ?? '',
      this.selectedCommentListName,
    );
  }

  protected openNewCommentModal(): void {
    if (!this.selectedTask?.Id) {
      console.error('Select a task first');
      return;
    }
    this.isCommentPanelOpen = true;
  }

  protected closeNewCommentModal(): void {
    this.isCommentPanelOpen = false;
  }

  /** Marks a Request for Action as completed in SharePoint and removes it from To Do. */
  protected async onCompleteTask(event: { task: any; completionText: string }): Promise<void> {
    const task = event?.task;
    const completionText = String(event?.completionText ?? '').trim();
    const taskId = String(task?.id ?? task?.Id ?? '').trim();
    if (!taskId) return;

    try {
      const completedByName = this.getLoggedInUserDisplayName();
      await this.commentService.completeRequestForAction(task, completionText, completedByName);

      this.todoService.removeTask(taskId);
      // Keep Redis in sync - otherwise the completed task reappears on the next
      // cache paint until a full Graph refresh replaces the snapshot.
      this.syncTaskCachesAfterMutation(task, 'complete');

      const completionNote = completionText
        ? `[Completed by ${completedByName}]: ${completionText}`
        : `[Completed by ${completedByName}]`;

      this.commentItems = this.commentItems.map(item => {
        if (String(item.id ?? '').trim() !== taskId) {
          return item;
        }

        const existingComment = String(
          item.eFormDetails?.['comment'] ?? item.eFormDetails?.['commentHtml'] ?? '',
        )
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<[^>]+>/g, '')
          .trim();
        const mergedComment = existingComment
          ? `${existingComment}\n\n${completionNote}`
          : completionNote;

        return {
          ...item,
          status: '',
          completedBy: completedByName,
          eFormDetails: {
            ...(item.eFormDetails ?? {}),
            status: '',
            completedBy: completedByName,
            comment: mergedComment,
            commentHtml: mergedComment.replace(/\n/g, '<br>'),
          },
        };
      });

      this.allFilesComponent?.refreshCommentSearchCache();
      this.invalidateCommentFilters();
      this.refreshView();
    } catch (err) {
      console.error('Failed to complete Request for Action:', err);
      this.errorMessage = err instanceof Error ? err.message : 'Failed to complete task';
      this.refreshView();
    }
  }

  // to update the comment section iteams for the new-comment modal in the todo section
  //also this is saved by the new-comment modal in the todo section grouping the comments by task id.
  protected onNewCommentSaved(ev: NewCommentSavedEvent): void {
    const taskId = String(ev?.taskId ?? '').trim();
    const base = taskId ? this.resolveCommentCardBase(taskId) : null;
    if (!base) return;

    this.prependCommentItem(buildCommentCardFromSavedEvent(ev, base, taskId, {
      commentAuthor: this.getLoggedInUserDisplayName(),
      fallbackEFormListId: this.selectedEFormListId ?? '',
    }));

    const parentTask =
      this.commentItems.find(item => String(item.id ?? '').trim() === String(ev.taskId ?? '').trim()) ??
      { id: ev.taskId };

    if (ev.category === 'action' && ev.assignedToName) {
      this.onDelegateSuccess(String(ev.taskId), ev.assignedToName, ev.assignedToEmail ?? '');
    } else {
      this.syncTaskCachesAfterMutation(parentTask, 'comment');
    }
  }

  /** After a successful SharePoint due-date write, keep Redis snapshots aligned. */
  protected onTaskDueDateChanged(event: { task: any; dueDate: string }): void {
    const dueDate = String(event?.dueDate ?? '').trim();
    if (!event?.task || !dueDate) return;
    this.syncTaskCachesAfterMutation(event.task, 'dueDate', { dueDate });
  }

  // to resolve the comment card base for the new-comment modal in the todo section
  private resolveCommentCardBase(taskId: string): CommentListItem | null {
    return (
      this.commentItems.find(item => String(item.id ?? '').trim() === taskId) ??
      this.buildCommentCardFromSelection(taskId)
    );
  }

  // to get the task submitter from the comment card base for the new-comment modal in the todo section
  private getTaskSubmitterFromBase(base: CommentListItem): string {
    return (
      base.submittedBy ??
      base.eFormDetails?.submittedBy ??
      base.eFormDetails?.submitter ??
      this.currentUser?.username ??
      ''
    );
  }
  protected onCommentLinkClick(event: Event): void {
    onCommentLinkClick(event);
  }

  private prependCommentItem(item: CommentListItem): void {
    this.commentItems = [item, ...this.commentItems];
    this.invalidateCommentFilters();
    this.refreshView();
  }

  private buildCommentCardFromSelection(taskId: string): (typeof this.commentItems)[0] | null {
    if (!this.selectedTask?.Id || String(this.selectedTask.Id) !== taskId) {
      return null;
    }

    const todoTask = this.todoService.getTaskById(taskId);
    if (todoTask) {
      return {
        id: taskId,
        name: todoTask.name,
        webUrl: String((todoTask as any).webUrl ?? ''),
        isFolder: false,
        submittedBy: todoTask.submittedBy,
        eFormDetails: {
          ...(todoTask.eFormDetails ?? {}),
          eFormListId: this.selectedEFormListId ?? todoTask.eFormDetails?.eFormListId,
          type: this.selectedEFormTitle ?? todoTask.eFormDetails?.type,
          category: todoTask.eFormDetails?.category ?? 'eForm',
          assignedTo: todoTask.eFormDetails?.assignedTo ?? todoTask.assignedTo,
          submittedDate: todoTask.eFormDetails?.submittedDate,
          status: todoTask.eFormDetails?.status ?? '',
        },
        isContentLoaded: true,
      };
    }

    return {
      id: taskId,
      name: this.selectedEFormTitle ?? 'Task',
      webUrl: '',
      isFolder: false,
      submittedBy: this.selectedSubmitter ?? undefined,
      eFormDetails: {
        eFormListId: this.selectedEFormListId,
        type: this.selectedEFormTitle,
        category: 'eForm',
        status: '',
      },
      isContentLoaded: true,
    };
  }

  // ============================================================
  // MOBILE NAVIGATION
  // ============================================================
  protected activeNavIndex: number = MOBILE_NAV_INDEX.todo;

  protected get mobileSectionTitle(): string {
    if (this.showAllFilesSection) return 'All Files';
    if (this.showToDoSection) return 'To Do';
    return 'HR';
  }

  protected get mobileDetailTitle(): string {
    return (
      this.selectedFolderName ||
      this.selectedSubmitter ||
      this.currentFolderName ||
      this.selectedEFormTitle ||
      'Details'
    );
  }

  /** Leave Comments / Attachments and return to the list the user came from. */
  protected closeMobileDetail(): void {
    if (this.showAllFilesSection) {
      this.mobileCardView = 'allFiles';
      this.activeNavIndex = MOBILE_NAV_INDEX.allFiles;
    } else if (this.showToDoSection) {
      this.mobileCardView = 'todo';
      this.activeNavIndex = MOBILE_NAV_INDEX.todo;
    } else {
      this.mobileCardView = 'task';
      this.activeNavIndex = MOBILE_NAV_INDEX.task;
    }
    this.refreshView();
  }

  protected showMobileComments(): void {
    this.hideComments = false;
    this.mobileCardView = 'comments';
    this.syncMobileSearchInput();
    this.refreshView();
  }

  protected showMobileAttachments(): void {
    this.mobileCardView = 'attachments';
    if (
      this.currentLibraryDriveId ||
      this.attachmentBrowsingRoot ||
      this.currentFolderId ||
      (this.showUserFile && this.userFiles.length > 0)
    ) {
      this.showUserFile = true;
      this.refreshView();
      return;
    }
    if (this.currentUser?.email) {
      if (this.hasLoadedPersonalFiles) {
        this.showMyFilesPanel();
      } else {
        this.openMyFiles();
      }
    }
    this.refreshView();
  }

  protected onNavItemClick(index: number): void {
    const onDetail =
      this.mobileCardView === 'comments' || this.mobileCardView === 'attachments';
    if (onDetail && index === this.activeNavIndex) {
      this.closeMobileDetail();
      return;
    }

    const action = resolveMobileNavClick(
      index,
      this.activeNavIndex,
      !!this.currentUser?.email
    );

    switch (action.type) {
      case 'none':
        return;
      case 'goHome':
        this.toggleUserInfo();
        return;
      case 'goTodo':
        this.toggleToDoSection();
        return;
      case 'goAllFiles':
        this.toggleAllFilesSection();
        return;
      case 'goComments':
        this.hideComments = false;
        Object.assign(this, getCommentsTabMobileState());
        if (action.syncSearch) {
          this.syncMobileSearchInput();
        }
        this.refreshView();
        return;
      case 'goAttachments':
        Object.assign(this, getAttachmentsTabMobileState());
        // Already browsing All Files / HR person attachments - only switch the tab.
        // Never call openMyFiles() here: that clears the folder and reloads tasks.
        if (
          this.currentLibraryDriveId ||
          this.attachmentBrowsingRoot ||
          this.currentFolderId ||
          (this.showUserFile && this.userFiles.length > 0)
        ) {
          this.showUserFile = true;
          this.refreshView();
          return;
        }
        if (action.shouldLoadFiles) {
          if (this.hasLoadedPersonalFiles) {
            this.showMyFilesPanel();
          } else {
            this.openMyFiles();
          }
        }
        return;
    }
  }
}