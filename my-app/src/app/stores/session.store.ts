import { Injectable, Injector, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { graphGet, isGraphBackgroundBackedOff } from '../microsoft-graph';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { AuthService } from '../services/auth.service';
import { FormConfigurationService } from '../services/form-configuration.service';
import { UserService } from '../services/user.service';
import { DelegateService } from '../services/delegate.service';
import { FileCrawlCacheService } from '../services/file-crawl.service';
import { TodoTaskStoreService } from '../services/todo-task-store.service';
import { SiteMetadataService } from '../services/site-metadata.service';
import { BackendApiService } from '../services/backend-api.service';
import { SharePointTaskQueryService } from '../services/sharepoint-task-query.service';
import { DriveFilesService } from '../services/drive-files.service';
import { TodoFetchService } from '../services/todo-fetch.service';
import { UiStore } from './ui.store';
import { CommentsStore } from './comments.store';
import { FolderTasksStore } from './folder-tasks.store';
import { TodoStore } from './todo.store';
import { AttachmentsStore } from './attachments.store';
import { HrFilesStore } from './hr-files.store';
import { AllFilesStore } from './all-files.store';

/**
 * The signed-in user: sign-in on start, profile, sign-out, Graph tokens, site info, and the background poll that keeps caches fresh.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class SessionStore {
  private readonly injector = inject(Injector);
  private readonly http = inject(HttpClient);
  private readonly userService = inject(UserService);
  private readonly authService = inject(AuthService);
  private readonly delegateService = inject(DelegateService);
  private readonly formConfigService = inject(FormConfigurationService);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly todoStore = inject(TodoTaskStoreService);
  private readonly siteMetadataService = inject(SiteMetadataService);
  private readonly backendApi = inject(BackendApiService);
  private readonly taskQuery = inject(SharePointTaskQueryService);
  private readonly driveFiles = inject(DriveFilesService);
  private readonly todoFetch = inject(TodoFetchService);

  // Other features
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get comments(): CommentsStore { return this.injector.get(CommentsStore); }
  private get folderTasks(): FolderTasksStore { return this.injector.get(FolderTasksStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get attachments(): AttachmentsStore { return this.injector.get(AttachmentsStore); }
  private get hr(): HrFilesStore { return this.injector.get(HrFilesStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  /**
   * Background poll for To Do watermarks. Folder Comments uses a longer cadence
   * (FOLDER_COMMENTS_POLL_MS) so a person with many lists is not re-fanned every tick.
   */
  static readonly SHAREPOINT_CACHE_POLL_MS = 10 * 60_000;

  /**
   * Idle top-up of the open HR/All Files Comments folder. Each run re-crawls every task
   * list for that folder, so keep it rare — every 3 min helped push idle tabs into 429s.
   */
  static readonly FOLDER_COMMENTS_POLL_MS = 20 * 60_000;

  /**
   * Returning to the browser tab tops up To Do only if the last top-up is older than
   * this. Each top-up reads one page from every To Do list (~19 Graph calls), and
   * switching tabs back and forth used to fire one every time, even mid-search.
   */
  static readonly TAB_RETURN_TODO_TOPUP_MIN_GAP_MS = 3 * 60_000;

  /** When the last To Do top-up started (ms), for the tab-return gap above. */
  lastTodoTopUpAtMs = 0;

  currentUser: {
    email: string;
    username: string;
    userPrincipalName: string;
    employeeId?: string;
  } | null = null;

  get isUserLoggedIn(): boolean { return !!this.currentUser; }

  // Site info lives in SiteMetadataService; these just read it (no second copy).
  get cachedSiteId(): string | null { return this.siteMetadataService.current?.siteId ?? null; }

  get cachedSiteWebUrl(): string { return this.siteMetadataService.current?.siteWebUrl ?? ''; }

  get cachedSiteLists(): any[] { return this.siteMetadataService.current?.lists ?? []; }

  /** Interval that picks up external SharePoint task edits into Redis while signed in. */
  sharePointCachePollTimer: ReturnType<typeof setInterval> | null = null;

  lastFolderCommentsPollMs = 0;

  /** Last mouse / keyboard / touch / scroll input — background Graph work stops when stale. */
  lastUserActivityMs = Date.now();

  readonly onUserActivity = (): void => {
    this.lastUserActivityMs = Date.now();
  };

  /**
   * Gate for background (not user-clicked) Graph traffic: polling, prefetch, warm-up.
   * Off while the tab is hidden, after a few idle minutes, and shortly after any 429 —
   * so a tab left open does not keep spending the tenant's SharePoint budget.
   */
  backgroundGraphAllowed(): boolean {
    if (typeof document !== 'undefined' && document.visibilityState === 'hidden') return false;
    if (Date.now() - this.lastUserActivityMs > AppConstants.backgroundIdleAfterMs) return false;
    return !isGraphBackgroundBackedOff();
  }

  readonly onDocumentVisibilityForCachePoll = (): void => {
    if (document.visibilityState === 'visible') {
      this.startSharePointCachePoll();
      void this.pollSharePointTaskCaches({ fromTabReturn: true });
    } else {
      this.stopSharePointCachePollTimerOnly();
    }
  };

  async checkLoginOnStart(): Promise<void> {
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
        this.ui.statusMessage = `Signed in as ${this.currentUser?.username || 'User'}`;
        this.authService.restorePostLoginPathIfNeeded();
        return;
      }
      await this.trySilentSignIn();
    } catch {
      await this.beginSignInRedirect();
    }
  }

  async trySilentSignIn(): Promise<void> {
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
      this.ui.statusMessage = `Signed in as ${this.currentUser?.username || 'User'}`;
    } catch (error) {
      const prompt = error instanceof InteractionRequiredAuthError ? 'login' : 'none';
      await this.beginSignInRedirect(prompt);
    }
  }

  async beginSignInRedirect(prompt: 'none' | 'login' = 'none'): Promise<void> {
    if (this.ui.isConnecting) return;
    this.ui.errorMessage = '';
    this.ui.isConnecting = true;
    this.ui.statusMessage = 'Redirecting to Microsoft sign-in...';
    try {
      await this.authService.loginRedirect(prompt);
    } catch {
      this.ui.isConnecting = false;
      this.ui.statusMessage = 'Could not start Microsoft sign-in. Please refresh the page.';
    }
  }

  async loadCurrentUserProfile(token: string): Promise<void> {
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
          this.todo.initializeToDoOnLogin();

          // Defer HR warm + delegates so they don't race the first To Do Graph wave.
          // People-list warm fills shared Redis early so the first HR Files open is fast.
          setTimeout(() => {
            void this.hr.loadHrPersonalFolderName();
            void this.delegateService.loadDelegates();
            void this.hr.warmHrPersonalRootFoldersCache();
          }, 750);
          resolve();
        },
        error: () => {
          this.currentUser = null;
          this.hr.hrPersonalFolderName = null;
          this.userService.setCurrentUser(null);
          void this.delegateService.loadDelegates();
          resolve();
        },
      });
    });
  }

  clearUserScopedSharePointState(): void {
    this.attachments.showUserFile = false;
    this.attachments.userFiles = [];
    this.hr.hrPersonalRootFolders = [];
    this.attachments.userRootFolderId = null;
    this.attachments.currentFolderId = null;
    this.attachments.currentFolderWebUrl = '';
    this.attachments.currentFolderName = '';
    this.attachments.folderStack = [];
    this.attachments.attachmentBrowsingRoot = null;
    this.attachments.currentLibraryDriveId = null;
    this.attachments.currentLibraryName = null;
    this.hr.hrPersonalDriveId = null;
    this.hr.hrPersonalFolderName = null;
    this.allFiles.selectedFolderName = '';
    this.allFiles.selectedAllFilesParentFolderName = '';
    this.allFiles.selectedAllFilesFolderId = null;
    this.allFiles.childSubjectSourceFetchedForFolderId = null;
    this.allFiles.childSubjectNamesByLibraryEFormId.clear();
    this.folderTasks.selectedHrPersonalTaskFolder = '';
    this.comments.commentItems = [];
    this.comments.commentsMessage = '';
    this.todo.userHrTasksLoaded = false;
    this.todo.todoFirstPassComplete = false;
    this.todo.todoProcurementDrainActive = false;
    this.todo.todoScopeCacheValid = false;
    this.todoStore.invalidate();
    this.attachments.userFileError = '';
    this.attachments.userFileWarning = '';
    this.attachments.userFileProgressMessage = '';
    this.attachments._folderMapDirty = true;
    this.driveFiles.clearUserFolderCache();
    this.comments.invalidateCommentFilters();
  }

  logout(): void {
    this.hr.hrPersonalFolderName = null;
    this.stopSharePointCachePoll();
    // Wipe memory + IndexedDB so a shared machine keeps nothing after sign-out.
    this.fileCrawlCache.clearAll();
    this.driveFiles.clearUserFolderCache();
    this.siteMetadataService.reset();
    this.hr.cachedDriveId = null;
    this.hr.driveResolveInFlight = null;
    this.taskQuery.reset();
    this.todo.todoFirstPassComplete = false;
    this.todo.todoProcurementDrainActive = false;
    this.todo.todoScopeCacheValid = false;
    this.todoFetch.todoListCursors.clear();
    this.todoFetch.todoProcScanStoppedEarly.clear();
    this.todoFetch.todoListWatermarks.clear();
    this.siteMetadataService.reset();
    this.authService.logoutRedirect();
    this.hr.hrPersonalFolderName = null;
  }

  async getSharePointToken(): Promise<string> {
    return this.authService.acquireSharePointToken();
  }

  /**
   * While signed in, periodically top up To Do Redis from SharePoint watermarks and
   * soft-refresh the open folder Comments cache so external SharePoint edits appear.
   */
  startSharePointCachePoll(): void {
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', this.onDocumentVisibilityForCachePoll);
      document.addEventListener('visibilitychange', this.onDocumentVisibilityForCachePoll);
      for (const type of ['pointerdown', 'keydown', 'wheel', 'touchstart'] as const) {
        document.removeEventListener(type, this.onUserActivity, true);
        document.addEventListener(type, this.onUserActivity, { capture: true, passive: true });
      }
    }
    if (typeof document !== 'undefined' && document.visibilityState === 'hidden') {
      return;
    }
    this.stopSharePointCachePollTimerOnly();
    this.sharePointCachePollTimer = setInterval(() => {
      void this.pollSharePointTaskCaches();
    }, SessionStore.SHAREPOINT_CACHE_POLL_MS);
  }

  stopSharePointCachePollTimerOnly(): void {
    if (this.sharePointCachePollTimer) {
      clearInterval(this.sharePointCachePollTimer);
      this.sharePointCachePollTimer = null;
    }
  }

  stopSharePointCachePoll(): void {
    this.stopSharePointCachePollTimerOnly();
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', this.onDocumentVisibilityForCachePoll);
      for (const type of ['pointerdown', 'keydown', 'wheel', 'touchstart'] as const) {
        document.removeEventListener(type, this.onUserActivity, true);
      }
    }
  }

  async pollSharePointTaskCaches(options: { fromTabReturn?: boolean } = {}): Promise<void> {
    if (!this.currentUser?.email) return;
    if (!this.backgroundGraphAllowed()) return;
    if (this.todo.superiorMode) return;
    // Do not compete with an active To Do / Comments Graph fan-out.
    if (
      this.todo.isLoadingTodoTasks ||
      this.todo.todoProcurementDrainActive ||
      this.comments.isLoadingComments ||
      this.comments.isLoadingMoreComments ||
      this.folderTasks.folderTaskGraphInFlight
    ) {
      return;
    }

    const recentlyToppedUp =
      options.fromTabReturn &&
      Date.now() - this.lastTodoTopUpAtMs < SessionStore.TAB_RETURN_TODO_TOPUP_MIN_GAP_MS;
    if (this.todoFetch.todoListWatermarks.size > 0 && !recentlyToppedUp) {
      await this.todo.refreshTodoTasksIncrementally();
    }

    const now = Date.now();
    if (now - this.lastFolderCommentsPollMs < SessionStore.FOLDER_COMMENTS_POLL_MS) {
      return;
    }
    this.lastFolderCommentsPollMs = now;
    this.folderTasks.softRefreshOpenFolderTaskCache({ onlyIfStale: true });
  }

  async ensureFormConfigLoaded(): Promise<void> {
    if (this.formConfigService.snapshot.length > 0) {
      return;
    }
    await firstValueFrom(this.formConfigService.load());
  }

  /** Fire-and-forget invalidate of the Node `/api/tasks` Redis read-through entry. */
  invalidateBackendTasksCache(task: any): void {
    const list = String(task?.listName ?? task?.eFormDetails?.listName ?? '').trim();
    const email = String(
      task?.userEmail ??
      task?.eFormDetails?.submittedByEmail ??
      this.currentUser?.email ??
      '',
    ).trim();
    if (!list || !email) return;
    void this.backendApi.invalidateTasksForPerson(email, list).catch(() => { });
  }

  async ensureSiteMetadata(
    _siteHost: string,
    _sitePath: string,
    token: string,
  ): Promise<void> {
    if (!this.cachedSiteId || this.cachedSiteLists.length === 0) {
      const metadata = await this.siteMetadataService.resolve(token);
      this.taskQuery.setSiteContext(metadata.siteId, metadata.lists);
    }
    // Other users' filter results, before any list is probed (once per session; cheap after).
    await this.taskQuery.loadSharedListFilterStatus();
  }

  getLoggedInUserDisplayName(): string {
    return this.userService.getCurrentUserName();
  }

  // GRAPH API - thin wrapper used inside this component
  // Services use the standalone graphGet() helper directly.
  graphGet(path: string, token: string, timeoutMs: number = AppConstants.graphDefaultTimeoutMs) {
    return graphGet(this.http, path, token, timeoutMs);
  }
}
