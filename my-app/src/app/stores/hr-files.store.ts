import { Injectable, Injector, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { graphGetWithRetry, clearGraphThrottleCooldown, toGraphPath, normalizeName } from '../microsoft-graph';
import { AuthService } from '../services/auth.service';
import { getCommentsTabMobileState, getTaskTabMobileState } from '../components/mobile navigation/mobile-navigation';
import { FileCrawlCacheService } from '../services/file-crawl.service';
import { SiteMetadataService } from '../services/site-metadata.service';
import { DriveFilesService } from '../services/drive-files.service';
import { sortCommentItemsByDateDesc } from '../utils/comment-utils';
import { hrFolderTaskCacheKey, driveFolderCacheKey } from '../utils/cache-keys';
import { sleep } from '../utils/timing';
import { ViewRefreshService } from './view-refresh.service';
import { UiStore } from './ui.store';
import { SessionStore } from './session.store';
import { CommentsStore } from './comments.store';
import { FolderTasksStore } from './folder-tasks.store';
import { TodoStore } from './todo.store';
import { AttachmentsStore } from './attachments.store';
import { AllFilesStore } from './all-files.store';

/**
 * HR Files: the people folders list, the signed-in user's personal folder ("My Files"), access-group check and background warm-up.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class HrFilesStore {
  private readonly injector = inject(Injector);
  private readonly http = inject(HttpClient);
  private readonly authService = inject(AuthService);
  private readonly fileCrawlCache = inject(FileCrawlCacheService);
  private readonly siteMetadataService = inject(SiteMetadataService);
  private readonly driveFiles = inject(DriveFilesService);

  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get comments(): CommentsStore { return this.injector.get(CommentsStore); }
  private get folderTasks(): FolderTasksStore { return this.injector.get(FolderTasksStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get attachments(): AttachmentsStore { return this.injector.get(AttachmentsStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  /** Persisted HR Files people-folder list so the first open paints from cache. */
  static readonly HR_ROOT_FOLDERS_CACHE_KEY = 'files:hr-root:v1';

  /** True while the HR Files folder list in the first section is loading - separate from attachments. */
  isLoadingHrFilesList = false;

  hrPersonalRootFolders: Array<{
    id: string;
    name: string;
    webUrl: string;
    isFolder: boolean;
  }> = [];

  hrPersonalFolderName: string | null = null;

  readonly targetLibraryName = sharePointConfig.hrPersonalListDisplayName || 'HRPersonal';

  cachedDriveId: string | null = null;

  hrPersonalDriveId: string | null = null;

  toggleUserInfo(): void {
    if (!this.session.isUserLoggedIn) { return; }
    this.attachments.closePreview();
    this.todo.cancelBackgroundTodoLoad();
    this.allFiles.leaveAllFilesTaskView();
    this.ui.showToDoSection = false;
    this.ui.showAllFilesSection = false;
    this.ui.hideComments = false;
    this.todo.selectedSubmitter = null;
    this.todo.selectedEFormListId = null;
    this.todo.selectedEFormTitle = null;
    Object.assign(this.ui, getTaskTabMobileState());
    this.comments.invalidateCommentFilters();
    this.ui.syncMobileSearchInput();
    if (this.hrPersonalRootFolders.length === 0 && !this.isLoadingHrFilesList) {
      this.loadHrFilesInFirstSection();
    } else if (this.hrPersonalRootFolders.length > 0) {
      // Warm Comments cache for visible people while the user browses the list.
      this.folderTasks.scheduleHrFolderTasksPrefetch(this.hrPersonalRootFolders);
    }
    // Instantly restore the last HR person Comments from cache after To Do / All Files.
    this.folderTasks.restoreHrFilesCommentsFromCache();
    this.view.refreshView();
  }

  // LOAD LOGGED-IN USER FILES (main entry point)
  openMyFiles(): void {
    if (!this.session.currentUser) {
      alert('You must log in first to access your profile.');
      this.attachments.userFileError = 'Please log in first.';
      this.attachments.showUserFile = false;
      return;
    }
    if (this.attachments.isLoadingUserFiles) return;

    // Attachments-only entry: do not reload Comments/tasks.
    this.attachments.showUserFile = true;
    this.attachments.isLoadingUserFiles = true;
    this.attachments.currentFolderId = null;
    this.attachments.currentFolderWebUrl = '';
    this.attachments.currentFolderName = '';
    this.attachments.folderStack = [];
    this.attachments.attachmentBrowsingRoot = null;
    this.attachments.currentLibraryDriveId = null;
    this.attachments.currentLibraryName = null;
    this.attachments.userRootFolderId = null;
    this.attachments.userFileProgressMessage = 'Connecting to SharePoint...';
    this.attachments.userFileError = '';
    this.attachments.userFileWarning = '';
    this.attachments.userFiles = [];
    this.attachments._folderMapDirty = true;
    this.attachments.startUserFileLoadingWatchdog();
    this.view.refreshView();

    void this.loadMyFilesFromDrive();
  }

  loadHrFilesInFirstSection(): void {
    if (!this.session.currentUser) {
      this.attachments.userFileError = 'Please log in first.';
      return;
    }
    if (this.isLoadingHrFilesList) return;

    this.todo.cancelBackgroundTodoLoad();
    clearGraphThrottleCooldown();
    if (this.folderTasks.hrTaskPrefetchTimer) {
      clearTimeout(this.folderTasks.hrTaskPrefetchTimer);
      this.folderTasks.hrTaskPrefetchTimer = null;
    }

    this.attachments.userFileError = '';
    this.attachments.userFileWarning = '';

    // Instant paint from cache (memory / previous session) so the first open is not blank.
    const cachedRoots =
      this.fileCrawlCache.get<typeof this.hrPersonalRootFolders>(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY) ??
      this.fileCrawlCache.getStale<typeof this.hrPersonalRootFolders>(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY);
    if (Array.isArray(cachedRoots) && cachedRoots.length > 0) {
      this.hrPersonalRootFolders = cachedRoots;
      this.isLoadingHrFilesList = false;
      this.attachments.userFileProgressMessage = '';
      this.attachments.clearUserFileLoadingWatchdog();
      this.view.refreshView();
      // Top up from SharePoint without wiping the painted list.
      void this.loadHrPersonalFilesForFirstSection({ softRefresh: true });
      return;
    }

    this.isLoadingHrFilesList = true;
    this.attachments.userFileProgressMessage = 'Connecting to SharePoint...';
    this.hrPersonalRootFolders = [];
    this.attachments.startUserFileLoadingWatchdog();
    this.view.refreshView();

    void this.loadHrPersonalFilesForFirstSection();
  }

  showMyFilesPanel(): void {
    this.attachments.showUserFile = true;
    this.view.refreshView();
  }

  refreshHrSection(): void {
    if (this.allFiles.selectedFolderName) {
      this.folderTasks.invalidateFolderTaskCache(this.targetLibraryName, this.allFiles.selectedFolderName);
    }

    this.fileCrawlCache.invalidate(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY);
    this.hrPersonalRootFolders = [];
    this.loadHrFilesInFirstSection();

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
    if (this.hrPersonalDriveId && this.attachments.currentLibraryDriveId === this.hrPersonalDriveId && folderId) {
      this.fileCrawlCache.invalidate(driveFolderCacheKey(this.hrPersonalDriveId, folderId));
      if (this.allFiles.selectedFolderName) {
        this.folderTasks.loadAllFilesTasksForLibrary(this.targetLibraryName, this.allFiles.selectedFolderName);
      }
      void this.attachments.loadDriveFolderContents(
        this.hrPersonalDriveId,
        folderId,
        folderWebUrl,
        this.attachments.currentFolderName || this.attachments.attachmentBrowsingRoot?.name || stackFolder?.name,
        { skipCache: true },
      );
    } else if (
      this.attachments.showUserFile &&
      this.attachments.userRootFolderId &&
      !this.attachments.currentLibraryDriveId &&
      !this.attachments.isLoadingUserFiles
    ) {
      this.attachments.isLoadingUserFiles = true;
      this.attachments.userFileProgressMessage = 'Refreshing attachments...';
      this.attachments.userFileError = '';
      this.attachments.userFileWarning = '';
      this.attachments.startUserFileLoadingWatchdog();
      this.view.refreshView();
      void this.loadMyFilesFromDrive();
    }

    this.view.refreshView();
  }

  async loadMyFilesFromDrive(): Promise<void> {
    // This load is slow. If the user opens a To Do task (or a library folder) before it
    // finishes, that view owns the Attachments panel and this result must be dropped.
    const todoSeq = this.attachments.todoAttachmentLoadSeq;
    const isSuperseded = () =>
      this.attachments.todoAttachmentLoadSeq !== todoSeq || !!this.attachments.currentLibraryDriveId;
    try {
      const token = await this.session.getSharePointToken();
      const targetDriveId = await this.getTargetDriveId(token);

      if (isSuperseded()) return;

      if (!targetDriveId) {
        const site = await this.siteMetadataService.resolve(token);
        const drives: any = await graphGetWithRetry(
          this.http,
          `/sites/${site.siteId}/drives?$select=id,name,webUrl`,
          token,
        );
        const available = drives?.value?.map((x: any) => x.name).join(', ') || 'none';
        this.attachments.userFileError = `Could not find "${this.targetLibraryName}". Available libraries: ${available}`;
        this.view.refreshView();
        return;
      }

      await this.loadCurrentUserPersonalFolderFiles(targetDriveId, token, isSuperseded);
    } catch (error: unknown) {
      if (!isSuperseded()) this.attachments.handleUserFilesError(error);
    } finally {
      if (isSuperseded()) {
        // Stop our watchdog so it can't flag the newer view as timed out.
        this.attachments.clearUserFileLoadingWatchdog();
      } else if (this.attachments.isLoadingUserFiles) {
        this.attachments.isLoadingUserFiles = false;
        this.attachments.clearUserFileLoadingWatchdog();
        this.view.refreshView();
      }
    }
  }

  async loadHrPersonalFilesForFirstSection(
    options: { softRefresh?: boolean } = {},
  ): Promise<void> {
    const softRefresh = !!options.softRefresh;
    try {
      const token = await this.session.getSharePointToken();
      if (!softRefresh) {
        this.attachments.userFileProgressMessage = 'Loading HR Files...';
        this.view.refreshView();
      }

      // Drive resolve (Graph) and group membership (SharePoint REST) run in parallel -
      // the old serial path waited for both before the first children page could start.
      const [targetDriveId, allAccess] = await Promise.all([
        this.getTargetDriveId(token),
        this.isUserInHrPersonalAllAccessGroup(),
      ]);

      if (!targetDriveId) {
        if (!softRefresh) {
          this.attachments.userFileError = `Could not find "${this.targetLibraryName}".`;
          this.view.refreshView();
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
              this.fileCrawlCache.set(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY, partial);
              this.isLoadingHrFilesList = false;
              this.attachments.userFileProgressMessage = '';
              this.attachments.clearUserFileLoadingWatchdog();
              this.view.refreshView();
              this.folderTasks.scheduleHrFolderTasksPrefetch(partial);
            },
          );
          this.fileCrawlCache.set(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY, this.hrPersonalRootFolders);
          this.folderTasks.scheduleHrFolderTasksPrefetch(this.hrPersonalRootFolders);
          return;
        } catch (err: any) {
          const status = err?.status ?? err?.error?.status;
          if (status !== 403 && status !== 401) throw err;
          // Not actually allowed to list root - fall through to personal folder.
        }
      }

      const userFolder = await this.driveFiles.findUserFolder(targetDriveId, token, this.session.currentUser);
      if (!userFolder?.id) {
        if (!softRefresh) {
          this.attachments.userFileError =
            `Could not find a personal folder for ${this.session.currentUser?.username || 'the logged-in user'} in "${this.targetLibraryName}".`;
          this.view.refreshView();
        }
        return;
      }

      this.hrPersonalFolderName = userFolder.name;
      this.hrPersonalRootFolders = [{ ...userFolder, webUrl: userFolder.webUrl ?? '', isFolder: true }];
      this.fileCrawlCache.set(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY, this.hrPersonalRootFolders);
      this.folderTasks.scheduleHrFolderTasksPrefetch(this.hrPersonalRootFolders);
    } catch (error: unknown) {
      if (!softRefresh) {
        this.attachments.handleUserFilesError(error);
      }
    } finally {
      if (this.isLoadingHrFilesList) {
        this.isLoadingHrFilesList = false;
        this.attachments.clearUserFileLoadingWatchdog();
        this.view.refreshView();
      }
    }
  }

  hrPersonalGroupChecked = false;

  hrPersonalGroupAllowsAllAccess = false;

  readonly hrPersonalAllAccessGroupId = 3;

  async isUserInHrPersonalAllAccessGroup(): Promise<boolean> {
    if (this.hrPersonalGroupChecked) return this.hrPersonalGroupAllowsAllAccess;
    this.hrPersonalGroupChecked = true;
    this.hrPersonalGroupAllowsAllAccess = false;

    const email = (this.session.currentUser?.email ?? '').trim().toLowerCase();
    const upn = (this.session.currentUser?.userPrincipalName ?? '').trim().toLowerCase();
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
  async warmHrPersonalRootFoldersCache(): Promise<void> {
    if (!this.session.currentUser) return;
    if (this.isLoadingHrFilesList) return;

    try {
      if (!this.fileCrawlCache.get(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY)) {
        const token = await this.session.getSharePointToken();
        const targetDriveId = await this.getTargetDriveId(token);
        if (!targetDriveId) return;
        this.hrPersonalDriveId = targetDriveId;

        if (!(await this.isUserInHrPersonalAllAccessGroup())) {
          const userFolder = await this.driveFiles.findUserFolder(targetDriveId, token, this.session.currentUser);
          if (!userFolder?.id) return;
          this.hrPersonalFolderName = userFolder.name;
          const single = [{ ...userFolder, webUrl: userFolder.webUrl ?? '', isFolder: true as const }];
          this.fileCrawlCache.set(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY, single);
        } else {
          // Background warm: big pages and one cache save at the end. Saving the growing
          // list after every 50-folder page cost a Graph call plus a Redis PUT (each with
          // its CORS preflight) per page.
          const folders = await this.fetchHrPersonalRootFolders(targetDriveId, token, undefined, 999);
          this.fileCrawlCache.set(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY, folders);
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
  scheduleLoginHrTaskWarm(): void {
    const roots =
      this.fileCrawlCache.get<typeof this.hrPersonalRootFolders>(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY) ??
      this.fileCrawlCache.getStale<typeof this.hrPersonalRootFolders>(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY);
    if (!Array.isArray(roots) || roots.length === 0) return;

    const names = roots
      .map(f => String(f?.name ?? '').trim())
      .filter(Boolean)
      .slice(0, AppConstants.hrFilesTaskPrefetchMaxBatch);
    if (names.length === 0) return;

    if (this.folderTasks.hrTaskPrefetchTimer) {
      clearTimeout(this.folderTasks.hrTaskPrefetchTimer);
    }
    this.folderTasks.hrTaskPrefetchTimer = setTimeout(() => {
      this.folderTasks.hrTaskPrefetchTimer = null;
      if (this.attachments.folderPrefetchPaused) {
        // To Do still owns Graph — retry once after the usual pause window.
        this.folderTasks.hrTaskPrefetchTimer = setTimeout(() => {
          this.folderTasks.hrTaskPrefetchTimer = null;
          if (this.attachments.folderPrefetchPaused) return;
          if (this.comments.isLoadingComments || this.comments.isLoadingMoreComments) return;
          void this.folderTasks.prefetchHrFolderTasks(names);
        }, 8000);
        return;
      }
      if (this.comments.isLoadingComments || this.comments.isLoadingMoreComments) return;
      void this.folderTasks.prefetchHrFolderTasks(names);
    }, AppConstants.hrFilesTaskPrefetchDelayMs);

    if (AppConstants.hrFilesWarmAllEnabled) {
      void this.warmAllHrPeopleSlowly();
    }
  }

  /**
   * HR all-access users: trickle-warm Comments for every person into shared Redis,
   * one at a time, so first opens are almost always cache hits for everyone.
   * Yields to any user-driven load and skips people cached within the TTL.
   */
  async warmAllHrPeopleSlowly(): Promise<void> {
    if (this.folderTasks.hrWarmAllRunning) return;
    this.folderTasks.hrWarmAllRunning = true;
    try {
      if (!(await this.isUserInHrPersonalAllAccessGroup())) return;
      // Let the first-batch prefetch and To Do get Graph first.
      await sleep(AppConstants.hrFilesWarmAllStartDelayMs);

      const roots =
        this.fileCrawlCache.getStale<typeof this.hrPersonalRootFolders>(HrFilesStore.HR_ROOT_FOLDERS_CACHE_KEY) ?? [];
      // Shuffle so several HR browsers spread across different people.
      const names = roots
        .map(f => String(f?.name ?? '').trim())
        .filter(Boolean)
        .sort(() => Math.random() - 0.5);

      for (const name of names) {
        if (!this.session.currentUser) return;
        if (this.fileCrawlCache.get(hrFolderTaskCacheKey(name))) continue;
        while (
          this.attachments.folderPrefetchPaused ||
          this.comments.isLoadingComments ||
          this.comments.isLoadingMoreComments ||
          this.todo.isLoadingTodoTasks ||
          this.isLoadingHrFilesList ||
          !this.session.backgroundGraphAllowed()
        ) {
          await sleep(5000);
          if (!this.session.currentUser) return;
        }
        await this.folderTasks.prefetchHrFolderTasks([name]);
        await sleep(AppConstants.hrFilesWarmAllGapMs);
      }
    } catch {
      // Best-effort warm — clicks still load normally.
    } finally {
      this.folderTasks.hrWarmAllRunning = false;
    }
  }

  async fetchHrPersonalRootFolders(
    driveId: string,
    token: string,
    onPage?: (
      folders: Array<{ id: string; name: string; webUrl: string; isFolder: boolean }>
    ) => void,
    pageSize = 50,
  ): Promise<Array<{ id: string; name: string; webUrl: string; isFolder: boolean }>> {
    const folders: Array<{ id: string; name: string; webUrl: string; isFolder: boolean }> = [];
    // Default small pages so the first HR Files paint does not wait for the full library
    // root; the background warm has nobody waiting and asks for big pages instead.
    let nextPath: string | null =
      `/drives/${driveId}/root/children?$select=id,name,webUrl,folder&$top=${pageSize}`;

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

  async loadCurrentUserPersonalFolderFiles(
    targetDriveId: string,
    token: string,
    isSuperseded: () => boolean = () => false,
  ): Promise<void> {
    this.attachments.userFileProgressMessage = 'Locating your Personal folder/Files...';
    this.view.refreshView();

    const userFolder = await this.driveFiles.findUserFolder(targetDriveId, token, this.session.currentUser);
    if (isSuperseded()) return;
    if (!userFolder?.id) {
      this.attachments.userFileError =
        `Could not find a personal folder for ${this.session.currentUser?.username || 'the logged-in user'} in "${this.targetLibraryName}".`;
      this.view.refreshView();
      return;
    }

    this.attachments.userRootFolderId = userFolder.id;
    this.hrPersonalFolderName = userFolder.name;

    const driveLoadResult = await this.driveFiles.getAllDriveItems(targetDriveId, token, userFolder.id, {
      onBatch: () => {
        if (!isSuperseded()) this.attachments.startUserFileLoadingWatchdog();
      },
      onProgress: progress => {
        if (isSuperseded()) return;
        this.attachments.userFileProgressMessage = progress;
        this.view.refreshView();
      },
    });

    // A To Do task or library folder was opened meanwhile - don't overwrite its files.
    if (isSuperseded()) return;
    this.attachments.applyDriveItems(driveLoadResult.items, userFolder.id);

    if (driveLoadResult.wasTruncated) {
      this.attachments.userFileWarning = 'SharePoint scan limit reached; showing partial results.';
    }
  }

  //Folder Name LOOKUP(runs at login,lightweight lookup)
  async loadHrPersonalFolderName(): Promise<void> {
    if (!this.session.currentUser) return;
    this.view.refreshView();
    try {
      const token = await this.session.getSharePointToken();
      const targetDriveId = await this.getTargetDriveId(token);
      if (!targetDriveId) return;
      const userFolder = await this.driveFiles.findUserFolder(targetDriveId, token, this.session.currentUser);
      if (!userFolder?.id) return;
      this.hrPersonalFolderName = userFolder.name;
    } catch {
      //silent ignore -- card will show fallback text
    } finally {
      this.view.refreshView();
    }
  }

  openHrFileListItem(item: { id: string; name: string; webUrl: string; isFolder: boolean }): void {
    if (item.isFolder) {
      void this.openHrFileListItemAsync(item);
      return;
    }

    if (item.webUrl) {
      window.open(item.webUrl, '_blank', 'noopener,noreferrer');
    }
  }

  async openHrFileListItemAsync(item: { id: string; name: string; webUrl: string; isFolder: boolean }): Promise<void> {
    if (!(await this.ensureHrPersonalDriveId())) {
      this.attachments.userFileError = 'HRPersonal library is not loaded yet. Wait for HR Files to finish loading, then try again.';
      this.view.refreshView();
      return;
    }

    // Free Graph capacity for this person's click-path load.
    this.todo.cancelBackgroundTodoLoad();
    clearGraphThrottleCooldown();
    this.attachments.pauseFolderPrefetch();

    this.attachments.showUserFile = true;
    this.ui.hideComments = false;
    this.ui.showToDoSection = false;
    this.ui.showAllFilesSection = false;
    this.attachments.currentLibraryDriveId = this.hrPersonalDriveId;
    this.attachments.currentLibraryName = this.targetLibraryName;
    this.attachments.attachmentBrowsingRoot = { id: item.id, name: item.name, webUrl: item.webUrl };
    this.attachments.currentFolderId = item.id;
    this.attachments.currentFolderWebUrl = item.webUrl;
    this.attachments.currentFolderName = item.name;
    this.attachments.folderStack = [];
    this.attachments.userRootFolderId = null;
    this.allFiles.selectedFolderName = item.name;
    this.folderTasks.selectedHrPersonalTaskFolder = item.name;
    this.folderTasks.lastHrFilesCommentsFolder = item.name;

    // Same-session instant return: paint memory/stale BEFORE any async work so
    // navigating away and back never flashes the loading state for cached people.
    const driveId = this.hrPersonalDriveId!;
    const filesKey = driveFolderCacheKey(driveId, item.id);
    const cachedFiles =
      this.fileCrawlCache.get<any[]>(filesKey) ??
      this.fileCrawlCache.getStale<any[]>(filesKey);
    if (Array.isArray(cachedFiles)) {
      this.attachments.applyDriveFolderContentsToUi(cachedFiles, item.id, item.webUrl, item.name);
    } else {
      this.attachments.userFiles = [];
      this.attachments._folderMapDirty = true;
      this.attachments.isLoadingUserFiles = true;
      this.attachments.userFileProgressMessage = 'Loading folder contents...';
    }

    const tasksKey = hrFolderTaskCacheKey(item.name);
    const cachedTasks =
      this.fileCrawlCache.get<any[]>(tasksKey) ??
      this.fileCrawlCache.getStale<any[]>(tasksKey);
    // Allow empty arrays — a completed crawl with 0 tasks is still a valid hit.
    if (Array.isArray(cachedTasks)) {
      this.comments.commentItems = sortCommentItemsByDateDesc(
        cachedTasks,
      );
      this.comments.commentsMessage = cachedTasks.length ? '' : 'No tasks found.';
      this.comments.isLoadingComments = false;
      // Soft-refresh / deep heal sets Loading more — don't imply the cache is final.
      this.comments.isLoadingMoreComments = false;
    } else {
      this.comments.commentItems = [];
      this.comments.commentsMessage = '';
    }
    this.comments.invalidateCommentFilters();
    // Mobile opens on Comments; Attachments are one tab away and keep loading in the background.
    Object.assign(this.ui, getCommentsTabMobileState());
    this.view.refreshView();
    void this.attachments.loadDriveFolderContents(driveId, item.id, item.webUrl, item.name);
    this.folderTasks.loadAllFilesTasksForLibrary(this.targetLibraryName, item.name);
  }

  async ensureHrPersonalDriveId(): Promise<boolean> {
    if (this.hrPersonalDriveId) return true;

    if (
      this.attachments.currentLibraryDriveId &&
      normalizeName(this.attachments.currentLibraryName ?? '') === normalizeName(this.targetLibraryName)
    ) {
      this.hrPersonalDriveId = this.attachments.currentLibraryDriveId;
      return true;
    }

    try {
      const token = await this.session.getSharePointToken();
      const driveId = await this.getTargetDriveId(token);
      if (!driveId) return false;
      this.hrPersonalDriveId = driveId;
      return true;
    } catch {
      return false;
    }
  }

  /** Dedupes concurrent drive lookups during login / HR Files open. */
  driveResolveInFlight: Promise<string | null> | null = null;

  async getTargetDriveId(token: string): Promise<string | null> {
    if (this.cachedDriveId && this.session.cachedSiteId) return this.cachedDriveId;
    if (this.driveResolveInFlight) return this.driveResolveInFlight;

    this.driveResolveInFlight = this.resolveTargetDriveId(token);
    try {
      return await this.driveResolveInFlight;
    } finally {
      this.driveResolveInFlight = null;
    }
  }

  async resolveTargetDriveId(token: string): Promise<string | null> {
    if (this.cachedDriveId && this.session.cachedSiteId) return this.cachedDriveId;

    const site = await this.siteMetadataService.resolve(token);

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
}
