// ============================================================
// APP CONSTANTS
// Centralised home for every magic number / key that was
// previously scattered as private static readonly members
// on the App component class.
// Usage:  import { AppConstants } from './app.constants';
// ============================================================

export const AppConstants = {
  // Session-storage key used to restore the pre-login URL
  postLoginPathKey: 'paperless_post_login_path',

  // localStorage prefix for per-drive folder-location cache
  folderCacheLsPrefix: 'sp_folder_',
  /** Persists which SharePoint list columns accept Graph $filter (avoids repeat 400 probes). */
  listFilterFieldLsPrefix: 'sp_list_filter_v2_',

  // How long (ms) a cached folder ID is considered fresh  (30 days)
  folderCacheTtlMs: 30 * 24 * 60 * 60 * 1000,// day,hour,minute,second,millisecond

  // Auto-refresh was removed: refreshing is user-driven via the toolbar refresh button,
  // which fetches only rows newer than what is already cached (see todoRefreshMaxPagesPerList).

  /**
   * UPDATED LOADING:
   * Graph list-item date filters / $orderby are rejected on Live SharePoint lists.
   * My Tasks prefers AssignedToLookupId filtering (cheap — follow all matching pages).
   * Procurement lists use Progress=Pending (also follow all pages of that filtered set).
   * Unfiltered page-scans stay hard-capped so blocked lists cannot flood Graph.
   */
  hrTasksFastLoadPageSize: 200,
  hrTasksFastLoadPageLimit: 20, // fallback pages when assignee LookupId filter is unavailable
  /** To Do page-scan fallback when AssignedToLookupId is not filterable. */
  hrTasksTodoLoadPageLimit: 5,
  /** To Do page-scan for procurement lists when Progress filter fails (keep capped). */
  hrTasksTodoProcPageLimit: 10,
  /** Parallel list fetches for To Do (keep low — shared Graph concurrency is 2). */
  hrTasksTodoListConcurrency: 2,
  /**
   * AssignedToLookupId filter pages per list (0 = all matching assignee rows).
   * Filtered queries are small — do not cap or Live-assigned tasks go missing.
   */
  hrTasksTodoLookupMaxPages: 0,
  /**
   * Unfiltered / cursor scan pages when LookupId is blocked (not used for proc Progress path).
   */
  todoInitialPagesPerList: 5,
  /**
   * Procurement Progress=Pending first-slice pages. Remaining pages still drain before
   * To Do clears "Loading more…" so older assigned Proc* rows are not left behind.
   */
  todoProcInitialPages: 30,
  /** Background drain chunk size for remaining Progress=Pending procurement cursors. */
  todoProcBackgroundPagesPerChunk: 40,
  /** Unused on login (auto-drain removed); kept for rare manual drain callers. */
  todoDrainPagesPerList: 0,
  /** Manual "Load older" continues from stored cursors without re-scanning from page 1. */
  todoLoadMorePagesPerList: 50,
  /**
   * Safety cap for the manual refresh. Refresh walks each list newest-first and stops at
   * the first row it has already ingested, so it normally reads a single page; this only
   * bounds the case where a list changed more in one sitting than a few pages can hold.
   */
  todoRefreshMaxPagesPerList: 5,
  hrTasksSubordinateLoadPageLimit: 5,
  /** Max parallel Graph list-item fetches when opening an HR Files person (avoids 429 throttling). */
  hrFilesTaskListConcurrency: 2,
  /**
   * First pass pages per list only when LookupId cannot be used.
   * Keep small — HR Files fans out across many lists; deep scans freeze the UI.
   */
  hrFilesTaskListPageLimit: 2,
  /**
   * After LookupId hits, scan newest page(s) so brand-new rows are not missed.
   * Keep small — these are unfiltered company-wide pages; LookupId is the real
   * source for people with hundreds of tasks.
   */
  hrFilesLookupSupplementPages: 2,
  hrFilesSourceEFormLookupSupplementPages: 2,
  /** Extra pages for Sick Certificate / source eForm lists when LookupId misses. */
  hrFilesSourceEFormScanPages: 6,
  /** Graph $top for person LookupId queries (larger pages = fewer round-trips). */
  hrFilesPersonLookupPageSize: 200,
  /**
   * Soft-refresh / poll: only the newest LookupId pages. lastModifiedDateTime desc
   * catches both new rows and updates to older ones without re-reading 800+ tasks.
   */
  hrFilesSoftRefreshLookupPages: 3,
  /**
   * Cold open first paint: newest-first LookupId pages only, then clear the spinner.
   * Deeper pages + assignee continue silently in the background.
   */
  hrFilesFirstPaintLookupPages: 3,
  /** One newest unfiltered page on soft refresh — catches assignee-only rows LookupId misses. */
  hrFilesSoftRefreshSupplementPages: 1,
  /** After the first person column returns hits, extra columns only need a top-up. */
  hrFilesLookupExtraFieldPages: 2,
  /**
   * Max pages for HR Files person LookupId filter (Requestor/Author).
   * Higher cap so "new since last visit" rows (e.g. yesterday) are not skipped.
   */
  hrFilesPersonLookupMaxPages: 50,
  /**
   * Extra pages for related-step fallback scan. Keep 0 for speed — eFormId
   * filters usually find siblings; raise only if related steps are missing.
   */
  hrFilesTaskListBackfillPageLimit: 0,
  /**
   * Idle-prefetch HR Files person tasks (like All Files folder prefetch).
   * Keep moderate — each person fans out across many task lists.
   */
  hrFilesTaskPrefetchMaxBatch: 12,
  hrFilesTaskPrefetchConcurrency: 2,
  hrFilesTaskPrefetchDelayMs: 500,
  hrFilesTaskPrefetchBatchGapMs: 400,

  /**
   * All Files document-library folder Comments only (not HRPersonal).
   * Progressive scan: show matches early, keep paging so Live lists with many
   * items still find folder tasks that sit past the first few pages.
   */
  docLibraryFolderTaskPageSize: 200,
  /** Soft cap for All Files folder task full-scan fallback (filters should find matches first). */
  docLibraryFolderTaskMaxPages: 5,

  /** UI recency window used for local visibility once tasks are loaded. */
  hrTasksRecentFetchDays: 30, //30 days instead of 30 days
  /** Comments list UI: hide items older than this unless "Show all history" is on. */
  hrTasksVisibleDays: 30, //30 days instead of 30 days

  // Graph API request timeouts
  graphDefaultTimeoutMs: 15_000,
  graphFileListingTimeoutMs: 60_000,

  // Hard cap on the user-file loading watchdog
  userFileLoadingWatchdogMs: 120_000,

  // Drive-traversal safety caps
  maxDriveFoldersToScan: 1200,
  maxDriveItemsToCollect: 5000,
  maxDriveTraversalMs: 90_000,

  // SharePoint folder attachment list (see filteredAttachments in app.ts)
  recentAttachmentVisibilityDays: 30,
  /** How many older files each "Load more" reveals (not days). */
  olderAttachmentsPageSize: 50,
} as const;