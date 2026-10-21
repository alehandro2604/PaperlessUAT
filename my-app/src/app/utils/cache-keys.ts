// Cache keys used with FileCrawlCacheService (memory + IndexedDB + shared Redis).
// One place for the key formats, so readers and writers always agree.

export function normalizeCacheKeyPart(value: string): string {
  return String(value ?? '').trim().toLowerCase();
}

export function hrFolderTaskCacheKey(folderName: string): string {
  // v9: only finished crawls are persisted — drops older snapshots that may be partial.
  return `tasks:hr-folder:v9:${normalizeCacheKeyPart(folderName)}`;
}

/** Shared-Redis `{ email, lookupId }` for an HR person folder (under the shared hr-folder prefix). */
export function hrPersonIdsCacheKey(folderName: string): string {
  return `tasks:hr-folder:ids:${normalizeCacheKeyPart(folderName)}`;
}

export function libFolderTaskCacheKey(listName: string, folderName: string): string {
  return `tasks:lib:v7:${normalizeCacheKeyPart(listName)}:${normalizeCacheKeyPart(folderName)}`;
}

export function driveFolderCacheKey(driveId: string, folderId: string): string {
  return `files:folder:${driveId}:${folderId}`;
}

export function childSubjectCacheKey(libraryName: string, folderEFormId: string): string {
  return `${normalizeCacheKeyPart(libraryName)}:${folderEFormId}`;
}
