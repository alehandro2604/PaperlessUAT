import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { AppConstants } from '../app.constants';
import { getFileCategory, getFileExtension, getFileIcon } from '../file-utils';
import { graphGetWithRetry, normalizeName, toGraphPath , graphPost} from '../microsoft-graph';
import { AttachmentFile } from '../components/attachments/attachment-list/attachment-list';
import {firstValueFrom} from "rxjs";
import { AppUser } from '../models/app-user.model';

/** A file or folder from a SharePoint document library, as shown in Attachments. */
export interface DriveFile extends AttachmentFile {
  parentId?: string;
  size?: number;
  fileExtension?: string;
  mimeType?: string;
  fileCategory?: string;
}

/** A folder we only need to find / link to. */
export interface DriveFolderRef {
  id: string;
  name: string;
  webUrl?: string;
}

/** The parts of the signed-in user used to find their personal HR folder. */
export type DriveUser = AppUser;

/** Fields we ask Graph for whenever we list a folder. */
const CHILD_FIELDS =
  'id,name,webUrl,lastModifiedDateTime,lastModifiedBy,size,file,folder,parentReference';

/**
 * Talks to Microsoft Graph about SharePoint files and folders.
 *
 * It only fetches and converts data - it never changes what is on screen.
 * The component (app.ts) decides what to do with the results.
 */
@Injectable({ providedIn: 'root' })
export class DriveFilesService {
  private readonly http = inject(HttpClient);

  /** `${driveId}-root` -> the signed-in user's personal folder (memory cache). */
  private readonly userFolderCache = new Map<string, DriveFolderRef>();

  // ============================================================
  // CONVERT
  // ============================================================

  /** Turns one Graph driveItem into the row shape Attachments uses. */
  toDriveFile(item: any, fallbackParentId: string): DriveFile {
    const ext = item.file ? getFileExtension(item.name) : '';
    const category = item.file ? getFileCategory(ext) : '';
    return {
      id: item.id,
      name: item.name,
      webUrl: item.webUrl ?? '',
      parentId: item.parentReference?.id ?? fallbackParentId,
      driveId: item.parentReference?.driveId,
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
      fileIcon: item.file ? getFileIcon(category, ext) : '',
    };
  }

  // ============================================================
  // ONE FOLDER
  // ============================================================

  /** Files and sub-folders directly inside one folder (one level, max 500). */
  async fetchFolderChildren(driveId: string, folderId: string, token: string): Promise<DriveFile[]> {
    const resp: any = await graphGetWithRetry(
      this.http,
      `/drives/${driveId}/items/${folderId}/children?$select=${CHILD_FIELDS}&$top=500`,// Graph limits to 500 items per page, so this is the max we can get in one call
      token,
      AppConstants.graphFileListingTimeoutMs,
    );
    return (resp?.value ?? [])
      .filter((item: any) => !!item.file || !!item.folder)
      .map((item: any) => this.toDriveFile(item, folderId));
  }

//============================================================
//preview file
//============================================================

async getPreviewUrl(driveId: string, fileId: string, token: string): Promise<string | null> {
  const res = await firstValueFrom(
    graphPost(this.http, `/drives/${driveId}/items/${fileId}/preview`,{}, token)
  ) as {getUrl? : string};

  try{
    const url = new URL(res.getUrl ?? '');// res.getUrl is the preview URL returned by Graph
    return url.protocol === 'https:' && url.hostname.endsWith('.sharepoint.com') ? url.href : null;//this is a security check to make sure the url is from sharepoint and not some other site
  }

  catch{
    return null;
  }
}

  // ============================================================
  // WHOLE FOLDER TREE
  // ============================================================

  /**
   * Walks a folder and all its sub-folders (5 folders at a time) and returns
   * every raw Graph item found. Stops early at the limits in AppConstants.
   *
   * @param onBatch    called before each batch (app.ts uses it to reset its timeout watchdog)
   * @param onProgress called after each batch with a "... items found" message
   */
  async getAllDriveItems(
    driveId: string,
    token: string,
    startFolderId = 'root',
    hooks: { onBatch?: () => void; onProgress?: (message: string) => void } = {},
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
      hooks.onBatch?.();

      const batchPromises = batch.map(async folderId => {
        visitedFolders += 1;
        let nextPath: string | null = folderId === 'root'
          ? `/drives/${driveId}/root/children?$select=${CHILD_FIELDS}&$top=200`
          : `/drives/${driveId}/items/${folderId}/children?$select=${CHILD_FIELDS}&$top=200`;
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
        hooks.onProgress?.(`Loading files... ${collected.length} items found`);
      } catch {
        hadFailures = true;
      }
    }

    return { items: collected, hadFailures, wasTruncated };
  }

  // ============================================================
  // THE USER'S PERSONAL FOLDER
  // ============================================================

  /**
   * Finds the signed-in user's personal folder in a drive.
   * Order: localStorage (survives reloads) -> memory -> Graph lookup.
   */
  async findUserFolder(driveId: string, token: string, user: DriveUser | null): Promise<DriveFolderRef | null> {
    if (!user) return null;
    const candidates = this.getUserFolderCandidates(user);
    if (!candidates.length) return null;

    const memKey = `${driveId}-root`;
    const lsKey = `${AppConstants.folderCacheLsPrefix}${driveId}_${user.userPrincipalName ?? ''}`;

    try {
      const stored = localStorage.getItem(lsKey);
      if (stored) {
        const parsed = JSON.parse(stored) as { id: string; name: string; webUrl: string; cachedAt: number };
        if (parsed?.id && (Date.now() - parsed.cachedAt) < AppConstants.folderCacheTtlMs) {
          this.userFolderCache.set(memKey, parsed);
          return { id: parsed.id, name: parsed.name, webUrl: parsed.webUrl };
        }
      }
    } catch { /* ignore corrupt cache */ }

    const memCached = this.userFolderCache.get(memKey);
    if (memCached) return memCached;

    const found = await this.findUserFolderDirect(driveId, token, user, candidates);
    if (found) {
      this.userFolderCache.set(memKey, found);
      try {
        localStorage.setItem(lsKey, JSON.stringify({
          ...found, webUrl: found.webUrl ?? '', cachedAt: Date.now(),
        }));
      } catch { /* storage full / blocked */ }
      return found;
    }

    // Do NOT deep-walk the drive (was up to ~1500 children Graph calls). Direct path +
    // root listing is enough for the card; My Files can resolve later on demand.
    return null;
  }

  /** Forget the remembered personal folder (memory + localStorage), e.g. on sign-out. */
  clearUserFolderCache(): void {
    this.userFolderCache.clear();
    try {
      Object.keys(localStorage)
        .filter(k => k.startsWith(AppConstants.folderCacheLsPrefix))
        .forEach(k => localStorage.removeItem(k));
    } catch { /* ignore */ }
  }

  /** Try exact folder names first, then score every root folder against the user's names. */
  private async findUserFolderDirect(
    driveId: string,
    token: string,
    user: DriveUser,
    candidates: string[],
  ): Promise<DriveFolderRef | null> {
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

      let bestMatch: DriveFolderRef | null = null;
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

  /** Every way the user's folder might be named: email, UPN, local part, display name (both orders), employee id. */
  private getUserFolderCandidates(user: DriveUser): string[] {
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

  /** 0-100: how closely a folder name matches one of the candidate names. */
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
}
