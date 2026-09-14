// ============================================================
// HR TASK FOLDER MATCHING
// Decides whether a SharePoint task row (raw or already mapped)
// belongs to the HR Files / All Files folder being viewed.
// Extracted from app.ts - none of this reads component state;
// the folder name and match hints arrive as arguments.
// ============================================================
import { Injectable } from '@angular/core';
import { normalizeName } from '../microsoft-graph';
import {
  extractTrailingFolderId, getHrTaskPersonLookupIds, doesHrTaskMatchPersonLookupId,
} from '../utils/hr-task-matching';
import {
  getHrPersonalFolderMatchTokens, candidatesMatchFolderPerson, isItemSubmittedByFolderPerson,
  normalizeTaskMatchText, stringifyTaskFieldValue, fieldMatchesFolderToken,
} from '../utils/person-name-matching';
import { HrTaskMapperService } from './hr-task-mapper.service';

@Injectable({ providedIn: 'root' })
export class HrTaskFolderMatchService {
  constructor(private readonly hrTaskMapper: HrTaskMapperService) { }
  /** True when a SharePoint list item belongs to the clicked document-library folder. */
  doesSharePointTaskMatchFolder(item: any, folderName: string): boolean {
    const folderTokens = getHrPersonalFolderMatchTokens(folderName);
    if (folderTokens.length === 0) return false;

    const f = item.fields ?? item;
    const folderEFormId = extractTrailingFolderId(folderName);
    if (folderEFormId) {
      const taskTitle = String(f.Title ?? f.Name ?? '').trim();
      const taskEFormId = this.extractTaskEFormListId(f);
      const taskTitleId = extractTrailingFolderId(taskTitle);

      if (taskEFormId === folderEFormId) return true;
      if (taskTitleId === folderEFormId) return true;

      // Child rows often reuse the parent folder Title but carry a different eFormListId.
      if (taskEFormId && taskEFormId !== folderEFormId) return false;
      if (taskTitleId && taskTitleId !== folderEFormId) return false;

      // Legacy rows without numeric ids - exact folder title only.
      if (normalizeName(taskTitle) === normalizeName(folderName)) return true;
      return false;
    }

    const candidates = [
      f.Title,
      f.Name,
      f.Folder,
      f.FolderName,
      f.DocumentFolder,
      f.RelatedFolder,
      f.FileLeafRef,
      f.FileDirRef,
      f.Path,
      f.Requestor,
      f.SubmittedBy,
      f.Author,
      f.CreatedBy,
      f.AssignedTo,
      f.Assigned,
      f.AssignedToSuperior1,
      f.AssignedToSuperior2,
      f.EmployeeName,
      f.Employee,
      f.Email,
      f.ADmail,
      f.Comment,
      f.Notes,
      f.ResolvedSubmitter,
    ];

    for (const value of candidates) {
      if (value == null || value === '') continue;
      const text = stringifyTaskFieldValue(value);
      const normalizedText = normalizeTaskMatchText(text);
      if (folderTokens.some(token => normalizedText.includes(token))) return true;

      const segments = text.split(/[/\\]/).map((s: string) => normalizeTaskMatchText(s));
      if (segments.some(segment => folderTokens.some(token => segment.includes(token)))) return true;
    }

    return false;
  }


  private extractTaskEFormListId(fields: Record<string, unknown>): string {
    const field11 = String(fields['field_11'] ?? '').trim();
    return String(
      fields['eFormListId'] ?? fields['ListId'] ?? (/^\d+$/.test(field11) ? field11 : '')
    ).trim();
  }

  /** Final guard for mapped Comments cards in All Files folder view. */
  doesMappedHrTaskMatchFolder(item: any, folderName: string): boolean {
    const folderEFormId = extractTrailingFolderId(folderName);
    if (!folderEFormId) return true;

    const rawFields = item?.eFormDetails?.rawFields;
    if (rawFields && typeof rawFields === 'object') {
      return this.doesSharePointTaskMatchFolder({ fields: rawFields }, folderName);
    }

    const taskEFormId = String(item?.eFormDetails?.eFormListId ?? '').trim();
    const taskTitle = String(item?.name ?? item?.eFormDetails?.type ?? '').trim();
    const taskTitleId = extractTrailingFolderId(taskTitle);

    if (taskEFormId === folderEFormId) return true;
    if (taskTitleId === folderEFormId) return true;
    if (taskEFormId && taskEFormId !== folderEFormId) return false;
    if (taskTitleId && taskTitleId !== folderEFormId) return false;
    return normalizeName(taskTitle) === normalizeName(folderName);
  }

  /**
   * Seed/supplement membership for an HR person folder.
   * When the row has person LookupIds, trust those over loose name text so
   * "Adrian Kind" does not pick up Adriana / Adrian-John rows from newest-page scans.
   */
  doesHrTaskBelongToFolderPerson(
    item: any,
    listName: string,
    folderName: string,
    personLookupId: string | null,
    folderMatchHints: string[],
  ): boolean {
    const lookupIds = getHrTaskPersonLookupIds(item);
    if (personLookupId && lookupIds.length > 0) {
      return lookupIds.includes(personLookupId);
    }
    return (
      doesHrTaskMatchPersonLookupId(item, personLookupId) ||
      this.doesHrTaskMatchFolderPerson(item, listName, folderName) ||
      this.doesHrTaskMatchFolderHints(item, listName, folderMatchHints)
    );
  }

  /** Same rules after mapSharePointItemToHrTask (uses mapped + rawFields). */
  doesMappedHrTaskBelongToFolderPerson(
    item: any,
    folderName: string,
    personLookupId: string | null,
    folderMatchHints: string[],
  ): boolean {
    const raw = item?.eFormDetails?.rawFields;
    const lookupSource = raw && typeof raw === 'object' ? { fields: raw } : item;
    const lookupIds = getHrTaskPersonLookupIds(lookupSource);
    if (personLookupId && lookupIds.length > 0) {
      return lookupIds.includes(personLookupId);
    }
    return (
      doesHrTaskMatchPersonLookupId(lookupSource, personLookupId) ||
      isItemSubmittedByFolderPerson(item, folderName) ||
      this.isItemSubmittedByFolderHints(item, folderMatchHints)
    );
  }

  private doesHrTaskMatchFolderPerson(item: any, listName: string, folderName: string): boolean {
    const f = item.fields ?? item;
    const submitter = this.hrTaskMapper.resolveTaskSubmitter(item, f, listName);
    return candidatesMatchFolderPerson([
      submitter,
      f.Requestor,
      f.SubmittedBy,
      f.Submitter,
      f.Author,
      f.CreatedBy,
      f.AssignedTo,
      f.Assigned,
      f.ResolvedSubmitter,
      f.commentSubmittedBy,
      f.EmployeeName,
      f.Employee,
      f.EmployeeEmail,
      f.RequestorEmail,
      item?.createdBy?.user?.displayName,
      item?.createdBy?.user?.email,
      item?.createdBy?.user?.userPrincipalName,
    ], folderName);
  }

  /** Extra match path when folder email was resolved from AllDomainUsers (no email in folder name). */
  private doesHrTaskMatchFolderHints(item: any, listName: string, hints: string[]): boolean {
    if (!hints.length) return false;
    const f = item.fields ?? item;
    const submitter = this.hrTaskMapper.resolveTaskSubmitter(item, f, listName);
    return this.candidatesMatchFolderHints([
      submitter,
      f.Requestor,
      f.SubmittedBy,
      f.Submitter,
      f.Author,
      f.CreatedBy,
      f.ResolvedSubmitter,
      f.commentSubmittedBy,
      f.EmployeeName,
      f.Employee,
      f.EmployeeEmail,
      f.RequestorEmail,
      item?.createdBy?.user?.displayName,
      item?.createdBy?.user?.email,
      item?.createdBy?.user?.userPrincipalName,
    ], hints);
  }

  private isItemSubmittedByFolderHints(
    item: { submittedBy?: string; eFormDetails?: Record<string, unknown> },
    hints: string[],
  ): boolean {
    if (!hints.length) return false;
    const rawFields = (item.eFormDetails?.['rawFields'] ?? {}) as Record<string, unknown>;
    return this.candidatesMatchFolderHints([
      item.submittedBy,
      item.eFormDetails?.['submitter'],
      item.eFormDetails?.['submittedBy'],
      item.eFormDetails?.['commentSubmittedBy'],
      rawFields['Requestor'],
      rawFields['SubmittedBy'],
      rawFields['Submitter'],
      rawFields['Author'],
      rawFields['CreatedBy'],
      rawFields['commentSubmittedBy'],
      rawFields['EmployeeName'],
      rawFields['Employee'],
      rawFields['EmployeeEmail'],
    ], hints);
  }

  private candidatesMatchFolderHints(candidates: unknown[], hints: string[]): boolean {
    const normalizedHints = hints
      .map(hint => normalizeTaskMatchText(hint))
      .filter(hint => hint.length >= 3);
    if (normalizedHints.length === 0) return false;

    const normalizedCandidates = candidates
      .map(value => normalizeTaskMatchText(stringifyTaskFieldValue(value)))
      .filter(Boolean);
    if (normalizedCandidates.length === 0) return false;

    return normalizedHints.some(hint =>
      normalizedCandidates.some(text => fieldMatchesFolderToken(text, hint))
    );
  }

}
