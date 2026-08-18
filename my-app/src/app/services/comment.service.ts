import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { BehaviorSubject, Observable, from } from 'rxjs';
import { firstValueFrom } from 'rxjs';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { normalizeSharePointFileUrl } from '../file-utils';
import { graphGet, graphPatch, graphPost, graphPut, withTimeout } from '../microsoft-graph';
import { AuthService } from './auth.service';
import { SiteMetadataService } from './site-metadata.service';

/**
 * Shape of a task passed into createCommentTask(). We only need the item id, the
 * list it lives in, and the raw SharePoint fields so we can detect the correct
 * assignee column and its current value.
 */
export interface CommentTask {
  id?: string;
  parentTaskId?: string;
  listName?: string;
  eFormListId?: string;
  eFormTitle?: string;
  commentText?: string;
  category?: string;
  submittedByName?: string;
  attachmentFile?: File | null;
  attachmentFileName?: string;
  attachmentUrl?: string;
  parentRawFields?: Record<string, unknown>;
  eFormDetails?: {
    rawFields?: Record<string, unknown>;
    assignedTo?: string;
    listName?: string;
  };
  createdDateTime?: string;
  fields?: Record<string, unknown>;
  listWebUrl?: string;
  siteWebUrl?: string;
  userEmail?: string;
  userUpn?: string;
  isComment?: boolean;
  isAttachment?: boolean;
  isTask?: boolean;
  isCommentTask?: boolean;
  assignedToUser?: {
    Title: string;
    Email: string;
    UserPrincipalName: string;
  } | null;
}

// Shape of a list column
interface ListColumn {
  name: string;
  displayName: string;
  readOnly?: boolean;
  required?: boolean;
  choices?: string[];
}

const STATUS_FIELD_CANDIDATES = [
  'LastState',
  'Status',
  'ApprovalStatus',
  'WorkflowStatus',
  'TaskOutcome',
  'Outcome',
  'CurrentStage',
  'Stage',
  'field_3',
  'field_4',
  'field_5',
  'Completed',
  'IsCompleted',
  'State',
  'Progress',
] as const;

const DEFAULT_COMPLETION_STATUS_VALUES = [
  'Completed',
  'Complete',
  'Approved',
  'Done',
  'Closed',
  'Finished',
  'Resolved',
] as const;

// Shape of a comment field map
interface CommentFieldMap {
  category: any;
  comment?: string;
  eFormListId?: string;
  submitter?: string;
  attachmentUrl?: string;
}

// Comment service
@Injectable({
  providedIn: 'root',
})
// Comment service class
export class CommentService {
  private readonly commentTasksSubject = new BehaviorSubject<CommentTask[]>([]);// subject to store the comment tasks
  private readonly loadingSubject = new BehaviorSubject<boolean>(false);// subject to store the loading state
  private readonly errorSubject = new BehaviorSubject<string | null>(null);// subject to store the error state

  constructor(
    private readonly http: HttpClient,// http client to make the API calls
    private readonly authService: AuthService,// auth service to acquire the SharePoint token
    private readonly siteMetadata: SiteMetadataService,// shared site + list metadata
  ) {}

  createCommentTask(task: CommentTask): Observable<CommentTask> {
    return from(this.createCommentTaskInternal(task));// return an observable to create the comment task
  }

  // private method to create the comment task
  private async createCommentTaskInternal(task: CommentTask): Promise<CommentTask> {
    this.loadingSubject.next(true);
    this.errorSubject.next(null);

    try {
      return await withTimeout(
        this.saveCommentToSharePoint(task),
        task.isAttachment && task.attachmentFile ? 120_000 : 45_000,
        'Timed out while saving comment to SharePoint.',
      );
    } catch (err: unknown) {
      const message = this.formatError(err);
      this.errorSubject.next(message);
      throw new Error(message);
    } finally {
      this.loadingSubject.next(false);
    }
  }


  async completeRequestForAction(
    task: {
      id?: string;
      listName?: string;
      eFormDetails?: { listName?: string; rawFields?: Record<string, unknown>; comment?: string };
    },
    completionText: string,
    completedByName: string,
  ): Promise<Record<string, unknown>> {
    return this.completeRequestForActionInternal(task, completionText, completedByName);
  }

  private async completeRequestForActionInternal(
    task: {
      id?: string;
      listName?: string;
      eFormDetails?: { listName?: string; rawFields?: Record<string, unknown>; comment?: string };
    },
    completionText: string,
    completedByName: string,
  ): Promise<Record<string, unknown>> {
    const itemId = String(task?.id ?? '').trim();
    const listName = String(task.listName ?? task.eFormDetails?.listName ?? '').trim();

    if (!itemId) {
      throw new Error('Cannot complete: task has no id');
    }
    if (!listName) {
      throw new Error('Cannot complete: task has no list name');
    }

    const token = await withTimeout(
      this.acquireTokenForCommentWrite(),
      20_000,
      'Timed out while acquiring SharePoint token.',
    );

    const siteId = sharePointConfig.siteId?.trim();
    if (!siteId) {
      throw new Error('No site ID available');
    }

    const listId = await this.resolveListId(siteId, listName, token);
    const columns = await this.loadListColumns(siteId, listId, token);
    const fieldMap = this.resolveCommentFieldMap(columns);
    const writable = new Set(columns.filter(col => !col.readOnly).map(col => col.name));

    let saved = await this.readItemFields(siteId, listId, itemId, token);
    const statusColumn = this.resolveStatusColumnForItem(columns, saved);
    const statusColumnMeta = statusColumn
      ? columns.find(col => col.name === statusColumn)
      : undefined;

    if (statusColumn && statusColumnMeta?.readOnly) {
      throw new Error(
        `Cannot complete: the Status field ("${statusColumnMeta.displayName || statusColumn}") is read-only in SharePoint. ` +
          'It may be managed by a workflow — ask your SharePoint admin how Request for Action items should be completed.',
      );
    }

    if (!statusColumn || !writable.has(statusColumn)) {
      throw new Error(
        `Cannot complete: list "${listName}" has no writable Status column, so SharePoint status could not be updated.`,
      );
    }

    const currentStatus = String(saved[statusColumn] ?? '').trim();
    if (this.isStatusComplete(currentStatus)) {
      console.info('[CommentService] Request for Action already complete in SharePoint', {
        listName,
        itemId,
        statusColumn,
        currentStatus,
      });
    } else {
      const statusValues = this.resolveCompletionStatusValues(statusColumn, columns, currentStatus);
      let statusUpdated = false;
      let lastError: unknown = null;

      for (const value of statusValues) {
        if (value.toLowerCase() === currentStatus.toLowerCase()) {
          continue;
        }

        try {
          await firstValueFrom(
            graphPatch(
              this.http,
              `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
              { [statusColumn]: value },
              token,
            ),
          );

          const verified = await this.readItemFields(siteId, listId, itemId, token);
          const newStatus = String(verified[statusColumn] ?? '').trim();
          if (
            this.isStatusComplete(newStatus) &&
            newStatus.toLowerCase() !== currentStatus.toLowerCase()
          ) {
            saved = verified;
            statusUpdated = true;
            break;
          }
        } catch (err) {
          lastError = err;
          const unknownField = this.extractUnknownField(err);
          if (unknownField) {
            continue;
          }
        }
      }

      if (!statusUpdated) {
        const allowed = statusValues.join(', ') || 'none detected';
        throw new Error(
          `Could not update SharePoint status from "${currentStatus || 'empty'}". ` +
            `Tried: ${allowed}. ` +
            (this.formatError(lastError) || 'The list may use status values not supported by the app.'),
        );
      }
    }

    const existingComment = this.readCommentFromFields(saved, fieldMap);
    const updatedComment = this.buildCompletionComment(existingComment, completionText, completedByName);
    const commentPatch = this.buildCommentPatch(fieldMap, updatedComment, writable);

    if (Object.keys(commentPatch).length) {
      try {
        await firstValueFrom(
          graphPatch(
            this.http,
            `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
            commentPatch,
            token,
          ),
        );
      } catch (err) {
        console.warn('[CommentService] Status updated but comment patch failed', err);
      }
    }

    const verified = await this.readItemFields(siteId, listId, itemId, token);
    const finalStatus = String(verified[statusColumn] ?? '').trim();
    if (!this.isStatusComplete(finalStatus)) {
      throw new Error(
        `SharePoint status is still "${finalStatus || 'empty'}" after completion. ` +
          'The status field may have been reset by a SharePoint workflow.',
      );
    }

    console.info('[CommentService] Request for Action marked complete in SharePoint', {
      listName,
      itemId,
      statusColumn,
      status: finalStatus,
    });
    return verified;
  }

  // private method to save the comment to SharePoint
  private async saveCommentToSharePoint(task: CommentTask): Promise<CommentTask> {
    const token = await withTimeout(
      this.acquireTokenForCommentWrite(),
      20_000,
      'Timed out while acquiring SharePoint token.',
    );

    const siteId = sharePointConfig.siteId?.trim();
    if (!siteId) {
      throw new Error('No site ID available');
    }

    const listName = task.listName?.trim();
    if (!listName) {
      throw new Error('No list name available. Select a Level 2 task group first.');
    }

    const eFormListId = String(task.eFormListId ?? '').trim();
    if (!eFormListId) {
      throw new Error('No eForm list ID available. Select a Level 2 task group first.');
    }

    const commentTextInput = String(task.commentText ?? '').trim();
    if (!commentTextInput && !task.attachmentFile) {
      throw new Error('Comment text is required');
    }

    const listId = await this.resolveListId(siteId, listName, token);
    const columns = await this.loadListColumns(siteId, listId, token);
    const fieldMap = this.resolveCommentFieldMap(columns);
    if (!fieldMap.comment && !fieldMap.eFormListId) {
      throw new Error(
        `List "${listName}" has no writable Comment or eFormListId columns. ` +
          'Cannot save a visible comment row.',
      );
    }
    const title = this.resolveCommentTitle(task.category, task.isAttachment === true);

    let attachmentUrl = String(task.attachmentUrl ?? '').trim();
    let attachmentFileName = String(task.attachmentFileName ?? '').trim();
    let commentText = commentTextInput;

    const isAttachmentSave =
      task.isAttachment === true || task.category === 'attachment' || !!task.attachmentFile;

    if (isAttachmentSave && task.attachmentFile) {
      const uploaded = await this.uploadAttachmentWithFallback(
        siteId,
        task.attachmentFile,
        eFormListId,
        token,
      );
      attachmentUrl = uploaded.webUrl;
      attachmentFileName = uploaded.name;
      commentText = this.buildAttachmentCommentHtml(commentTextInput, attachmentFileName, attachmentUrl);
    }

    const parentFields = await this.fetchParentFields(
      siteId,
      listId,
      String(task.parentTaskId ?? '').trim(),
      token,
      task.parentRawFields ?? task.eFormDetails?.rawFields,
    );

    const attempts = this.buildCreateAttempts(
      title,
      commentText,
      eFormListId,
      String(task.submittedByName ?? '').trim(),
      attachmentUrl,
      attachmentFileName,
      title,
      fieldMap,
      columns,
      parentFields,
    );

    const response = await this.postListItemAttempts(
      siteId,
      listId,
      listName,
      attempts,
      token,
      fieldMap,
      eFormListId,
      commentText,
    );

    return {
      ...task,
      id: response.id,
      fields: response.fields,
      listWebUrl: response.webUrl,
      attachmentUrl,
      attachmentFileName,
    };
  }

  // private method to build the attachment comment HTML
  //
  // IMPORTANT: `attachmentUrl` arriving here is already percent-encoded
  // (SharePoint's REST API returns ServerRelativeUrl/webUrl with spaces as
  // %20, etc.). Do NOT run it through encodeURI()/encodeURIComponent() again —
  // that double-encodes it (%20 -> %2520) and breaks the link. Only HTML-escape
  // it so it's safe to sit inside an href="..." attribute.
  private buildAttachmentCommentHtml(comment: string, fileName: string, attachmentUrl: string): string {
    const safeComment = this.escapeHtml(comment.trim()).replace(/\n/g, '<br>');
    const safeName = this.escapeHtml(fileName);
    const safeUrl = this.escapeHtml(normalizeSharePointFileUrl(attachmentUrl));
    const link =
      `<a class="comment-link" href="${safeUrl}" target="_blank" rel="noopener noreferrer">${safeName}</a>`;
    return safeComment ? `${safeComment}<br><br>${link}` : link;
  }

  // private method to escape the HTML
  /** Escape plain text so it is safe to embed inside HTML (e.g. comment body with a link). */
  private escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')   // &  → &amp;  (must be first, or other replacements would double-encode)
      .replace(/</g, '&lt;')    // <  → &lt;   (prevents browser treating user text as an HTML tag)
      .replace(/>/g, '&gt;')    // >  → &gt;   (closes any accidental tag-like text)
      .replace(/"/g, '&quot;')  // "  → &quot; (safe inside href="..." or other attributes)
      .replace(/'/g, '&#39;');  // '  → &#39;  (safe inside single-quoted attributes)
  }

  private async resolveSiteDriveId(siteId: string, token: string): Promise<string> {//this private method is used to resolve the site drive id
    const drive: any = await firstValueFrom(
      graphGet(this.http, `/sites/${siteId}/drive?$select=id`, token),
    );
    if (!drive?.id) {
      throw new Error('Could not resolve SharePoint document library for attachments');
    }
    return drive.id;
  }

  private getSiteWebUrl(): string {
    return `https://${sharePointConfig.siteHostName}/${sharePointConfig.sitePath}`;
  }

  private async uploadAttachmentWithFallback(
    siteId: string,
    file: File,
    eFormListId: string,
    graphToken: string,
  ): Promise<{ webUrl: string; name: string }> {
    try {
      const driveId = await this.resolveSiteDriveId(siteId, graphToken);
      return await this.uploadAttachmentViaGraph(driveId, file, eFormListId, graphToken);
    } catch (graphError) {
      console.warn('Graph attachment upload failed, trying SharePoint REST', graphError);
      try {
        const restToken = await this.authService.acquireSharePointRestWriteToken();
        return await this.uploadAttachmentViaRest(file, eFormListId, restToken);
      } catch (restError) {
        const graphMessage = this.formatError(graphError);
        const restMessage = this.formatError(restError);
        throw new Error(
          `Could not upload attachment. Graph: ${graphMessage}. REST: ${restMessage}`,
        );
      }
    }
  }

  private async uploadAttachmentViaGraph(
    driveId: string,
    file: File,
    eFormListId: string,
    token: string,
  ): Promise<{ webUrl: string; name: string }> {
    const safeName = file.name.replace(/[<>:"/\\|?*#%]+/g, '_');
    const itemPath = `PaperlessComments/${eFormListId}/${Date.now()}_${safeName}`
      .split('/')
      .map(segment => encodeURIComponent(segment))
      .join('/');

    const uploaded: any = await firstValueFrom(
      graphPut(
        this.http,
        `/drives/${driveId}/root:/${itemPath}:/content`,
        file,
        token,
        file.type || 'application/octet-stream',
        AppConstants.graphFileListingTimeoutMs,
      ),
    );

    return this.extractUploadedFileLink(uploaded, safeName);
  }

  private async uploadAttachmentViaRest(
    file: File,
    eFormListId: string,
    token: string,
  ): Promise<{ webUrl: string; name: string }> {
    const siteWebUrl = this.getSiteWebUrl();
    const safeName = file.name.replace(/[<>:"/\\|?*#%]+/g, '_');
    const fileName = `${Date.now()}_${safeName}`;
    const arrayBuffer = await file.arrayBuffer();

    const folderCandidates = [
      `/${sharePointConfig.sitePath}/Shared Documents/PaperlessComments/${eFormListId}`,
      `/${sharePointConfig.sitePath}/Shared Documents/PaperlessComments`,
      `/${sharePointConfig.sitePath}/PaperlessComments/${eFormListId}`,
    ];

    let lastError: unknown = null;
    for (const folderPath of folderCandidates) {
      try {
        const escapedFolder = folderPath.replace(/'/g, "''");
        const escapedFile = fileName.replace(/'/g, "''");
        const resp: any = await firstValueFrom(
          this.http.post(
            `${siteWebUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedFolder}')/Files/add(url='${escapedFile}',overwrite=true)`,
            arrayBuffer,
            {
              headers: {
                Authorization: `Bearer ${token}`,
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/octet-stream',
              },
            },
          ),
        );

        const serverRelative = String(
          resp?.d?.ServerRelativeUrl ?? resp?.d?.serverRelativeUrl ?? '',
        ).trim();
        if (!serverRelative) {
          throw new Error('SharePoint REST upload did not return a file path');
        }

        return {
          webUrl: normalizeSharePointFileUrl(
            `https://${sharePointConfig.siteHostName}${serverRelative}`,
          ),
          name: String(resp?.d?.Name ?? resp?.d?.name ?? safeName),
        };
      } catch (err) {
        lastError = err;
      }
    }

    throw lastError ?? new Error('Could not upload attachment via SharePoint REST');
  }

  private extractUploadedFileLink(uploaded: any, fallbackName: string): { webUrl: string; name: string } {
    const webUrl = String(
      uploaded?.webUrl ??
      uploaded?.['@microsoft.graph.downloadUrl'] ??
      uploaded?.link?.webUrl ??
      '',
    ).trim();
    if (!webUrl) {
      throw new Error('Attachment uploaded but no SharePoint link was returned');
    }

    return {
      webUrl: normalizeSharePointFileUrl(webUrl),
      name: String(uploaded?.name ?? fallbackName),
    };
  }

  private buildCreateAttempts(
    title: string,
    commentText: string,
    eFormListId: string,
    submittedByName: string,
    attachmentUrl: string,
    _attachmentFileName: string,
    categoryValue: string,
    fieldMap: CommentFieldMap,
    columns: ListColumn[],
    parentFields: Record<string, unknown>,
  ): Record<string, unknown>[] {
    const writable = new Set(
      columns.filter(col => !col.readOnly).map(col => col.name),
    );
    const eFormValue = this.coerceFieldValue(eFormListId);
    const attempts: Record<string, unknown>[] = [];

    const applyComment = (fields: Record<string, unknown>): void => {
      for (const key of this.getCommentColumnKeys(fieldMap.comment, writable)) {
        fields[key] = commentText;
      }
    };

    const applyEFormListId = (fields: Record<string, unknown>): void => {
      for (const key of this.getEFormListIdColumnKeys(fieldMap.eFormListId, writable)) {
        fields[key] = eFormValue;
      }
    };

    const applySubmitter = (fields: Record<string, unknown>): void => {
      if (!submittedByName) return;
      const keys = new Set<string>();
      if (fieldMap.submitter) keys.add(fieldMap.submitter);
      for (const key of ['Requestor', 'SubmittedBy', 'Submitter']) {
        keys.add(key);
      }
      for (const key of keys) {
        if (writable.has(key)) {
          fields[key] = submittedByName;
        }
      }
    };

    const applyAttachmentLink = (fields: Record<string, unknown>): void => {
      if (!attachmentUrl || !fieldMap.attachmentUrl || !writable.has(fieldMap.attachmentUrl)) {
        return;
      }
      fields[fieldMap.attachmentUrl] = attachmentUrl;
    };

    const applyCategory = (fields: Record<string, unknown>): void => {
      if (!categoryValue || !fieldMap.category || !writable.has(fieldMap.category)) {
        return;
      }
      fields[fieldMap.category] = categoryValue;
    };

    const buildBody = (includeRequiredFromParent: boolean): Record<string, unknown> => {
      const body: Record<string, unknown> = { Title: title };
      applyComment(body);
      applyEFormListId(body);
      applySubmitter(body);
      applyAttachmentLink(body);
      applyCategory(body);
      this.applyParentAssignee(body, writable, parentFields);
      if (includeRequiredFromParent) {
        this.applyRequiredFields(body, columns, parentFields);
      }
      return body;
    };

    // Primary payload: Title + comment + eFormListId (+ assignee from parent when possible).
    attempts.push(buildBody(true));
    attempts.push(buildBody(false));

    const unique = new Map<string, Record<string, unknown>>();
    for (const attempt of attempts) {
      if (!this.hasMinimumCommentFields(attempt, fieldMap)) continue;
      unique.set(JSON.stringify(attempt), attempt);
    }

    if (!unique.size) {
      throw new Error('Could not build a valid SharePoint payload for this comment.');
    }

    return [...unique.values()];
  }

  /** Writable comment columns used on Paperless HR lists. */
  private getCommentColumnKeys(
    resolved: string | undefined,
    writable: Set<string>,
  ): string[] {
    const keys: string[] = [];
    for (const candidate of [resolved, 'field_10', 'Comment', 'Notes']) {
      if (candidate && writable.has(candidate) && !keys.includes(candidate)) {
        keys.push(candidate);
      }
    }
    return keys;
  }

  /** Writable eFormListId columns used to group Level 2 tasks and comments. */
  private getEFormListIdColumnKeys(
    resolved: string | undefined,
    writable: Set<string>,
  ): string[] {
    const keys: string[] = [];
    for (const candidate of [resolved, 'eFormListId', 'field_11']) {
      if (candidate && writable.has(candidate) && !keys.includes(candidate)) {
        keys.push(candidate);
      }
    }
    return keys;
  }

  private hasMinimumCommentFields(
    fields: Record<string, unknown>,
    fieldMap: CommentFieldMap,
  ): boolean {
    const hasComment = !!fieldMap.comment && Object.prototype.hasOwnProperty.call(fields, fieldMap.comment);
    const hasEForm = !!fieldMap.eFormListId &&
      Object.prototype.hasOwnProperty.call(fields, fieldMap.eFormListId);
    return hasComment && hasEForm;
  }

  /**
   * Set the assignee lookup id on the new comment item.
   *
   * Requirement (Paperless): comments should be assigned back to the ORIGINAL SUBMITTER
   * of the parent task item (Created By / Author), not the current assignee.
   *
   * We do this by preferring "submitter/requestor" person lookup ids from the parent
   * item when present (Requestor/SubmittedBy/Submitter), falling back to Author,
   * then finally to the original assignee.
   */
  private applyParentAssignee(
    target: Record<string, unknown>,
    writable: Set<string>,
    parentFields: Record<string, unknown>,
  ): void {
    const extractLookupId = (value: unknown): number | string | null => {
      if (typeof value === 'number' || typeof value === 'string') return value;
      if (!value || typeof value !== 'object') return null;
      const obj = value as Record<string, unknown>;
      const candidate =
        obj['LookupId'] ??
        obj['lookupId'] ??
        obj['id'] ??
        obj['Id'];
      if (typeof candidate === 'number' || typeof candidate === 'string') return candidate;
      return null;
    };

    const submitterCandidate =
      extractLookupId(parentFields['RequestorLookupId']) ??
      extractLookupId(parentFields['SubmittedByLookupId']) ??
      extractLookupId(parentFields['SubmitterLookupId']) ??
      extractLookupId(parentFields['RequestorId']) ??
      extractLookupId(parentFields['SubmittedById']) ??
      extractLookupId(parentFields['SubmitterId']) ??
      extractLookupId(parentFields['Requestor']) ??
      extractLookupId(parentFields['SubmittedBy']) ??
      extractLookupId(parentFields['Submitter']) ??
      extractLookupId(parentFields['CreatedBy']) ??
      extractLookupId(parentFields['Author']) ??
      extractLookupId(parentFields['AuthorLookupId']);

    for (const key of ['AssignedToLookupId', 'AssignedToId', 'field_7LookupId']) {
      if (!writable.has(key)) continue;
      const value = submitterCandidate ?? parentFields[key];
      if (value == null || value === '') continue;
      target[key] = value;
    }
  }

  /** Copy only required writable values from the parent item, using the same shape Graph returned. */
  private applyRequiredFields(
    target: Record<string, unknown>,
    columns: ListColumn[],
    parentFields: Record<string, unknown>,
  ): void {
    for (const column of columns) {
      if (!column.required || column.readOnly) continue;
      const name = column.name;
      if (!name || name === 'Title' || Object.prototype.hasOwnProperty.call(target, name)) continue;

      const parentValue = parentFields[name];
      if (parentValue == null || parentValue === '') continue;
      if (typeof parentValue === 'object') continue;

      target[name] = parentValue;
    }
  }

  private coerceFieldValue(value: string): string | number {
    return /^\d+$/.test(value) ? Number(value) : value;
  }

  private async postListItemAttempts(
    siteId: string,
    listId: string,
    listName: string,
    attempts: Record<string, unknown>[],
    token: string,
    fieldMap: CommentFieldMap,
    eFormListId: string,
    commentText: string,
  ): Promise<{ id: string; fields: Record<string, unknown>; webUrl: string }> {
    let lastError: unknown = null;

    for (const fields of attempts) {
      const mutable = { ...fields };
      for (let retry = 0; retry < 4; retry++) {
        try {
          const created = await firstValueFrom(
            graphPost(
              this.http,
              `/sites/${siteId}/lists/${listId}/items`,
              { fields: mutable },
              token,
              AppConstants.graphDefaultTimeoutMs,
            ),
          ) as { id: string; fields?: Record<string, unknown> };

          const verified = await this.verifyAndRepairCreatedItem(
            siteId,
            listId,
            listName,
            created.id,
            token,
            fieldMap,
            eFormListId,
            commentText,
            mutable,
          );

          const webUrl =
            `https://${sharePointConfig.siteHostName}/${sharePointConfig.sitePath}` +
            `/Lists/${listName}/DispForm.aspx?ID=${created.id}`;

          console.info('[CommentService] Comment saved to SharePoint', {
            listName,
            itemId: created.id,
            webUrl,
            fieldsWritten: Object.keys(mutable),
          });

          return { id: created.id, fields: verified, webUrl };
        } catch (err) {
          lastError = err;
          const unknownField = this.extractUnknownField(err);
          if (unknownField && unknownField in mutable) {
            delete mutable[unknownField];
            if (!this.hasMinimumCommentFields(mutable, fieldMap)) {
              break;
            }
            continue;
          }
          break;
        }
      }
    }

    throw lastError ?? new Error('Failed to create comment item');
  }

  private async verifyAndRepairCreatedItem(
    siteId: string,
    listId: string,
    listName: string,
    itemId: string,
    token: string,
    fieldMap: CommentFieldMap,
    eFormListId: string,
    commentText: string,
    fieldsWritten: Record<string, unknown>,
  ): Promise<Record<string, unknown>> {
    const readBack: any = await firstValueFrom(
      graphGet(
        this.http,
        `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`,
        token,
      ),
    );
    const saved = (readBack?.fields ?? {}) as Record<string, unknown>;

    const patch: Record<string, unknown> = {};
    if (fieldMap.comment) {
      const savedComment = String(saved[fieldMap.comment] ?? '').trim();
      if (!savedComment && commentText) {
        patch[fieldMap.comment] = commentText;
      }
    }
    if (fieldMap.eFormListId) {
      const savedEForm = String(saved[fieldMap.eFormListId] ?? '').trim();
      if (!savedEForm && eFormListId) {
        patch[fieldMap.eFormListId] = this.coerceFieldValue(eFormListId);
      }
    }

    if (Object.keys(patch).length) {
      try {
        await firstValueFrom(
          graphPatch(
            this.http,
            `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
            patch,
            token,
          ),
        );
        console.warn('[CommentService] Repaired missing fields after create', {
          listName,
          itemId,
          patch,
          fieldsWritten: Object.keys(fieldsWritten),
        });
        return { ...saved, ...patch };
      } catch (err) {
        console.warn('[CommentService] Created item but could not repair fields', err);
      }
    }

    return saved;
  }

  private async loadListColumns(
    siteId: string,
    listId: string,
    token: string,
  ): Promise<ListColumn[]> {
    const columnsResp: any = await firstValueFrom(
      graphGet(
        this.http,
        `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,readOnly,required,choice`,
        token,
      ),
    );
    return (columnsResp?.value ?? [])
      .map((col: any) => ({
        name: String(col.name ?? '').trim(),
        displayName: String(col.displayName ?? '').trim(),
        readOnly: col.readOnly === true,
        required: col.required === true,
        choices: Array.isArray(col.choice?.choices)
          ? col.choice.choices.map((choice: unknown) => String(choice ?? '').trim()).filter(Boolean)
          : [],
      }))
      .filter((col: ListColumn) => col.name);
  }

  private resolveCommentFieldMap(columns: ListColumn[]): CommentFieldMap {
    return {
      comment: this.pickColumn(columns, ['comment', 'notes'], ['Comment', 'Notes', 'field_10']),
      eFormListId: this.pickColumn(
        columns,
        ['eformlistid', 'eform list id'],
        ['eFormListId', 'field_11'],
      ),
      submitter: this.pickColumn(
        columns,
        ['requestor', 'submittedby', 'submitter', 'author'],
        ['Requestor', 'SubmittedBy', 'Submitter', 'Author'],
      ),
      attachmentUrl: this.pickColumn(
        columns,
        ['attachmenturl', 'attachmentlink', 'documentlink', 'link'],
        ['AttachmentUrl', 'AttachmentLink', 'DocumentLink', 'Link'],
      ),
      category: this.pickColumn(
        columns,
        ['category', 'type'],
        ['Category', 'Type', 'field_9'],
      ),
    };
  }

  private pickColumn(
    columns: ListColumn[],
    displayTokens: string[],
    internalCandidates: string[],
  ): string | undefined {
    const writable = columns.filter(col => !col.readOnly);

    for (const candidate of internalCandidates) {
      const match = writable.find(
        col => col.name.toLowerCase() === candidate.toLowerCase(),
      );
      if (match) return match.name;
    }

    for (const col of writable) {
      const display = col.displayName.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (displayTokens.some(token => display.includes(token.replace(/[^a-z0-9]/g, '')))) {
        return col.name;
      }
    }

    return undefined;
  }

  private async fetchParentFields(
    siteId: string,
    listId: string,
    parentTaskId: string,
    token: string,
    fallback?: Record<string, unknown>,
  ): Promise<Record<string, unknown>> {
    if (!parentTaskId) {
      return fallback ?? {};
    }

    try {
      const resp: any = await firstValueFrom(
        graphGet(
          this.http,
          `/sites/${siteId}/lists/${listId}/items/${parentTaskId}?$expand=fields`,
          token,
        ),
      );
      return resp?.fields ?? fallback ?? {};
    } catch {
      return fallback ?? {};
    }
  }

  private async acquireTokenForCommentWrite(): Promise<string> {
    try {
      return await this.authService.acquireSharePointWriteToken();
    } catch {
      return this.authService.acquireSharePointToken();
    }
  }

  private extractUnknownField(err: unknown): string | null {
    const message = this.formatError(err);
    const match = message.match(/field '([^']+)' is not recognized/i);
    return match?.[1] ?? null;
  }

  private formatError(err: unknown): string {
    if (err instanceof HttpErrorResponse) {
      const body = err.error as {
        error?: { message?: string; details?: Array<{ message?: string }> };
        message?: string;
      } | string;

      if (typeof body === 'string' && body.trim()) return body;

      const details = (body as { error?: { details?: Array<{ message?: string }> } })?.error?.details;
      const detailText = details?.map(d => d.message).filter(Boolean).join(' ');
      const main =
        (body as { error?: { message?: string } })?.error?.message ||
        (body as { message?: string })?.message ||
        err.message;

      return [main, detailText].filter(Boolean).join(' — ') || 'Failed to create comment';
    }

    if (err instanceof Error && err.message) {
      return err.message;
    }

    return 'Failed to create comment';
  }

  private resolveCommentTitle(category?: string, isAttachment = false): string {
    // Action requests can also include attachments; keep the title as "Request for Action"
    // instead of letting the attachment flag override it.
    if (category === 'action') return 'Request for Action';
    if (isAttachment || category === 'attachment') return 'New Attachment';
    return 'General Comment';
  }

  private resolveStatusColumn(columns: ListColumn[]): string | undefined {
    return this.pickColumn(
      columns,
      [
        'status',
        'laststate',
        'approvalstatus',
        'workflowstatus',
        'taskoutcome',
        'outcome',
        'currentstage',
        'stage',
        'completed',
        'iscompleted',
        'state',
        'progress',
      ],
      [...STATUS_FIELD_CANDIDATES],
    );
  }

  private resolveStatusColumnFromFields(fields: Record<string, unknown>): string | undefined {
    for (const key of STATUS_FIELD_CANDIDATES) {
      const value = fields[key];
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return key;
      }
    }
    return undefined;
  }

  private resolveStatusColumnForItem(
    columns: ListColumn[],
    fields: Record<string, unknown>,
  ): string | undefined {
    const fromItem = this.resolveStatusColumnFromFields(fields);
    if (fromItem && columns.some(col => col.name === fromItem)) {
      return fromItem;
    }
    return this.resolveStatusColumn(columns);
  }

  private resolveCompletionStatusValues(
    statusColumn: string,
    columns: ListColumn[],
    currentStatus: string,
  ): string[] {
    const column = columns.find(col => col.name === statusColumn);
    const choices = column?.choices ?? [];
    const completionLike = choices.filter(choice => this.isStatusComplete(choice));
    const pendingLike = new Set(
      choices
        .filter(choice => this.isStatusPendingLike(choice))
        .map(choice => choice.toLowerCase()),
    );

    const candidates: string[] = [];
    const addCandidate = (value: string) => {
      const trimmed = String(value ?? '').trim();
      if (!trimmed) return;
      if (trimmed.toLowerCase() === currentStatus.toLowerCase()) return;
      if (pendingLike.has(trimmed.toLowerCase())) return;
      if (!candidates.some(existing => existing.toLowerCase() === trimmed.toLowerCase())) {
        candidates.push(trimmed);
      }
    };

    for (const value of completionLike) {
      addCandidate(value);
    }

    for (const value of DEFAULT_COMPLETION_STATUS_VALUES) {
      if (
        !choices.length ||
        choices.some(choice => choice.toLowerCase() === value.toLowerCase())
      ) {
        addCandidate(value);
      }
    }

    for (const value of choices) {
      if (!this.isStatusPendingLike(value) && !this.isStatusRejectedLike(value)) {
        addCandidate(value);
      }
    }

    return candidates;
  }

  private buildCommentPatch(
    fieldMap: CommentFieldMap,
    updatedComment: string,
    writable: Set<string>,
  ): Record<string, unknown> {
    const patch: Record<string, unknown> = {};
    for (const key of this.getCommentColumnKeys(fieldMap.comment, writable)) {
      patch[key] = updatedComment;
    }
    return patch;
  }

  private isStatusComplete(status: string): boolean {
    const normalized = String(status ?? '').trim().toLowerCase();
    return (
      normalized.includes('complet') ||
      normalized.includes('approv') ||
      normalized.includes('done') ||
      normalized.includes('closed') ||
      normalized.includes('finished') ||
      normalized.includes('resolved')
    );
  }

  private isStatusPendingLike(status: string): boolean {
    const normalized = String(status ?? '').trim().toLowerCase();
    return (
      normalized.includes('pending') ||
      normalized.includes('progress') ||
      normalized.includes('submitted') ||
      normalized.includes('open') ||
      normalized.includes('new')
    );
  }

  private isStatusRejectedLike(status: string): boolean {
    const normalized = String(status ?? '').trim().toLowerCase();
    return normalized.includes('reject') || normalized.includes('denied') || normalized.includes('cancel');
  }

  private async readItemFields(
    siteId: string,
    listId: string,
    itemId: string,
    token: string,
  ): Promise<Record<string, unknown>> {
    const readBack: any = await firstValueFrom(
      graphGet(
        this.http,
        `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`,
        token,
      ),
    );
    return (readBack?.fields ?? {}) as Record<string, unknown>;
  }

  private readCommentFromFields(
    fields: Record<string, unknown>,
    fieldMap: CommentFieldMap,
  ): string {
    const writable = new Set(Object.keys(fields));
    const keys = this.getCommentColumnKeys(fieldMap.comment, writable);
    for (const key of keys) {
      const value = String(fields[key] ?? '').trim();
      if (value) {
        return value;
      }
    }
    return String(fields['Comment'] ?? fields['Notes'] ?? fields['field_10'] ?? '').trim();
  }

  private buildCompletionComment(
    existingComment: string,
    completionText: string,
    completedByName: string,
  ): string {
    const stamp = new Date().toISOString().slice(0, 10);
    const byLine = `[Completed by ${completedByName} on ${stamp}]`;
    const note = completionText ? `${byLine}: ${completionText}` : byLine;
    return existingComment ? `${existingComment}\n\n${note}` : note;
  }

  private async resolveListId(_siteId: string, listName: string, token: string): Promise<string> {
    const targetList = await this.siteMetadata.findList(token, listName);
    if (!targetList?.id) {
      throw new Error(`List "${listName}" not found on site`);
    }
    return targetList.id;
  }
}