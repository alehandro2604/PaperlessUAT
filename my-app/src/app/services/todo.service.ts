//this file manages the to-do tasks state and provides methods to interact with them

import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { BehaviorSubject, Observable, firstValueFrom } from 'rxjs';
import { graphGet, graphPatch } from '../microsoft-graph';
import { ToDoTask } from '../types/todo-task.interface';
import { AuthService } from './auth.service';
import { CommentService, CommentTask } from './comment.service';
import { DelegateService, DomainUser } from './delegate.service';
import { SiteMetadataService } from './site-metadata.service';
import { UserService } from './user.service';

export interface TaskComment {
  id: string;
  taskId: string;
  category: string;
  comment: string;
  createdAtIso: string;
}

/** Candidate internal names for SharePoint due-date columns (matches app.ts mapping). */
const DUE_DATE_FIELD_CANDIDATES = ['DueDate', 'DueDate0', 'Due', 'Due_x0020_Date'] as const;

@Injectable({
  providedIn: 'root'
})
export class TodoService {
  private readonly commentsByTaskId$ = new BehaviorSubject<Record<string, TaskComment[]>>({});
  private tasks = new BehaviorSubject<ToDoTask[]>([]);
  private loading = new BehaviorSubject<boolean>(false);

  constructor(
    private readonly http: HttpClient,
    private readonly authService: AuthService,
    private readonly siteMetadata: SiteMetadataService,
    private readonly commentService: CommentService,
    private readonly delegateService: DelegateService,
    private readonly userService: UserService,
  ) {}

  async addComment(
    taskId: number | string,
    comment: {
      comment: string;
      category: string;
      listName: string;
      eFormListId: string;
      eFormTitle?: string;
      isAttachment?: boolean;
      attachmentFile?: File | null;
      attachmentFileName?: string;
      assignedToUser?: DomainUser | null;
    }
  ): Promise<CommentTask> {
    const id = String(taskId ?? '').trim();
    if (!id) {
      throw new Error('No task ID provided');
    }

    const parentTask = this.getTaskById(id);
    const parentRawFields = (parentTask?.eFormDetails?.rawFields ?? {}) as Record<string, unknown>;
    const listName =
      String(comment.listName ?? '').trim() ||
      String(parentTask?.eFormDetails?.listName ?? '').trim();

    if (!listName) {
      throw new Error('No list name available. Select a Level 2 task group first.');
    }

    const created = await firstValueFrom(
      this.commentService.createCommentTask({
        parentTaskId: id,
        listName,
        eFormListId: comment.eFormListId,
        eFormTitle: comment.eFormTitle,
        commentText: comment.comment,
        category: comment.category,
        submittedByName: this.userService.getCurrentUserName(),
        parentRawFields,
        attachmentFile: comment.attachmentFile ?? null,
        attachmentFileName: comment.attachmentFileName ?? comment.attachmentFile?.name ?? '',
        assignedToUser: comment.assignedToUser ?? null,
        isComment: true,
        isAttachment: comment.isAttachment === true,
        isTask: false,
        isCommentTask: true,
      }),
    );

    const assignedToUser = comment.assignedToUser ?? null;
    if (comment.category === 'action' && assignedToUser) {
      await this.assignActionRequestToUser(
        id,
        listName,
        parentTask,
        parentRawFields,
        assignedToUser,
      );
      created.assignedToUser = assignedToUser;
    }

    const entry: TaskComment = {
      id: created.id ?? globalThis.crypto?.randomUUID?.() ?? String(Date.now()),
      taskId: id,
      category: String(comment.category ?? '').trim(),
      comment: String(comment.comment ?? '').trim(),
      createdAtIso: new Date().toISOString(),
    };

    const current = this.commentsByTaskId$.value;
    const existing = current[id] ?? [];
    this.commentsByTaskId$.next({
      ...current,
      [id]: [entry, ...existing],
    });

    return created;
  }

  getCommentsForTask(taskId: string): Observable<TaskComment[]> {
    return new Observable<TaskComment[]>(subscriber => {
      const sub = this.commentsByTaskId$.subscribe(map => {
        subscriber.next(map[taskId] ?? []);
      });
      return () => sub.unsubscribe();
    });
  }

  isLoading(): boolean {
    return this.loading.value;
  }

  getLoading$() {
    return this.loading.asObservable();
  }

  setLoading(loading: boolean): void {
    this.loading.next(loading);
  }

  setTasksFromSharePoint(rawItems: any[], options?: { clearLoading?: boolean }): void {
    const mapped: ToDoTask[] = rawItems.map(item => ({
      id:            item.id,
      name:          item.name,
      description:   item.description ?? '',
      status:        this.mapStatus(item.status),
      assignedTo:    item.eFormDetails?.assignedTo ?? '',
      isAssignedToCurrentUser: item.isAssignedToCurrentUser === true,
      // Never emit the literal "Unknown User" — prefer empty string when submitter can't be resolved.
      submittedBy:   String(
        item.submittedBy ??
        item.eFormDetails?.submitter ??
        item.eFormDetails?.submittedBy ??
        item.author?.title ??
        item.createdBy?.title ??
        ''
      ).trim(),
      listName:      item.listName ?? '',
      submittedDate: item.submittedDate ?? '',
      statusDate:    item.statusDate ?? '',
      createdDate:   new Date(item.submittedDate ?? Date.now()),
      updatedDate:   new Date(item.lastModifiedDateTime ?? Date.now()),
      eFormDetails:  item.eFormDetails,
      webUrl:        item.webUrl
    }));

    this.tasks.next(mapped);
    if (options?.clearLoading !== false) {
      this.loading.next(false);
    }
  }

  getTasksForCurrentUser(): Observable<ToDoTask[]> {
    return this.tasks.asObservable();
  }

  getTaskById(taskId: string): ToDoTask | undefined {
    const id = String(taskId ?? '').trim();
    if (!id) return undefined;
    return this.tasks.value.find(task => String(task.id ?? '').trim() === id);
  }

  private async assignActionRequestToUser(
    taskId: string,
    listName: string,
    parentTask: ToDoTask | undefined,
    parentRawFields: Record<string, unknown>,
    assignedToUser: DomainUser,
  ): Promise<void> {
    const assigneeEmail = (assignedToUser.Email || assignedToUser.UserPrincipalName || '').trim();
    if (!assigneeEmail) {
      throw new Error(
        `No email found for ${assignedToUser.Title || 'the selected user'}. Choose a user with an email address.`,
      );
    }

    await this.delegateService.updateAssignedTo(
      {
        id: taskId,
        listName,
        eFormDetails: {
          listName,
          rawFields: parentRawFields,
          assignedTo: parentTask?.assignedTo ?? parentTask?.eFormDetails?.assignedTo ?? '',
        },
      },
      assignedToUser,
    );

    this.updateLocalTaskAssignee(taskId, assignedToUser);
  }

  private updateLocalTaskAssignee(taskId: string, assignedToUser: DomainUser): void {
    const assigneeLabel = assignedToUser.Title || assignedToUser.Email || assignedToUser.UserPrincipalName;
    const updated = this.tasks.value.map(task =>
      task.id === taskId
        ? {
            ...task,
            assignedTo: assigneeLabel,
            eFormDetails: {
              ...(task.eFormDetails ?? {}),
              assignedTo: assigneeLabel,
            },
            updatedDate: new Date(),
          }
        : task,
    );
    this.tasks.next(updated);
  }

  private mapStatus(raw: string): ToDoTask['status'] {
    const s = (raw ?? '').toLowerCase();
    if (s.includes('reject') || s.includes('denied')) return 'rejected';
    if (this.isPendingApprovalStatus(s)) return 'pending';
    if (s.includes('approv') || s.includes('complet')) return 'completed';
    if (s.includes('progress') || s.includes('stage')) return 'in_progress';
    return 'pending';
  }

  /** "Pending Approval" / "Awaiting Approval" must stay open — not completed. */
  private isPendingApprovalStatus(status: string): boolean {
    return /\b(pending|awaiting|waiting|needs?)\b.*\bapprov/.test(status)
      || /\bawait(ing|s)?\s+approv/.test(status);
  }

  completeTask(taskId: string): void {
    const updated = this.tasks.value.map(t =>
      t.id === taskId
        ? { ...t, status: 'completed' as const, updatedDate: new Date() }
        : t
    );
    this.tasks.next(updated);
  }

  removeTask(taskId: string): void {
    const id = String(taskId ?? '').trim();
    if (!id) return;
    this.tasks.next(this.tasks.value.filter(task => String(task.id ?? '').trim() !== id));
  }

  rejectTask(taskId: string): void {
    const updated = this.tasks.value.map(t =>
      t.id === taskId
        ? { ...t, status: 'rejected' as const, updatedDate: new Date() }
        : t
    );
    this.tasks.next(updated);
  }

  /**
   * Writes due date to the SharePoint list item via Graph, then syncs local state.
   * `dueDate` must be `YYYY-MM-DD`.
   */
  async updateTaskDueDate(task: ToDoTask | Record<string, unknown>, dueDate: string): Promise<void> {
    const itemId = String((task as any)?.id ?? '').trim();
    const listName = String(
      (task as any)?.listName ?? (task as any)?.eFormDetails?.listName ?? '',
    ).trim();
    const normalized = String(dueDate ?? '').trim();

    if (!itemId) throw new Error('Cannot update due date: task has no id');
    if (!listName) throw new Error('Cannot update due date: task has no listName');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
      throw new Error(`Invalid due date "${dueDate}". Expected YYYY-MM-DD.`);
    }

    const token = await this.acquireWriteToken();
    const { siteId } = await this.siteMetadata.resolve(token);
    if (!siteId) throw new Error('Could not resolve SharePoint site ID');

    const list = await this.siteMetadata.findList(token, listName);
    if (!list?.id) throw new Error(`List "${listName}" not found on site`);

    const rawFields = (task as any)?.eFormDetails?.rawFields as Record<string, unknown> | undefined;
    const fieldNames = await this.resolveDueDateFieldNames(siteId, list.id, token, rawFields);

    // Prefer the format already stored on the item; otherwise try date-only then noon UTC.
    const existing = this.readExistingDueDateValue(rawFields, fieldNames);
    const values = this.buildDueDatePayloadValues(normalized, existing);

    let lastError: unknown = null;
    for (const fieldName of fieldNames) {
      for (const value of values) {
        try {
          await firstValueFrom(
            graphPatch(
              this.http,
              `/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`,
              { [fieldName]: value },
              token,
              12_000,
            ),
          );
          this.updateLocalTaskDueDate(itemId, normalized, fieldName);
          return;
        } catch (err) {
          lastError = err;
          const status = Number((err as { status?: number })?.status);
          const msg = this.formatGraphError(err).toLowerCase();
          // Fatal auth/permission errors — don't burn time on more field names.
          if (
            status === 401 ||
            status === 403 ||
            msg.includes('access denied') ||
            msg.includes('forbidden') ||
            msg.includes('unauthorized')
          ) {
            throw new Error(
              `Could not update due date (permission denied). ${this.formatGraphError(err)}`.trim(),
            );
          }
          console.warn('Due date Graph PATCH failed:', fieldName, value, err);
        }
      }
    }

    throw new Error(
      `Could not update due date on list "${listName}" item ${itemId}. ${this.formatGraphError(lastError)}`.trim(),
    );
  }

  private async acquireWriteToken(): Promise<string> {
    try {
      const token = await this.authService.acquireSharePointWriteToken();
      if (token) return token;
    } catch (err) {
      console.warn('SharePoint write token failed; falling back to read token.', err);
    }
    const fallback = await this.authService.acquireSharePointToken();
    if (!fallback) throw new Error('Unable to acquire SharePoint token');
    return fallback;
  }

  private async resolveDueDateFieldNames(
    siteId: string,
    listId: string,
    token: string,
    rawFields?: Record<string, unknown>,
  ): Promise<string[]> {
    const present = DUE_DATE_FIELD_CANDIDATES.filter(
      name => rawFields != null && Object.prototype.hasOwnProperty.call(rawFields, name),
    );
    if (present.length) return [...present];

    // Discover date columns whose name looks like DueDate when the item has never had a value.
    try {
      const cols: any = await firstValueFrom(
        graphGet(
          this.http,
          `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,dateTime`,
          token,
          10_000,
        ),
      );
      const discovered = (cols?.value ?? [])
        .filter((col: any) => col?.dateTime || /due/i.test(String(col?.name ?? col?.displayName ?? '')))
        .map((col: any) => String(col?.name ?? '').trim())
        .filter((name: string) => !!name && /due/i.test(name));
      if (discovered.length) return discovered;
    } catch (err) {
      console.warn('Could not load list columns for due date field discovery.', err);
    }

    return [...DUE_DATE_FIELD_CANDIDATES];
  }

  private readExistingDueDateValue(
    rawFields: Record<string, unknown> | undefined,
    fieldNames: string[],
  ): string {
    if (!rawFields) return '';
    for (const name of fieldNames) {
      const value = String(rawFields[name] ?? '').trim();
      if (value) return value;
    }
    return '';
  }

  private buildDueDatePayloadValues(normalized: string, existing: string): string[] {
    const noonIso = `${normalized}T12:00:00Z`;
    // If the existing value looks like a full datetime, prefer ISO first.
    if (/T\d{2}:\d{2}/.test(existing)) {
      return [noonIso, normalized];
    }
    return [normalized, noonIso];
  }

  private updateLocalTaskDueDate(taskId: string, dueDate: string, fieldName: string): void {
    const updated = this.tasks.value.map(task => {
      if (String(task.id ?? '').trim() !== taskId) return task;
      const details = { ...(task.eFormDetails ?? {}) };
      const rawFields = { ...(details.rawFields ?? {}) };
      rawFields[fieldName] = dueDate;
      details.dueDate = dueDate;
      details.rawFields = rawFields;
      return {
        ...task,
        eFormDetails: details,
        updatedDate: new Date(),
      };
    });
    this.tasks.next(updated);
  }

  private formatGraphError(err: unknown): string {
    const e = err as {
      error?: { error?: { message?: string }; message?: string };
      message?: string;
      status?: number;
    };
    const status = e?.status != null ? `HTTP ${e.status}` : '';
    const message = String(
      e?.error?.error?.message ?? e?.error?.message ?? e?.message ?? '',
    ).trim();
    return [status, message].filter(Boolean).join(': ');
  }
}
