import { Injectable, Injector, inject } from '@angular/core';
import { Observable } from 'rxjs';
import { FormConfigurationService } from '../services/form-configuration.service';
import { EForm } from '../models/form-configuration.model';
import { TodoService } from '../services/todo.service';
import { UserService } from '../services/user.service';
import { DelegateService } from '../services/delegate.service';
import { CommentService } from '../services/comment.service';
import { SubordinaryTaskService } from '../services/subordinaryTask.service';
import { HrTaskFolderMatchService } from '../services/hr-task-folder-match.service';
import { AppDropdownOption } from '../components/app-dropdown/app-dropdown.component';
import { NewCommentSavedEvent, CommentListItem } from '../models/comment.models';
import { buildCommentCardFromSavedEvent } from '../utils/comment-card-builders';
import { isSyntheticCommentCard, shouldHideCommentStatus, isCommentTypeItem, extractCommentCardTaskId, matchesApprovalFilter, onCommentLinkClick } from '../utils/comment-utils';
import { isItemSubmittedBy, getDateSearchTokens, getTaskEFormTitle, extractTrailingFolderId } from '../utils/hr-task-matching';
import { buildHrTaskDetailRows } from '../utils/hr-task-detail';
import { extractPersonName, collectTaskAssigneeValues } from '../utils/person-name-matching';
import { getTaskListName, isSameTask } from '../utils/task-identity';
import { HrTaskDetailTone, HrTaskDetailView } from '../models/hr-task-detail.model';
import { ViewRefreshService } from './view-refresh.service';
import { UiStore } from './ui.store';
import { SessionStore } from './session.store';
import { FolderTasksStore } from './folder-tasks.store';
import { TodoStore } from './todo.store';
import { AllFilesStore } from './all-files.store';

/**
 * The Comments panel: the list of task/comment cards, its filters, search and paging, and the actions on a card (complete, delegate, new comment, view eForm).
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class CommentsStore {
  private readonly injector = inject(Injector);
  private readonly todoService = inject(TodoService);
  private readonly userService = inject(UserService);
  private readonly delegateService = inject(DelegateService);
  private readonly commentService = inject(CommentService);
  private readonly formConfigService = inject(FormConfigurationService);
  private readonly subordinaryTaskService = inject(SubordinaryTaskService);
  private readonly folderMatch = inject(HrTaskFolderMatchService);

  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get ui(): UiStore { return this.injector.get(UiStore); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get folderTasks(): FolderTasksStore { return this.injector.get(FolderTasksStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  // this is to check if the item is a user comment card
  isUserCommentCard(item: { id?: unknown }): boolean {
    return String(item?.id ?? '').startsWith('comment:');
  }

  commentsMessage = 'Loading comments...';

  isLoadingComments = false;

  /** True while a Comments network fetch is still running after the first page is shown. */
  isLoadingMoreComments = false;

  /** After Comments-search click: scroll/highlight this task id in the Comments panel. */
  pendingCommentFocusId: string | null = null;

  highlightedCommentId: string | null = null;

  readonly approvalFilterOptions: AppDropdownOption[] = [
    { value: '', label: 'All statuses', tone: 'neutral' },
    { value: 'approved', label: 'Complete', tone: 'complete' },
    { value: 'pending', label: 'Pending', tone: 'pending' },
    { value: 'general-comments', label: 'General Comments', tone: 'comment' },
    { value: 'new-attachments', label: 'New Attachments', tone: 'attachment' },
    { value: 'rfa', label: 'Request For Action', tone: 'action' },
  ];

  getTaskFilterOptions(forms: EForm[] | null | undefined): AppDropdownOption[] {
    return [
      { value: '', label: 'All Tasks' },
      ...(forms ?? []).map(form => ({
        value: form.taskListName || form.title,
        label: this.ui.formatFormTitle(form.title),
      })),
    ];
  }

  getAssigneeOptions(assignedTo: string, eform?: { status?: string; completedBy?: string; eFormDetails?: any }): AppDropdownOption[] {
    return this.getAssignees(assignedTo).map(person => ({
      value: person,
      label: this.formatAssigneeLabel(person, eform),
    }));
  }

  /**
   * Multi-assignee dropdown only while the task is still open.
   */
  showAssigneeDropdown(
    assignedTo: string,
    eform?: { status?: string; completedBy?: string; eFormDetails?: any },
  ): boolean {
    return this.getAssignees(assignedTo).length >= 2 && !this.isEformComplete(eform);
  }

  /** True when the comment/task card should use Completed By instead of Assigned To. */
  isEformComplete(eform?: { status?: string; eFormDetails?: any }): boolean {
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
  getCommentStatusTone(eform?: { status?: string; eFormDetails?: any }): HrTaskDetailTone {
    const status = String(eform?.eFormDetails?.status ?? '').toLowerCase();
    if (/reject|denied|cancel/.test(status)) return 'danger';
    if (this.isEformComplete(eform)) return 'success';
    if (/pending|progress|waiting|awaiting/.test(status)) return 'warning';
    return 'neutral';
  }

  /** Completed By when approved; falls back to last editor only when no completer is stored. */
  getLastModifiedByName(
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
  formatAssigneeLabel(person: string, eform?: { status?: string; completedBy?: string; eFormDetails?: any }): string {
    const name = String(person ?? '').trim();
    if (!name) return '';
    if (this.hasAssigneeCompleted(name, eform)) {
      return `${name} ✓`;
    }
    return name;
  }

  /**
   * A card the Comments panel shows as done: the eForm is complete ("Completed by"),
   * or a Request for Action whose assignees all carry the ✓. Other pending eForms are
   * kept even if one approver already acted, since their workflow is still open.
   */
  isCompletedCommentCard(eform: { status?: string; completedBy?: string; eFormDetails?: any }): boolean {
    if (this.isEformComplete(eform)) return true;

    const category = String(eform.eFormDetails?.category ?? '').trim().toLowerCase();
    const commentCategory = String(eform.eFormDetails?.commentCategory ?? '').trim().toLowerCase();
    if (category !== 'request for action' && commentCategory !== 'action') return false;

    const assignees = this.getAssignees(this.getAssignedToDisplay(eform as any) ?? '');
    return assignees.length > 0 && assignees.every(person => this.hasAssigneeCompleted(person, eform));
  }

  hasAssigneeCompleted(person: string, eform?: { status?: string; completedBy?: string; eFormDetails?: any }): boolean {
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

  getEformStatus(eform?: { status?: string; eFormDetails?: any }): string {
    return String(eform?.eFormDetails?.status ?? eform?.status ?? '').toLowerCase();
  }

  getCompletedByNames(eform: { completedBy?: string; eFormDetails?: any }): string[] {
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

  commentItems: Array<{
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

  // FILTER STATE - cached, only recomputed when dirty
  selectedTaskFilter = '';

  selectedApprovalFilter = '';

  searchQuery = '';

  // eForms loaded from SharePoint, used to populate the task-type filter dropdown.
  taskFilterForms$!: Observable<EForm[]>;

  _filteredCommentItems: Array<typeof this.commentItems[0]> = [];

  _filtersDirty = true;

  // FILTERED COMMENT ITEMS - latest-first pages + Load more
  get filteredCommentItems(): Array<typeof this.commentItems[0]> {
    if (!this._filtersDirty) return this._filteredCommentItems;
    this._filteredCommentItems = this.filterCommentItems();
    this._filtersDirty = false;
    return this._filteredCommentItems;
  }

  /** Placeholder card count for the Comments skeleton while tasks load. */
  readonly commentsSkeletonRows = [1, 2, 3, 4, 5, 6];

  /** Short, user-facing Comments loader text (no raw folder/email dump). */
  get commentsLoadingLabel(): string {
    if (this.isLoadingMoreComments) {
      return 'Loading more tasks...';
    }
    const custom = this.commentsMessage?.trim();
    if (this.isLoadingComments) {
      // Prefer progress text (e.g. "Scanning ProcTaskArchive...") over a generic label.
      if (custom) return custom;
      if (this.folderTasks.selectedHrPersonalTaskFolder || this.allFiles.selectedFolderName) {
        return 'Loading tasks...';
      }
      return 'Loading comments...';
    }
    return custom || 'Loading comments...';
  }

  /** First paint shows the newest page; "Load more" reveals the next page from the loaded/cached set. */
  visibleItemCount = 20;

  readonly commentsInitialPageSize = 20;

  readonly itemPageSize = 20;

  get pagedCommentItems(): Array<typeof this.commentItems[0]> {
    return this.filteredCommentItems.slice(0, this.visibleItemCount);
  }

  get hasMoreItems(): boolean {
    return this.filteredCommentItems.length > this.visibleItemCount;
  }

  get remainingItemCount(): number {
    return Math.max(this.filteredCommentItems.length - this.visibleItemCount, 0);
  }

  loadMoreItems(): void {
    this.visibleItemCount += this.itemPageSize;
    this.view.refreshView();
  }

  invalidateCommentFilters(resetPagination = true): void {
    this._filtersDirty = true;
    if (resetPagination) {
      this.visibleItemCount = this.commentsInitialPageSize;
    }
  }

  focusPendingCommentIfNeeded(): void {
    const id = String(this.pendingCommentFocusId ?? '').trim();
    if (!id) return;

    const idx = this.filteredCommentItems.findIndex((item) => String(item?.id ?? '') === id);
    if (idx < 0) return;

    this.visibleItemCount = Math.max(this.visibleItemCount, idx + 1);
    this.highlightedCommentId = id;
    this.pendingCommentFocusId = null;
    this.view.refreshView();

    setTimeout(() => {
      const el = document.getElementById(`comment-item-${id}`);
      el?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }, 80);
  }

  /** "list|title" - the file a task or comment belongs to when it has no eForm list ID. */
  private getTaskFileKey(item: { listName?: unknown; name?: unknown; eFormDetails?: Record<string, unknown> }): string {
    return `${getTaskListName(item)}|${getTaskEFormTitle(item).toLowerCase()}`;
  }

  /** "list#id" - item ids are only unique within one SharePoint list. */
  private getTaskItemKey(item: { id?: unknown; listName?: unknown; eFormDetails?: Record<string, unknown> }): string {
    return `${getTaskListName(item)}#${String(item.id ?? '').trim()}`;
  }

  /** The opened To Do group's tasks (list#id) and the files they belong to (list|title). */
  private getSelectedTaskKeys(): { items: Set<string>; files: Set<string> } {
    const items = new Set<string>();
    const files = new Set<string>();
    const groupTitle = String(this.todo.selectedEFormTitle ?? '').trim().toLowerCase();
    for (const id of this.todo.selectedGroupTaskIds) {
      // The same id can exist in several lists, so pick the one whose title is the opened group's.
      const sameId = this.todoService.getTasksById(id);
      const candidates = sameId.length
        ? sameId
        : this.commentItems.filter(item => String(item.id ?? '').trim() === id);
      const task = groupTitle
        ? candidates.find(c => getTaskEFormTitle(c as any).toLowerCase() === groupTitle)
        : (candidates.length === 1 ? candidates[0] : undefined);
      if (!task) continue;
      items.add(this.getTaskItemKey(task as any));
      const fileKey = this.getTaskFileKey(task as any);
      // Without a list and a real title the key would match unrelated items.
      if (!fileKey.startsWith('|') && !fileKey.endsWith('|unknown title')) files.add(fileKey);
    }
    return { items, files };
  }

  filterCommentItems(): Array<typeof this.commentItems[0]> {
    let filtered = this.commentItems;
    if (this.todo.selectedSubmitter && !this.ui.showAllFilesSection) {
      // Tasks with no eForm list ID (e.g. "Monitor Workflow") can't be matched on that ID:
      // every comment without one would match. Link them on list + Title instead.
      const noFormIdTasks = this.todo.selectedEFormListId === 'Unknown eForm'
        ? this.getSelectedTaskKeys()
        : null;

      filtered = filtered.filter(item => {
        const matchesEFormList = !this.todo.selectedEFormListId || (
          (String(item.eFormDetails?.eFormListId ?? '').trim() || 'Unknown eForm') === this.todo.selectedEFormListId
        );
        if (!matchesEFormList) {
          return false;
        }

        if (
          noFormIdTasks &&
          !noFormIdTasks.items.has(this.getTaskItemKey(item)) &&
          !noFormIdTasks.files.has(this.getTaskFileKey(item))
        ) {
          return false;
        }

        const isCommentEntry = isCommentTypeItem(item);
        const selectedTaskId = String(this.selectedTask?.Id ?? '').trim();
        const groupTaskIds = new Set(this.todo.selectedGroupTaskIds);

        if (!isCommentEntry) {
          const matchesEFormTitle = !this.todo.selectedEFormTitle || (
            getTaskEFormTitle(item) === this.todo.selectedEFormTitle
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

        if (this.ui.showToDoSection && !this.todo.superiorMode) {
          if (isCommentEntry) {
            return true;
          }
          if (!isItemSubmittedBy(item, this.todo.selectedSubmitter)) {
            return false;
          }
          // When a submitter is selected in the To Do view, the Comments panel should
          // show the submitter's full eForm history (Pending + Completed), not only
          // items currently assigned to the logged-in user.
          return true;
        }

        if (this.ui.showToDoSection && this.todo.superiorMode && this.todo.selectedSubmitter) {
          if (isCommentEntry) {
            return true;
          }
          return this.isItemAssignedToSubordinate(item, this.todo.selectedSubmitter);
        }

        return true;
      });
    }

    // To Do is for open work: completed cards stay out of its Comments panel.
    // All Files keeps them, as the place to look up past work.
    // A card whose task is still open in the To Do list always stays (e.g. after Claim,
    // the claimer is the sole assignee and could otherwise look "all ✓" while Pending).
    if (this.ui.showToDoSection) {
      filtered = filtered.filter(item =>
        !!this.todoService.getTaskById(String(item.id ?? '')) || !this.isCompletedCommentCard(item),
      );
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

    if (!this.ui.showAllFilesSection && this.selectedTaskFilter) {
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

    if (this.ui.showAllFilesSection) {
      const folderName = String(this.allFiles.selectedFolderName ?? '').trim();
      if (folderName && extractTrailingFolderId(folderName)) {
        filtered = filtered.filter(item => this.folderMatch.doesMappedHrTaskMatchFolder(item, folderName));
      }
    }

    if (this.folderTasks.isHrPersonalFilesContext()) {
      filtered = filtered.filter(item => !isCommentTypeItem(item));
    }

    return filtered;
  }

  getPrimaryItemAssignee(item: { eFormDetails?: Record<string, unknown>; assignedTo?: unknown }): string {
    const details = item.eFormDetails ?? {};
    const raw = String(details['assignedTo'] ?? item.assignedTo ?? '').trim();
    return this.userService.parseAssignees(raw)[0] ?? '';
  }

  getSubordinateItemAssigneeValues(item: { eFormDetails?: Record<string, unknown>; assignedTo?: unknown }): string[] {
    const primary = this.getPrimaryItemAssignee(item);
    return primary ? [primary] : [];
  }

  isItemAssignedToSubordinate(item: { eFormDetails?: Record<string, unknown>; assignedTo?: unknown }, subordinateName: string): boolean {
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

  getAssignees(assignedTo: string | undefined | null): string[] {
    return this.userService.parseAssignees(assignedTo);
  }

  /** Assigned To text for comment cards - rebuilds from raw SharePoint fields when blank. */
  getAssignedToDisplay(eform: { eFormDetails?: any; assignedTo?: string } | null | undefined): string {
    const details = eform?.eFormDetails ?? {};
    const direct = String(details.assignedTo ?? eform?.assignedTo ?? '').trim();
    if (direct) return direct;

    const rawFields = details.rawFields;
    if (rawFields && typeof rawFields === 'object') {
      return collectTaskAssigneeValues(rawFields as Record<string, any>).join(', ');
    }
    return '';
  }

  onTaskFilterChange(value: string): void {
    this.selectedTaskFilter = value;
    this.invalidateCommentFilters();
    this.view.refreshView();
  }

  onApprovalFilterChange(value: string): void {
    this.selectedApprovalFilter = value;
    this.invalidateCommentFilters();
    this.view.refreshView();
  }

  onSearchChange(event: Event): void {
    this.searchQuery = (event.target as HTMLInputElement).value;
    this.invalidateCommentFilters();
    this.view.refreshView();
  }

  /** Hide status badge for general comments, attachments, and request-for-action cards. */
  shouldHideCommentStatus(item: {
    id?: unknown;
    name?: unknown;
    eFormDetails?: Record<string, unknown>;
  }): boolean {
    return shouldHideCommentStatus(item);
  }

  selectedEFormContent: {
    fileName: string;
    content: string;
    contentType: string;
    uploadDate: string;
    eFormDetails?: any;
    detailRows?: { label: string; value: string; sectionBreak?: boolean }[];
    detailView?: HrTaskDetailView;
  } | null = null;

  isLoadingEFormContent = false;

  // Handles the delegate form submission. The SharePoint write is delegated to
  // the service, and the UI is only updated after that write is verified - so an
  // unverified/failed write shows an error instead of a fake success.
  async onDelegateTask(payload: { task: any; newAssignee: any }): Promise<void> {
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
      this.ui.showModal('Delegate Failed', `Could not update SharePoint: ${msg}`, 'error');
    }
  }

  onDelegateSuccess(itemId: string, newAssigneeName: string, newAssigneeEmail = ''): void {
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
    this.folderTasks.syncTaskCachesAfterMutation(item ?? { id: itemId }, 'delegate', {
      assignedToName: newAssigneeName,
      assignedToEmail: newAssigneeEmail,
    });
    this.invalidateCommentFilters();
    this.view.refreshView();
    this.ui.showModal('Task Delegated', `Task has been delegated to ${newAssigneeName}.`, 'success');
  }

  /** Groups an HR task's fields into the sections shown by the details panel. */
  buildHrTaskDetailView(d: any, task: any): HrTaskDetailView {
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

    const me = this.normalizePersonName(this.session.getLoggedInUserDisplayName());
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

  normalizePersonName(value: string): string {
    return String(value || '').toLowerCase().replace(/[^a-z]/g, '');
  }

  openHrTaskForm(task: any): void {
    if (!task.id) return;

    try {
      const d = task.eFormDetails;
      const detailRows = buildHrTaskDetailRows(d, task);
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
    this.view.refreshView();
  }

  closeFormViewer(): void { this.selectedEFormContent = null; }

  isCommentPanelOpen = false;

  selectedTask: { Id: number | string } | null = null;

  selectedCommentListName = '';

  get showRequestForActionCategory(): boolean {
    return this.formConfigService.allowsRequestForAction(
      this.todo.selectedEFormListId ?? '',
      this.todo.selectedEFormTitle ?? '',
      this.selectedCommentListName,
    );
  }

  openNewCommentModal(): void {
    if (!this.selectedTask?.Id) {
      console.error('Select a task first');
      return;
    }
    this.isCommentPanelOpen = true;
  }

  closeNewCommentModal(): void {
    this.isCommentPanelOpen = false;
  }

  /** Marks a Request for Action as completed in SharePoint and removes it from To Do. */
  async onCompleteTask(event: {
    task: any;
    completionText: string;
    done?: (error?: string | null) => void;
  }): Promise<void> {
    const task = event?.task;
    const completionText = String(event?.completionText ?? '').trim();
    const taskId = String(task?.id ?? task?.Id ?? '').trim();
    if (!taskId) {
      event?.done?.('Cannot complete: the task has no id.');
      return;
    }

    try {
      const completedByName = this.session.getLoggedInUserDisplayName();
      const savedFields = await this.commentService.completeRequestForAction(task, completionText, completedByName);
      // Show the value SharePoint now holds (e.g. "Complete"), so the card's pill flips from Pending.
      const savedProgress = String(savedFields?.['Progress'] ?? '').trim();
      const completedStatus =
        ['Progress', 'LastState', 'Status', 'ApprovalStatus', 'WorkflowStatus', 'TaskOutcome']
          .map(key => String(savedFields?.[key] ?? '').trim())
          .find(value => /complet/i.test(value)) || 'Completed';

      // Ids repeat across lists, so remove/update only this list's item.
      const listName = getTaskListName(task);
      this.todoService.removeTask(taskId, listName);
      // Keep Redis in sync - otherwise the completed task reappears on the next
      // cache paint until a full Graph refresh replaces the snapshot.
      this.folderTasks.syncTaskCachesAfterMutation(task, 'complete');

      const completionNote = completionText
        ? `[Completed by ${completedByName}]: ${completionText}`
        : `[Completed by ${completedByName}]`;

      this.commentItems = this.commentItems.map(item => {
        if (!isSameTask(item, taskId, listName)) {
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
          status: completedStatus,
          completedBy: completedByName,
          eFormDetails: {
            ...(item.eFormDetails ?? {}),
            ...(savedProgress ? { progress: savedProgress, eFormProgress: savedProgress } : {}),
            rawFields: { ...(item.eFormDetails?.['rawFields'] ?? {}), ...(savedFields ?? {}) },
            status: completedStatus,
            completedBy: completedByName,
            comment: mergedComment,
            commentHtml: mergedComment.replace(/\n/g, '<br>'),
          },
        };
      });

      this.ui.allFilesComponent?.refreshCommentSearchCache();
      this.invalidateCommentFilters();
      this.view.refreshView();
      event.done?.();
    } catch (err) {
      console.error('Failed to complete Request for Action:', err);
      const message = err instanceof Error ? err.message : 'Failed to complete task';
      // The popup shows the error itself; only fall back to the page banner without one.
      if (event.done) {
        event.done(message);
      } else {
        this.ui.errorMessage = message;
      }
      this.view.refreshView();
    }
  }

  // to update the comment section iteams for the new-comment modal in the todo section
  //also this is saved by the new-comment modal in the todo section grouping the comments by task id.
  onNewCommentSaved(ev: NewCommentSavedEvent): void {
    const taskId = String(ev?.taskId ?? '').trim();
    const base = taskId ? this.resolveCommentCardBase(taskId) : null;
    if (!base) return;

    this.prependCommentItem(buildCommentCardFromSavedEvent(ev, base, taskId, {
      commentAuthor: this.session.getLoggedInUserDisplayName(),
      fallbackEFormListId: this.todo.selectedEFormListId ?? '',
    }));

    const parentTask =
      this.commentItems.find(item => String(item.id ?? '').trim() === String(ev.taskId ?? '').trim()) ??
      { id: ev.taskId };

    if (ev.category === 'action' && ev.assignedToName) {
      this.onDelegateSuccess(String(ev.taskId), ev.assignedToName, ev.assignedToEmail ?? '');
    } else {
      this.folderTasks.syncTaskCachesAfterMutation(parentTask, 'comment');
    }
  }

  /** After a successful SharePoint due-date write, keep Redis snapshots aligned. */
  onTaskDueDateChanged(event: { task: any; dueDate: string }): void {
    const dueDate = String(event?.dueDate ?? '').trim();
    if (!event?.task || !dueDate) return;
    this.folderTasks.syncTaskCachesAfterMutation(event.task, 'dueDate', { dueDate });
  }

  /** A claim is a delegation to yourself, so reuse the delegate cache sync. */
  onTaskClaimed(event: { task: any; name: string; email: string }): void {
    if (!event?.task) return;
    this.folderTasks.syncTaskCachesAfterMutation(event.task, 'delegate', {
      assignedToName: event.name,
      assignedToEmail: event.email,
    });
  }

  // to resolve the comment card base for the new-comment modal in the todo section
  resolveCommentCardBase(taskId: string): CommentListItem | null {
    return (
      this.commentItems.find(item => String(item.id ?? '').trim() === taskId) ??
      this.buildCommentCardFromSelection(taskId)
    );
  }

  onCommentLinkClick(event: Event): void {
    onCommentLinkClick(event);
  }

  prependCommentItem(item: CommentListItem): void {
    this.commentItems = [item, ...this.commentItems];
    this.invalidateCommentFilters();
    this.view.refreshView();
  }

  buildCommentCardFromSelection(taskId: string): (typeof this.commentItems)[0] | null {
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
          eFormListId: this.todo.selectedEFormListId ?? todoTask.eFormDetails?.eFormListId,
          type: this.todo.selectedEFormTitle ?? todoTask.eFormDetails?.type,
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
      name: this.todo.selectedEFormTitle ?? 'Task',
      webUrl: '',
      isFolder: false,
      submittedBy: this.todo.selectedSubmitter ?? undefined,
      eFormDetails: {
        eFormListId: this.todo.selectedEFormListId,
        type: this.todo.selectedEFormTitle,
        category: 'eForm',
        status: '',
      },
      isContentLoaded: true,
    };
  }
}
