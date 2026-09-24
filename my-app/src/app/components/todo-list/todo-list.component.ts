import { Component, OnInit, OnDestroy, Output, EventEmitter, Input } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Subscription } from 'rxjs';
import { TodoService } from '../../services/todo.service';
import { UserService } from '../../services/user.service';
import { ToDoTask } from '../../types/todo-task.interface';
import { DomainUser, SubordinaryTaskService } from '../../services/subordinaryTask.service';
import { NewCommentCompleteComponent } from '../new-comment/complete/complete';
import { ChangeTaskDateComponent } from './change task date/date';
import { ClaimComponent } from '../Claimable Task/claim';

export interface SubmitterSelection {
  name: string;
  email?: string;
  eFormListId?: string;
  eFormTitle?: string;
  taskId?: string;
  taskIds?: string[];
}

export interface SubmitterGroup{
  displayName: any;
  submitterKey: string;
  submitterName: string;
  email?: string;
  subGroups: { groupKey: string; eFormListId: string; taskLabel: string; tasks: ToDoTask[] }[];
  totalCount: number;
}

export interface LocalToDoTask {
  id: string;
  name: string;
  description: string;
  status: 'pending' | 'in_progress' | 'completed' | 'rejected' | 'APPROVED';
  assignedTo: string;
  listName: string;
  submittedDate: string;
  submittedBy: string;
  statusDate: string;
  createdDate: Date;
  updatedDate: Date;
  eFormDetails?: any;
  webUrl?: string;
}

@Component({
  selector: 'app-todo-list',
  standalone: true,
  imports: [CommonModule, NewCommentCompleteComponent, ChangeTaskDateComponent, ClaimComponent],
  templateUrl: './todo-list.component.html',
  styleUrls: ['./todo-list.component.css']
})


export class TodoListComponent implements OnInit, OnDestroy {
  /**
   * GROUPING OVERVIEW (UI)
   * - Level 1 (accordion): submitter / employee name
   * - Level 2 (nested list): task types (grouped by eFormListId + type/title)
   *
   * Implementation detail:
   * Tasks are grouped into a flat map keyed by `submitter:eFormListId:eFormTitle`.
   */
  private static readonly GROUP_KEY_SEP = ':'; // separator between submitter and eFormListId in each group key

  /** True while AppComponent is still fetching HR tasks for To Do. */
  @Input() loadingTasks = false;

  /** True only while the user-triggered "Load remaining tasks" request is in flight. */
  @Input() loadingMoreTasks = false;

  get isTodoLoading(): boolean {
    return this.isLoading || this.loadingTasks || (this.showSubordinateTasks && !!this.subordinateTasksLoading);
  }

  tasks: ToDoTask[] = [];
  searchTerm = '';
  private tasksSubscription: Subscription | null = null;
  private loadingSubscription: Subscription | null = null;
  isLoading = false;
  expandedUsers: Set<string> = new Set();
  currentUser: any;
  submitter: string = '';
  subordinateTasks: ToDoTask[] = [];
  showSubordinateTasks = false;
  showFutureTasks = false;
  subordinateTasksLoading = false;
  subordinateTasksError: string | null = null;
  subordinateTasksLoaded = false;
  subordinateTasksLoadedError = false;
  subordinateTasksLoadedErrorMsg = '';
  subordinateEmails: any;
  isLoadingSubordinates: boolean | undefined;
  subordinateEmployees: DomainUser[] = [];
  private subordinateEmployeesLoadSeq = 0;
  constructor(
    private todoService: TodoService,
    private userService: UserService,
    private subordinaryTaskService: SubordinaryTaskService
  ) { }

  /**
   * Cache for `activeGroupedTasks` because grouping is called from template getters
   * and can be expensive (especially subordinate mode).
   */
  private groupedTasksCache: { [submitterName: string]: ToDoTask[] } | null = null;
  private groupedTasksCacheKey = '';

  // Toggles between "tasks assigned to me" (default) and "tasks where I am
  // Superior 1 or Superior 2".
  onSubordinateToggle(checked: boolean): void {
    this.showSubordinateTasks = checked;
    this.showFutureTasks = false
    this.expandedSubmitters.clear();
    this.expandedSubmitterGroups.clear();
    this.invalidateGroupedTasksCache();
    this.submitterSelected.emit(null);
    this.superiorModeChange.emit(checked);
    if (checked) {
      // Load subordinate employees (cached after first load) so the UI can group
      // subordinate tasks without needing a manual refresh.
      void this.loadSubordinateEmployees(false);
      return;
    }
    this.subordinateEmployees = [];
    this.subordinateTasksError = null;
  }

  onFutureTasksToggle(checked: boolean): void {
    const wasSubordinate = this.showSubordinateTasks;
    this.showFutureTasks = checked;
    this.showSubordinateTasks = false;
    this.expandedSubmitters.clear();
    this.expandedSubmitterGroups.clear();
    this.invalidateGroupedTasksCache();
    this.submitterSelected.emit(null);
    // Leaving subordinate mode notifies the parent; future-only is a local filter.
    if (wasSubordinate) {
      this.superiorModeChange.emit(false);
    }
  }

  // Grouped tasks by user (tasks assigned to current user)
  get groupedTasks(): { [userName: string]: ToDoTask[] } {
    const groups: { [userName: string]: ToDoTask[] } = {};
    const currentUserEmail = this.userService.getCurrentUserEmail();
    const currentUserName = this.userService.getCurrentUserName();

    const relevantTasks = this.tasks.filter(task => {
      const assignedToUser = this.isAssignedToCurrentUser(task.assignedTo, currentUserName, currentUserEmail);
      const submittedByUser = task.submittedBy === currentUserName;
      return assignedToUser || submittedByUser;
    });

    relevantTasks.forEach(task => {
      // Never show the literal "Unknown User" (and don't render blanks as groups).
      let userName = String(task.submittedBy ?? '').trim();
      if (!userName) userName = '__unknown__';

      if (!groups[userName]) {
        groups[userName] = [];
      }
      groups[userName].push(task);
    });
    return groups;
  }

  get userNames(): string[] {
    return Object.keys(this.groupedTasks);
  }


  expandedSubmitters: Set<string> = new Set();
  @Output() submitterSelected = new EventEmitter<SubmitterSelection | null>();
  @Output() superiorModeChange = new EventEmitter<boolean>();
  @Output() taskCompleted = new EventEmitter<{ task: ToDoTask; completionText: string }>();
  @Output() taskDueDateChanged = new EventEmitter<{ task: ToDoTask; dueDate: string }>();

  /** Level-1 accordion state (only one submitter group expanded at a time). */
  expandedSubmitterGroups: Set<string> = new Set();

  toggleSubmitterGroup(submitterKey: string): void {
    if (this.expandedSubmitterGroups.has(submitterKey)) {
      this.expandedSubmitterGroups.delete(submitterKey);
      this.expandedSubmitters.clear();
      this.submitterSelected.emit(null);
    } else {
      this.expandedSubmitterGroups.clear(); // accordion: only one submitter open at a time
      this.expandedSubmitterGroups.add(submitterKey);
      this.expandedSubmitters.clear();
      this.submitterSelected.emit(null);
    }
  }
  
  isSubmitterGroupExpanded(submitterKey: string): boolean {
    return this.expandedSubmitterGroups.has(submitterKey);
  }

  trackBySubmitterGroup = (_: number, group: SubmitterGroup): string => group.submitterKey;

  trackBySubGroup = (_: number, sub: { groupKey: string }): string => sub.groupKey;
  //this make the asign to field to display more then 1 user name or email assigned to the same task
  private isAssignedToCurrentUser(assignedTo: string | undefined | null, currentUserName: string, currentUserEmail: string): boolean {
    return this.userService.matchesAssigneeField(assignedTo, currentUserName, currentUserEmail);
  }

  // My tasks — each row is one submitter + one eForm List ID.
  private buildTasksAssignedToMe(): { [submitterName: string]: ToDoTask[] } {
    const groups: { [submitterName: string]: ToDoTask[] } = {};
    const futureCutoff = this.getFutureTasksCutoffDate();
    for (const task of this.tasks as any[]) {
      if (!this.isAssignedToMeTodoStatus(task)) continue;

      const isMine = this.isTaskAssignedToCurrentUserTask(task);
      if (!isMine) continue;

      // Future toggle: due after today+14 days. Default list hides those far-future tasks.
      const isFuture = this.isFutureDueTask(task, futureCutoff);
      if (this.showFutureTasks ? !isFuture : isFuture) continue;

      let submitter = String(task.submittedBy ?? '').trim();
      if (!submitter) submitter = '__unknown__';
      const eFormListId = String(task.eFormDetails?.eFormListId ?? '').trim() || 'Unknown eForm';
      const eFormTitle = this.getTaskEFormTitle(task);
      const key = this.buildGroupKey(submitter, eFormListId, eFormTitle);

      if (!groups[key]) groups[key] = [];
      groups[key].push(task);
    }
  
    return groups;
  }

  /** Calendar date (local) that is today + 14 days, as `YYYY-MM-DD`. */
  private getFutureTasksCutoffDate(): string {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 14);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  /** `dd/MM/yyyy` cutoff shown on the Future Tasks toggle (today + 14 days). */
  get futureTasksCutoffDisplay(): string {
    const iso = this.getFutureTasksCutoffDate();
    const [y, m, d] = iso.split('-');
    if (!y || !m || !d) return '';
    return `${d}/${m}/${y}`;
  }

  /** True when the task’s due date is strictly after today + 14 days (e.g. after 21/08/2026 today). */
  private isFutureDueTask(task: ToDoTask | null | undefined, cutoffIso?: string): boolean {
    const due = this.getTaskDueDate(task);
    if (!due) return false;
    return due > (cutoffIso ?? this.getFutureTasksCutoffDate());
  }

  /** Pending and in-progress tasks assigned to the logged-in user. */
  private isAssignedToMeTodoStatus(task: any): boolean {
    const rawFields = (task?.eFormDetails?.rawFields ?? {}) as Record<string, unknown>;
    const progress = String(rawFields['Progress'] ?? task?.eFormDetails?.progress ?? '').toLowerCase().trim();
    // Live To Do for Proc* is gated on Progress, not LastState/Status.
    if (progress === 'pending') return true;
    if (progress === 'complete' || progress === 'completed') return false;

    const { rawStatus, normalizedStatus, fromSharePoint } = this.getTodoStatusSignals(task);
    const combined = `${normalizedStatus} ${rawStatus} ${fromSharePoint}`.toLowerCase();

    if (combined.includes('reject') || combined.includes('denied')) {
      return false;
    }

    if (this.isPendingApprovalStatus(combined)) {
      return true;
    }

    if (
      normalizedStatus === 'completed' ||
      normalizedStatus === 'rejected' ||
      this.isCompletedOrApprovedStatus(combined)
    ) {
      return false;
    }

    if (this.isRequestForActionTask(task)) {
      return true;
    }

    if (
      normalizedStatus === 'pending' ||
      normalizedStatus === 'in_progress' ||
      combined.includes('pending') ||
      combined.includes('progress') ||
      combined.includes('open') ||
      combined.includes('assign') ||
      combined.includes('await') ||
      combined.includes('action')
    ) {
      return true;
    }

    if (!combined.trim()) {
      return true;
    }

    return false;
  }

  /** Workflow states like "Pending Approval" must remain in To Do. */
  private isPendingApprovalStatus(status: string): boolean {
    return /\b(pending|awaiting|waiting|needs?)\b.*\bapprov/.test(status)
      || /\bawait(ing|s)?\s+approv/.test(status);
  }

  private isCompletedOrApprovedStatus(status: string): boolean {
    if (this.isPendingApprovalStatus(status)) return false;
    if (status.includes('complet')) return true;
    if (/\bapproved\b/.test(status) && !/\b(unapproved|not approved)\b/.test(status)) return true;
    return status.includes('approv');
  }

  /** Prefer displayed status; fall back to SharePoint raw fields (status may be cleared for comment cards). */
  private getTodoStatusSignals(task: any): {
    rawStatus: string;
    normalizedStatus: string;
    fromSharePoint: string;
  } {
    const rawFields = (task?.eFormDetails?.rawFields ?? {}) as Record<string, unknown>;
    const fromSharePoint = String(
      rawFields['Status'] ??
      rawFields['LastState'] ??
      rawFields['ApprovalStatus'] ??
      rawFields['WorkflowStatus'] ??
      ''
    ).toLowerCase();
    const rawStatus = String(
      task.eFormDetails?.status || task.status || fromSharePoint || '',
    ).toLowerCase();
    const normalizedStatus = String(task.status ?? '').toLowerCase();
    return { rawStatus, normalizedStatus, fromSharePoint };
  }

  // Subordinate tasks — same grouping by eForm List ID as above.
  private buildTasksAsSuperior(): { [submitterName: string]: ToDoTask[] } {
    const groups: { [submitterName: string]: ToDoTask[] } = {};
    const employees = this.activeSubordinateEmployees;

    const superiorTasks = this.tasks.filter((task: any) => {
      if (!this.isActiveSubordinateTask(task)) {
        return false;
      }

      const assigneeValues = this.getSubordinateTaskAssigneeValues(task);
      if (!assigneeValues.length) {
        return false;
      }

      return this.subordinaryTaskService.isTaskAssignedToAnyAssigneeValue(assigneeValues);
    });

    superiorTasks.forEach(task => {
      const assigneeValues = this.getSubordinateTaskAssigneeValues(task);
      if (!assigneeValues.length) return;

      // A task may be assigned to multiple subordinates; show it under each matching employee.
      const matchedEmployees = employees.filter(user =>
        assigneeValues.some(value =>
          this.subordinaryTaskService.isTaskAssignedToEmployeeForDomainUser(value, user)
        )
      );
      if (!matchedEmployees.length) return;

      const eFormListId = this.getTaskEFormListId(task);
      const eFormTitle = this.getTaskEFormTitle(task);

      for (const employee of matchedEmployees) {
        /**
         * IMPORTANT:
         * Use a *stable unique key* for the "submitter" portion when in subordinate mode.
         * Full names can collide (same-name employees) which causes groups to overwrite each other
         * and makes rows appear/disappear as data refreshes.
         */
        // Internal key can be a placeholder, but we never display "Unknown User".
        const submitterKey = String(employee.ADmail || employee.PinNo || employee.FullName || '__unknown__').trim();
        const groupKey = this.buildGroupKey(submitterKey, eFormListId, eFormTitle);

        if (!groups[groupKey]) {
          groups[groupKey] = [];
        }
        groups[groupKey].push(task);
      }
    });

    return groups;
  }

  get subordinateEmployeeCount(): number {
    return this.activeSubordinateEmployees.length;
  }

  private get activeSubordinateEmployees(): DomainUser[] {
    return this.subordinateEmployees.length > 0
      ? this.subordinateEmployees
      : this.subordinaryTaskService.getCachedSubordinates();
  }

  private getPrimaryTaskAssignee(task: any): string {
    const details = task.eFormDetails ?? {};
    const raw = String(details.assignedTo ?? task.assignedTo ?? '').trim();
    return this.userService.parseAssignees(raw)[0] ?? '';
  }

  private getTaskAssigneeValues(task: any): string[] {
    const details = task.eFormDetails ?? {};
    const rawValues = [details.assignedTo, task.assignedTo]
      .map(value => String(value ?? '').trim())
      .filter(Boolean);

    return [...new Set(
      rawValues.flatMap(value => this.userService.parseAssignees(value))
    )];
  }

  private getSubordinateTaskAssigneeValues(task: any): string[] {
    const primary = this.getPrimaryTaskAssignee(task);
    return primary ? [primary] : [];
  }

  /** True when the logged-in user is on the task assignee field(s). */
  private isTaskAssignedToCurrentUserTask(task: any): boolean {
    if (task.isAssignedToCurrentUser === true) {
      return true;
    }

    const currentUserEmail = this.userService.getCurrentUserEmail();
    const currentUserName = this.userService.getCurrentUserName();
    const assigneeValues = this.getTaskAssigneeValues(task);

    if (!assigneeValues.length) {
      return this.isAssignedToCurrentUser(task.assignedTo, currentUserName, currentUserEmail);
    }

    return assigneeValues.some(value =>
      this.isAssignedToCurrentUser(value, currentUserName, currentUserEmail)
    );
  }

  private isActiveSubordinateTask(task: any): boolean {
    const { rawStatus, normalizedStatus, fromSharePoint } = this.getTodoStatusSignals(task);
    const status = `${rawStatus} ${fromSharePoint}`.trim();
    if (
      status.includes('approv') ||
      status.includes('complet') ||
      normalizedStatus === 'completed'
    ) {
      return false;
    }
    if (
      normalizedStatus === 'pending' ||
      normalizedStatus === 'rejected' ||
      normalizedStatus === 'in_progress'
    ) {
      return true;
    }
    return (
      status.includes('pending') ||
      status.includes('progress') ||
      status.includes('reject') ||
      (this.isRequestForActionTask(task) && !fromSharePoint)
    );
  }

  // Picks which grouped map the template renders (normal vs subordinate mode).
  get activeGroupedTasks(): { [submitterName: string]: ToDoTask[] } {
    const cacheKey = [
      this.showSubordinateTasks ? '1' : '0',
      this.showFutureTasks ? '1' : '0',
      this.tasks.length,
      this.subordinateEmployees.length,
      this.subordinaryTaskService.getCachedSubordinates().length,
    ].join('|');

    if (this.groupedTasksCache && this.groupedTasksCacheKey === cacheKey) {
      return this.groupedTasksCache;
    }

    this.groupedTasksCacheKey = cacheKey;
    // Note: this returns a *flat* map keyed by `submitter:eFormListId`.
    this.groupedTasksCache = this.showSubordinateTasks
      ? this.buildTasksAsSuperior()
      : this.buildTasksAssignedToMe();
    return this.groupedTasksCache;
  }

  private invalidateGroupedTasksCache(): void {
    this.groupedTasksCache = null;
    this.groupedTasksCacheKey = '';
  }

  get submitterNames(): string[] {
    return Object.keys(this.activeGroupedTasks).filter(
      name => this.activeGroupedTasks[name].length > 0
    );
  }

  get filteredSubmitterNames(): string[] {
    if (!this.searchTerm.trim()) return this.submitterNames;
    const query = this.searchTerm.toLowerCase().trim();
    return this.submitterNames.filter(name => {
      if (name.toLowerCase().includes(query)) return true;
      if (this.showSubordinateTasks && this.getEmployeeFullName(name).toLowerCase().includes(query)) return true;
      return this.activeGroupedTasks[name].some(task => this.taskMatchesSearch(task, query));
    });
  }

  get groupedBySubmitter(): SubmitterGroup[] {
    /**
     * Converts the flat map (`activeGroupedTasks`) into a list of submitter groups:
     * SubmitterGroup = { displayName, email, subGroups[], totalCount }
     *
     * This is what the template renders as:
     * submitter row -> nested "task type" rows.
     */
    const flat = this.activeGroupedTasks; // existing flat map: { "submitter:eFormListId:eFormTitle": tasks[] }
    const bySubmitter = new Map<string, SubmitterGroup>();
    const uniqueTaskIdsBySubmitter = new Map<string, Set<string>>();
  
    for (const groupKey of Object.keys(flat)) {
      const tasks = flat[groupKey];
      if (!tasks.length) continue;
  
      const { submitter, eFormListId } = this.parseGroupKey(groupKey);
  
      if (!bySubmitter.has(submitter)) {
        let displayName = this.showSubordinateTasks
          ? this.getSubordinateDisplayName(groupKey)
          : this.formatSubmitterDisplayName(submitter);

        if (!String(displayName ?? '').trim()) {
          displayName = this.getGroupTaskLabel(groupKey) || 'Assigned task';
        }

        if (!String(displayName ?? '').trim()) {
          continue;
        }

        bySubmitter.set(submitter, {
          submitterKey: submitter,
          displayName,
          email: this.showSubordinateTasks
            ? this.getEmployeeEmail(groupKey)
            : (this.isSubmitterEmail(submitter) ? submitter : undefined),
          subGroups: [],
          totalCount: 0,
          submitterName: ''
        });
        uniqueTaskIdsBySubmitter.set(submitter, new Set<string>());
      }
  
      const entry = bySubmitter.get(submitter);
      if (!entry) continue;
      entry.subGroups.push({
        groupKey,
        eFormListId: eFormListId && eFormListId !== 'Unknown eForm' ? eFormListId : '',
        taskLabel: this.getGroupTaskLabel(groupKey),
        tasks,
      });

      // Count tasks uniquely per submitter (prevents double-counting if upstream list contains duplicates).
      const idSet = uniqueTaskIdsBySubmitter.get(submitter)!;
      for (const task of tasks as any[]) {
        const id = String((task as any)?.id ?? '').trim();
        if (id) idSet.add(id);
      }
    }
  
    for (const [submitter, entry] of bySubmitter.entries()) {
      entry.totalCount = uniqueTaskIdsBySubmitter.get(submitter)?.size ?? 0;
    }

    return [...bySubmitter.values()];
  }
  
  get filteredGroupedBySubmitter(): SubmitterGroup[] {
    const all = this.groupedBySubmitter;
    if (!this.searchTerm.trim()) return all;
    const query = this.searchTerm.toLowerCase().trim();
  
    return all
      .map(group => {
        const nameMatches =
          group.displayName.toLowerCase().includes(query) ||
          (group.email ?? '').toLowerCase().includes(query);
        if (nameMatches) return group;
  
        const filteredSubGroups = group.subGroups.filter(sg =>
          sg.taskLabel.toLowerCase().includes(query) ||
          sg.tasks.some(t => this.taskMatchesSearch(t, query))
        );
        return filteredSubGroups.length ? { ...group, subGroups: filteredSubGroups } : null;
      })
      .filter((g): g is SubmitterGroup => g !== null);
  }
  
  // make this to display only the initials and not numbers
  getGroupInitial(displayName: string): string {
    const v = String(displayName ?? '').trim();
    if (!v || v === '—') return '?';
    return (v.match(/[A-Z]/g)?.join('') || '?').toUpperCase();  }

  onSearchChange(event: Event): void {
    this.searchTasks((event.target as HTMLInputElement).value);
  }

  searchTasks(term: string): void {
    this.searchTerm = term ?? '';
  }



  private taskMatchesSearch(task: ToDoTask, query: string): boolean {
    const extended = task as ToDoTask & {
      listName?: string;
      submittedDate?: string;
      eFormDetails?: Record<string, unknown>;
    };
    const searchableValues = [
      extended.name,
      extended.description,
      extended.status,
      extended.assignedTo,
      extended.listName,
      extended.submittedBy,
      extended.submittedDate,
      ...Object.values(extended.eFormDetails ?? {}),
    ];
    return searchableValues.some(value => String(value ?? '').toLowerCase().includes(query));
  }


  toggleSubmitter(submitterName: string): void {
    if (this.expandedSubmitters.has(submitterName)) {
      this.expandedSubmitters.clear();
      this.submitterSelected.emit(null);
    } else {
      this.expandedSubmitters.clear();
      this.expandedSubmitters.add(submitterName);
      // Pull submitter, id, and title back out of the row key when the user clicks a task group.
      const { submitter, eFormListId, eFormTitle } = this.parseGroupKey(submitterName);
      const groupTasks = this.activeGroupedTasks[submitterName] ?? [];
      const taskIds = groupTasks
        .map(task => String((task as any)?.id ?? '').trim())
        .filter(Boolean);
      const firstTaskId = taskIds[0];
      this.submitterSelected.emit({
        name: submitter,
        eFormListId: eFormListId || undefined,
        eFormTitle: eFormTitle && eFormTitle !== 'Unknown Title' ? eFormTitle : undefined,
        email: this.showSubordinateTasks ? this.getEmployeeEmail(submitterName) : undefined,
        taskId: firstTaskId,
        taskIds,
      });
    }
  }

  isSubmitterExpanded(submitterName: string): boolean {
    return this.expandedSubmitters.has(submitterName);
  }

  ngOnInit(): void {
    this.loadTasks();
  }

  ngOnDestroy(): void {
    this.tasksSubscription?.unsubscribe();
    this.loadingSubscription?.unsubscribe();
  }

  private loadTasks(): void {
    this.isLoading = this.todoService.isLoading();
    this.loadingSubscription = this.todoService.getLoading$().subscribe(loading => {
      this.isLoading = loading;
    });
    this.tasksSubscription = this.todoService.getTasksForCurrentUser().subscribe(
      tasks => {
        this.tasks = tasks;
        this.invalidateGroupedTasksCache();
        if (this.showSubordinateTasks) {
          this.subordinateEmployees = this.subordinaryTaskService.getCachedSubordinates();
        }
        if (!this.todoService.isLoading()) {
          this.isLoading = false;
        }
      },
      () => {
        this.isLoading = false;
      }
    );
  }

  private async loadSubordinateEmployees(forceRefresh = false): Promise<void> {
    const currentUser = this.userService.getCurrentUser();
    const identifiers = [
      currentUser?.email,
      currentUser?.userPrincipalName,
      currentUser?.username,
    ]
      .map(value => (value ?? '').trim())
      .filter(Boolean);

    if (identifiers.length === 0) return;

    const loadSeq = ++this.subordinateEmployeesLoadSeq;
    this.subordinateTasksLoading = true;
    this.subordinateTasksError = null;

    try {
      const employees = await this.subordinaryTaskService.ensureSubordinatesForManager(
        identifiers,
        forceRefresh
      );
      if (loadSeq !== this.subordinateEmployeesLoadSeq || !this.showSubordinateTasks) return;

      this.subordinateEmployees = employees;
      this.invalidateGroupedTasksCache();
      this.subordinateTasksLoaded = true;
    } catch {
      if (loadSeq !== this.subordinateEmployeesLoadSeq || !this.showSubordinateTasks) return;
      this.subordinateTasksError = 'Could not load subordinate employees.';
      this.subordinateEmployees = [];
    } finally {
      if (loadSeq === this.subordinateEmployeesLoadSeq) {
        this.subordinateTasksLoading = false;
      }
    }
  }

  getEmployeeFullName(employeeName: string): string {
    const employee = this.findSubordinateEmployee(employeeName);
    return employee?.FullName ?? '';
  }

  getEmployeeEmail(employeeName: string): string {
    const employee = this.findSubordinateEmployee(employeeName);
    return employee?.ADmail ?? '';
  }

  getSubordinateDisplayName(groupKey: string): string {
    const rawKey = this.parseGroupKey(groupKey).submitter;
    if (rawKey === '__unknown__') return '';
    return this.getEmployeeFullName(groupKey) || this.parseGroupKeyForDisplay(groupKey);
  }

  // Show the eForm List ID on each row in the template.
  getGroupEFormListId(groupKey: string): string {
    const eFormListId = this.parseGroupKey(groupKey).eFormListId;
    return eFormListId && eFormListId !== 'Unknown eForm' ? eFormListId : '';
  }

  /** Human-readable task type for a group, e.g. "Sick Leave" instead of "SickLeaveByAppointment". */
  /*this is the label that will be displayed in the task list and the eform Id*/
  getGroupTaskLabel(groupKey: string): string {
    const tasks = this.activeGroupedTasks[groupKey];
    const title = this.getTaskEFormTitle(tasks?.[0] ?? {});
    if (title && title !== 'Unknown Title') return title;
    return this.showSubordinateTasks ? 'Subordinate pending tasks' : 'Tasks assigned to you';
  }

  private formatTaskType(title: string | undefined): string {
    return (title ?? '')
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
      .trim();
  }
  private getTaskEFormListId(task: { eFormDetails?: { eFormListId?: unknown } }): string {
    return String(task.eFormDetails?.eFormListId ?? '').trim() || 'Unknown eForm';
  }

  /**
   * Title used for Level-2 grouping and the row label.
   * Prefer the SharePoint item Title — `eFormDetails.type` / `name` are the list
   * display name (e.g. "ProcTasks"), not the task title.
   */
  private getTaskEFormTitle(task: { eFormDetails?: Record<string, unknown>; name?: unknown }): string {
    const details = (task as any)?.eFormDetails ?? {};
    const rawFields = (details?.rawFields ?? {}) as Record<string, unknown>;
    const raw =
      rawFields['Title'] ??
      details['Title'] ??
      details['title'];
    return String(raw ?? '').trim() || 'Unknown Title';
  }


  private buildGroupKey(submitter: string, eFormListId: string, eFormTitle: string): string {
    // Key format: submitter:eFormListId:eFormTitle
    return [submitter, eFormListId, eFormTitle].join(TodoListComponent.GROUP_KEY_SEP);
  }

  private parseGroupKey(groupKey: string): { submitter: string; eFormListId: string; eFormTitle: string } {
    const sep = TodoListComponent.GROUP_KEY_SEP;
    const parts = String(groupKey ?? '').split(sep);
    if (parts.length <= 1) return { submitter: groupKey, eFormListId: '', eFormTitle: '' };
    const submitter = parts[0] ?? '';
    const eFormListId = parts[1] ?? '';
    const eFormTitle = parts.slice(2).join(sep); // preserve any separators in title
    return { submitter, eFormListId, eFormTitle };
  }
  // ── END: Group tasks by eForm List ID ────────────────────────────

  private getEmployeeDisplayName(employee: DomainUser): string {
    return employee.FullName || employee.ADmail || 'Unknown User';
  }

  /** Finds a subordinate employee by email/full name/pin (case-insensitive). */
  private findSubordinateEmployee(groupKeyOrSubmitter: string): DomainUser | undefined {
    const submitter = this.parseGroupKey(groupKeyOrSubmitter).submitter;
    const target = String(submitter ?? '').trim().toLowerCase();
    if (!target) return undefined;

    return this.activeSubordinateEmployees.find(employee => {
      const email = String(employee.ADmail ?? '').trim().toLowerCase();
      const fullName = String(employee.FullName ?? '').trim().toLowerCase();
      const pin = String(employee.PinNo ?? '').trim().toLowerCase();
      return (email && email === target) || (fullName && fullName === target) || (pin && pin === target);
    });
  }

  private isTaskAssignedToEmployee(assignedTo: string | undefined | null, employee: DomainUser): boolean {
    return this.userService.matchesAssigneeField(assignedTo, employee.FullName, employee.ADmail) ||
      this.userService.matchesAssigneeField(assignedTo, employee.ADmail, employee.ADmail) ||
      (employee.PinNo
        ? this.userService.matchesAssigneeField(assignedTo, employee.PinNo, employee.ADmail)
        : false);
  }

  onCompleteTask(task: ToDoTask, event: { completionText: string }): void {
    this.taskCompleted.emit({
      task,
      completionText: String(event?.completionText ?? '').trim(),
    });
  }

  /** Keep local subgroup task in sync after a successful SharePoint write. */
  onChangeTaskDate(task: ToDoTask, event: { dueDate: string }): void {
    const dueDate = String(event?.dueDate ?? '').trim();
    if (!task || !dueDate) return;

    const t = task as ToDoTask & { dueDate?: string; eFormDetails?: Record<string, unknown> };
    const details = { ...(t.eFormDetails ?? {}) };
    const rawFields = { ...((details['rawFields'] as Record<string, unknown> | undefined) ?? {}) };
    details['dueDate'] = dueDate;

    let wroteRaw = false;
    for (const key of ['DueDate', 'DueDate0', 'Due', 'Due_x0020_Date'] as const) {
      if (Object.prototype.hasOwnProperty.call(rawFields, key)) {
        rawFields[key] = dueDate;
        wroteRaw = true;
        break;

    
      }
    }
    if (!wroteRaw) rawFields['DueDate'] = dueDate;
    details['rawFields'] = rawFields;

    t.dueDate = dueDate;
    t.eFormDetails = details;
    this.invalidateGroupedTasksCache();
    this.taskDueDateChanged.emit({ task: t, dueDate });
  }

  /** Due date shown on the level-2 card (SharePoint DueDate / local override). Only display the clock if the tasks has duedate column in sharepoint*/
  getTaskDueDate(task: ToDoTask | null | undefined): string | null {
    if (!task) return null;
    const details = (task as any)?.eFormDetails ?? {};
    const raw = String(
      details?.dueDate ??
        (task as any)?.dueDate ??
        details?.DueDate ??
        details?.rawFields?.DueDate ??
        details?.rawFields?.DueDate0 ??
        details?.rawFields?.Due ??
        details?.rawFields?.Due_x0020_Date ??
        '',
    ).trim();
    if (!raw) return null;
    const iso = raw.match(/^(\d{4}-\d{2}-\d{2})/);
    if (iso) return iso[1];
    const dt = new Date(raw);
    if (Number.isNaN(dt.getTime())) return null;
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;

  }

  /** `dd/MM/yyyy` for the level-2 card — avoids UTC date-pipe day shifts. */
  getTaskDueDateDisplay(task: ToDoTask | null | undefined): string | null {
    const iso = this.getTaskDueDate(task);
    if (!iso) return null;
    const [y, m, d] = iso.split('-');
    if (!y || !m || !d) return null;
    return `${d}/${m}/${y}`;
  }

  hasDueDateColumn(task: ToDoTask | null | undefined): boolean {
    const details = (task as any)?.eFormDetails ?? {};
    const raw = details.rawFields ?? {};
    const hasKey = ['DueDate', 'DueDate0', 'Due', 'Due_x0020_Date']
      .some(key => Object.prototype.hasOwnProperty.call(raw, key));
    if (hasKey) return true;
    return !!String(details.dueDate ?? (task as any)?.dueDate ?? '').trim();
  }

  /** Graph omits empty columns from `fields`, so scan the whole subgroup. */
  subGroupHasDueDateColumn(sub: { tasks?: ToDoTask[] } | null | undefined): boolean {
    return (sub?.tasks ?? []).some(task => this.hasDueDateColumn(task));
  }

  /** TEMP testing: show Claim when SharePoint Claimable is No/false. Flip back to Yes/true later. */
  isClaimableTask(task: ToDoTask | null | undefined): boolean {
    if (!task || !this.isTaskAssignedToCurrentUserTask(task)) return false;

    const value = this.getClaimableValue(task);
    if (value == null) return false;
    if (typeof value === 'boolean') return value === false;
    const s = String(value).trim().toLowerCase();
    return s === 'no' || s === 'false' || s === '0';
  }

  private getClaimableValue(task: ToDoTask | null | undefined): unknown {
    const details = (task as any)?.eFormDetails ?? {};
    const raw = (details.rawFields ?? {}) as Record<string, unknown>;
    const sources = [raw, details];

    for (const source of sources) {
      for (const [key, value] of Object.entries(source)) {
        if (key.replace(/[^a-z0-9]/gi, '').toLowerCase() === 'claimable') {
          return value;
        }
      }
    }
    return null;
  }

  /** True only for "Request for Action" comment/action entries. */
  isRequestForActionTask(task: unknown): boolean {
    const t: any = task as any;
    const details: any = t?.eFormDetails ?? {};

    // UI uses eFormDetails.category (e.g. "Request for Action")
    const category = String(details?.category ?? '')
      .trim()
      .toLowerCase();
    if (category === 'request for action') {
      return true;
    }

    // Newer saves: commentCategory === 'action'
    const commentCategory = String(details?.commentCategory ?? '')
      .trim()
      .toLowerCase();
    if (commentCategory === 'action') {
      return true;
    }

    // Older saves: Title explicitly set to "Request for Action"
    const rawFields = (details?.rawFields ?? {}) as Record<string, unknown>;
    const title = String(rawFields?.['Title'] ?? '')
      .trim()
      .toLowerCase();
    return title === 'request for action' || title.startsWith('request for action:');
  }
  //end of request for action complete button logic

  onRejectTask(taskId: string): void {
    this.todoService.rejectTask(taskId);
  }

  getStatusColor(status: ToDoTask['status']): string {
    switch (status) {
      case 'pending':
        return '#ffc107';
      case 'in_progress':
        return '#17a2b8';
      case 'completed':
        return '#28a745';
      case 'rejected':
        return '#dc3545';
      default:
        return '#6c757d';
    }
  }

  formatDate(date: Date): string {
    return new Date(date).toLocaleDateString();
  }

  toggleUserExpansion(userName: string): void {
    if (this.expandedUsers.has(userName)) {
      this.expandedUsers.delete(userName);
    } else {
      this.expandedUsers.add(userName);
    }
  }

  isUserExpanded(userName: string): boolean {
    return this.expandedUsers.has(userName);
  }

  isTaskAssignedToCurrentUser(task: ToDoTask): boolean {
    return this.isTaskAssignedToCurrentUserTask(task);
  }

  /** Submitter portion of a group key (safe for templates). */
  parseGroupKeyForDisplay(groupKey: string): string {
    return this.parseGroupKey(groupKey ?? '').submitter;
  }

  formatSubmitterDisplayName(submitter: string): string {
    const value = this.parseGroupKey(submitter ?? '').submitter.trim();
    if (value === '__unknown__') return '';
    if (!value.includes('@')) return value;
    const local = value.split('@')[0] ?? value;
    return local
      .split(/[._-]+/)
      .filter(Boolean)
      .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
      .join(' ');
  }

  getSubmitterInitial(submitter: string): string {
    return (this.formatSubmitterDisplayName(submitter).charAt(0) || '?').toUpperCase();
  }

  isSubmitterEmail(submitter: string): boolean {
    return this.parseGroupKey(submitter ?? '').submitter.includes('@');
  }
}