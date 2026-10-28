import { ChangeDetectorRef, Component, HostListener, OnDestroy, OnInit, ViewChild, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { FormConfigurationService } from './services/form-configuration.service';
import { MobileNavigationComponent } from './components/mobile navigation/mobile-navigation';
import { TodoListComponent } from './components/todo-list/todo-list.component';
import { AllFilesComponent } from './components/All-files/all-files';
import { TaskService } from './services/task.service';
import { PowerappsModalComponent } from './components/powerApps/powerapps-modal.component';
import { NewCommentComponent } from './components/new-comment/new-comment';
import { LoadingScreenComponent } from './components/loading screen/loading-screen';
import { DelegateComponent } from './components/delegate-component/delegate';
import { ManualRefreshComponent } from './components/manual-refresh/refresh';
import { AppDropdownComponent } from './components/app-dropdown/app-dropdown.component';
import { HrFilesComponent } from './components/hr-files/hr-files.component';
import { AttachmentList } from './components/attachments/attachment-list/attachment-list';
import { PreviewAttachments } from './components/preview-attachments/preview';
import { ViewRefreshService } from './stores/view-refresh.service';
import { UiStore } from './stores/ui.store';
import { SessionStore } from './stores/session.store';
import { CommentsStore } from './stores/comments.store';
import { FolderTasksStore } from './stores/folder-tasks.store';
import { TodoStore } from './stores/todo.store';
import { AttachmentsStore } from './stores/attachments.store';
import { HrFilesStore } from './stores/hr-files.store';
import { AllFilesStore } from './stores/all-files.store';
import { ALL_FILES_LIBRARY_RULES, ALL_FILES_PLUS_ALLOWED_EMAILS, ALLOWED_LIBRARIES_FOR_ALL_USERS } from './hr-task-lists.config';
import { TodoService } from './services/todo.service';

/**
 * Root component: the page layout (app.html) plus wiring.
 *
 * All behaviour lives in the feature stores under ./stores (one per area of the
 * screen). This class only hands those stores to the template, connects the two
 * child components the stores call into, and starts the app (ngOnInit).
 */
@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, TodoListComponent, AllFilesComponent, PowerappsModalComponent, NewCommentComponent, LoadingScreenComponent, MobileNavigationComponent, DelegateComponent, ManualRefreshComponent, AppDropdownComponent, HrFilesComponent, AttachmentList, PreviewAttachments],
  templateUrl: './app.html',
  styleUrls: ['./app.css', './components/attachments/attachments.css', './form-list.css', './components/mobile navigation/mobile-navigation.css', './components/task-details/task-details.css', './components/comments/comment-card.css'],
})
export class AppComponent implements OnInit, OnDestroy {
  // Feature stores - the template reads state and calls methods through these.
  protected readonly view = inject(ViewRefreshService);
  protected readonly ui = inject(UiStore);
  protected readonly session = inject(SessionStore);

  get canSeeAllFilesPlus(): boolean {
    const email = this.session.currentUser?.email?.toLowerCase();
    return !!email && ALL_FILES_PLUS_ALLOWED_EMAILS.some(e => e.toLowerCase() === email);
  }

  /** All Files: the task the new-comment panel is attached to (To Do uses the selected task). */
  allFilesCommentTarget: { taskId: string | number; listName: string; eFormListId: string; eFormTitle: string } | null = null;

  /** To Do: open the new-comment panel. All Files: open it for the open folder's task. */
  onPlusClick(): void {
    if (!this.ui.showAllFilesSection) {
      this.comments.openNewCommentModal();
      return;
    }
    const item: any = this.comments.commentItems.find(i => i.id && (i as any).eFormDetails?.listName)
      ?? this.comments.commentItems.find(i => i.id);
    if (!item) {
      console.error('Select a folder with a task first');
      return;
    }
    const d = item.eFormDetails ?? {};
    this.allFilesCommentTarget = {
      taskId: item.id,
      listName: String(item.listName ?? d.listName ?? '').trim(),
      eFormListId: String(d.eFormListId ?? '').trim(),
      eFormTitle: String(d.type ?? d.eFormTitle ?? '').trim(),
    };
    this.comments.isCommentPanelOpen = true;
  }

  onNewCommentClosed(): void {
    this.allFilesCommentTarget = null;
    this.comments.closeNewCommentModal();
  }

  /** All Files "+": mirrors the Power Apps DisplayMode rules for the open folder. */
  get canAddOnAllFiles(): boolean {
    const reason = this.allFilesPlusReason();
    const logKey = `${this.attachments.currentLibraryName}|${this.allFiles.selectedFolderName}|${reason}`;
    if (logKey !== this.lastPlusReason) {
      this.lastPlusReason = logKey;
      console.debug(`[All Files +] ${this.attachments.currentLibraryName ?? '-'} / ${this.allFiles.selectedFolderName} -> ${reason ?? 'denied'}`);
    }
    return reason !== null;
  }

  private lastPlusReason: string | undefined;

  /** Which rule allows the "+" for the open folder, or null when none does. */
  private allFilesPlusReason(): string | null {
    const libraryName = this.attachments.currentLibraryName;
    if (!libraryName) return null; // nothing selected yet
    const form = this.formConfigService.getFormByDocLibrary(libraryName);
    const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
    const keys = [libraryName, form?.title ?? ''].map(norm).filter(Boolean);
    // Exact match on the signed-in user's email or display name (no fuzzy name matching).
    const isMe = (person: string): boolean => {
      const p = norm(person);
      return !!p && (p === norm(this.session.currentUser?.email ?? '') || p === norm(this.session.currentUser?.username ?? ''));
    };

    // Rules 5, 7, 8: libraries open to everyone
    if (ALLOWED_LIBRARIES_FOR_ALL_USERS.some(n => norm(n) === norm(libraryName))) return 'open library';
    // Rule 13: allow-list
    if (this.canSeeAllFilesPlus) return 'allow-listed email';
    // Rule 4: SAP workflow
    if (this.comments.commentItems.some(i => this.formConfigService.isSapOrderWorkflow(i.eFormDetails?.eFormListId, i.eFormDetails?.type, (i as any).listName ?? i.eFormDetails?.listName))) return 'SAP workflow';

    // Rule 1: assigned to the user, or in the To Do list (own + subordinates' tasks)
    const folder = norm(this.allFiles.selectedFolderName);
    if (isMe(this.allFiles.selectedFolderAssignedTo)) return `assigned to me (${this.allFiles.selectedFolderAssignedTo})`;
    if (folder && this.todoService.currentTasks.some(t => (t.status === 'pending' || t.status === 'in_progress') && norm(String(t.name ?? '')) === folder)) return 'folder is in my To Do';

    // Rules 2, 3, 6, 9-12: by library - group membership and/or the person who opened the file
    const rule = ALL_FILES_LIBRARY_RULES.find(r => keys.some(k => k.includes(r.libraryKey)));
    if (!rule) return null;
    if (rule.groupId && this.session.groupIds.has(rule.groupId)) return `group ${rule.libraryKey}`;
    return rule.allowOpener && isMe(this.allFiles.selectedFolderCreatedBy)
      ? `opened by me (${this.allFiles.selectedFolderCreatedBy})`
      : null;
  }
  


  protected readonly comments = inject(CommentsStore);
  protected readonly folderTasks = inject(FolderTasksStore);
  protected readonly todo = inject(TodoStore);
  protected readonly attachments = inject(AttachmentsStore);
  protected readonly hr = inject(HrFilesStore);
  protected readonly allFiles = inject(AllFilesStore);
  private readonly formConfigService = inject(FormConfigurationService);
  private readonly taskService = inject(TaskService);
  private readonly todoService = inject(TodoService);
  private readonly cdr = inject(ChangeDetectorRef);

  // Child components some features call into (e.g. refresh All Files' folder list).
  @ViewChild(AllFilesComponent) set allFilesComponentRef(c: AllFilesComponent | undefined) {
    this.ui.allFilesComponent = c;
  }
  @ViewChild(TodoListComponent) set todoListComponentRef(c: TodoListComponent | undefined) {
    this.ui.todoListComponent = c;
  }

  constructor() {
    // Let every feature redraw the screen through ViewRefreshService.
    this.view.attach(() => this.cdr.detectChanges());
  }

  /** Close the menu when the user clicks anywhere else on the page or presses Escape. */
  @HostListener('document:click')
  @HostListener('document:keydown.escape')
  protected closeUserMenu(): void {
    this.ui.isUserMenuOpen = false;
  }

  async ngOnInit(): Promise<void> {
    this.comments.taskFilterForms$ = this.formConfigService.hrForms$;
    this.allFiles.documentLibraryTaskMapSub = this.formConfigService.documentLibraryTaskMap$
      .subscribe((map) => { this.allFiles.documentLibraryTaskMap = map ?? {}; });
    await this.session.checkLoginOnStart();
    try {
      const tasks = await this.taskService.loadTaskLists();
      this.attachments.allowedHrTaskLists = tasks;
    } catch (error) {
      console.warn('Could not load SharePoint task lists.', error);
      this.attachments.allowedHrTaskLists = [];
    }
  }

  ngOnDestroy(): void {
    this.allFiles.documentLibraryTaskMapSub?.unsubscribe();
    this.session.stopSharePointCachePoll();
    if (this.attachments.folderPrefetchTimer) {
      clearTimeout(this.attachments.folderPrefetchTimer);
      this.attachments.folderPrefetchTimer = null;
    }
    if (this.folderTasks.hrTaskPrefetchTimer) {
      clearTimeout(this.folderTasks.hrTaskPrefetchTimer);
      this.folderTasks.hrTaskPrefetchTimer = null;
    }
  }
}
