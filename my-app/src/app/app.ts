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
  protected readonly comments = inject(CommentsStore);
  protected readonly folderTasks = inject(FolderTasksStore);
  protected readonly todo = inject(TodoStore);
  protected readonly attachments = inject(AttachmentsStore);
  protected readonly hr = inject(HrFilesStore);
  protected readonly allFiles = inject(AllFilesStore);
  private readonly formConfigService = inject(FormConfigurationService);
  private readonly taskService = inject(TaskService);
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
