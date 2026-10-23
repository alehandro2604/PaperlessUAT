import { Injectable, Injector, inject } from '@angular/core';
import { sharePointConfig } from '../sharepoint.config';
import { MobileCardView, MOBILE_NAV_INDEX, getAttachmentsTabMobileState, getCommentsTabMobileState, resolveMobileNavClick } from '../components/mobile navigation/mobile-navigation';
import { syncMobileSearchTerm } from '../components/mobile navigation/mobile-search.utils';
import { TodoListComponent } from '../components/todo-list/todo-list.component';
import { AllFilesComponent } from '../components/All-files/all-files';
import { formatFormTitle } from '../form-utils';
import { ViewRefreshService } from './view-refresh.service';
import { SessionStore } from './session.store';
import { CommentsStore } from './comments.store';
import { TodoStore } from './todo.store';
import { AttachmentsStore } from './attachments.store';
import { HrFilesStore } from './hr-files.store';
import { AllFilesStore } from './all-files.store';

/**
 * Screen layout: which section/tab is showing, mobile navigation, the user menu, modals and the manual refresh button.
 *
 * Moved out of AppComponent (app.ts). Other features are reached through the lazy
 * getters below (injector.get at call time), so features that use each other do
 * not create a circular dependency when Angular builds them.
 */
@Injectable({ providedIn: 'root' })
export class UiStore {
  private readonly injector = inject(Injector);


  // Other features
  private get view(): ViewRefreshService { return this.injector.get(ViewRefreshService); }
  private get session(): SessionStore { return this.injector.get(SessionStore); }
  private get comments(): CommentsStore { return this.injector.get(CommentsStore); }
  private get todo(): TodoStore { return this.injector.get(TodoStore); }
  private get attachments(): AttachmentsStore { return this.injector.get(AttachmentsStore); }
  private get hr(): HrFilesStore { return this.injector.get(HrFilesStore); }
  private get allFiles(): AllFilesStore { return this.injector.get(AllFilesStore); }

  /** Set by AppComponent through @ViewChild once the child components exist. */
  allFilesComponent?: AllFilesComponent;
  todoListComponent?: TodoListComponent;

  /** Top-bar name; says "UAT" when running against the UAT SharePoint site. */
  readonly appSystemName = sharePointConfig.label;

  /** Initials for the signed-in user's avatar in the top bar. */
  getUserInitials(name: string | null | undefined): string {
    const base = String(name ?? '').split('@')[0].replace(/[._-]+/g, ' ');
    const initials = base.split(/\s+/).filter(Boolean).map(w => w[0]).join('');
    return (initials || '?').slice(0, 3).toUpperCase();
  }

  isUserMenuOpen = false;

  toggleUserMenu(event: Event): void {
    event.stopPropagation();
    this.isUserMenuOpen = !this.isUserMenuOpen;
  }

  isConnecting = false;

  statusMessage = 'Signing-in...';

  errorMessage = '';

  showToDoSection = true;

  showAllFilesSection = false;

  /** Once true, All Files stays in the DOM (hidden) so listings/search caches survive tab switches. */
  allFilesMounted = false;

  searchTerm = '';

  hideComments = true;

  mobileCardView: MobileCardView = 'todo';

  /** Turns a PascalCase/camelCase eForm title (e.g. "SickLeaveByAppointment") into spaced words ("Sick Leave By Appointment"). */
  formatFormTitle(title: string | null | undefined): string {
    return formatFormTitle(title);
  }

  syncMobileSearchInput(): void {
    this.searchTerm = syncMobileSearchTerm({
      showToDoSection: this.showToDoSection,
      showAllFilesSection: this.showAllFilesSection,
      searchQuery: this.comments.searchQuery,
      todoSearchTerm: this.todoListComponent?.searchTerm,
      allFilesSearchTerm: this.allFilesComponent?.searchTerm,
    });
  }

  // MANUAL REFRESH (refresh button) - reloads only the active section.
  onManualRefresh(): void {
    if (!this.session.currentUser) {
      this.showModal('Login Required', 'You must log in first to refresh.', 'error');
      return;
    }

    if (this.showAllFilesSection) {
      this.allFiles.refreshAllFilesSection();
      return;
    }

    if (this.showToDoSection) {
      this.todo.refreshToDoSection();
      return;
    }

    this.hr.refreshHrSection();
  }

  get isManualRefreshLoading(): boolean {
    return (
      this.comments.isLoadingComments ||
      this.attachments.isLoadingUserFiles ||
      this.todo.isLoadingTodoTasks ||
      this.hr.isLoadingHrFilesList ||
      !!this.allFilesComponent?.isLoading ||
      !!this.allFilesComponent?.isLoadingAllLibraries ||
      !!this.allFilesComponent?.isLoadingCommentTasks ||
      !!this.allFilesComponent?.isSearchingMoreAttachments
    );
  }

  modal = {
    visible: false,
    title: '',
    message: '',
    type: 'info' as 'info' | 'success' | 'warning' | 'error',
    action: null,
  };

  showModal(
    title: string,
    message: string,
    type: 'info' | 'success' | 'warning' | 'error' = 'info',
  ): void {
    this.modal = { visible: true, title, message, type, action: null };
    this.view.refreshView();
  }

  closeModal(): void {
    this.modal.visible = false;
    this.modal.action = null;
    this.view.refreshView();
  }

  showPowerAppsModal = false;

  openPowerAppsModal(): void {
    this.showPowerAppsModal = true;
  }

  closePowerAppsModal(): void {
    this.showPowerAppsModal = false;
  }

  activeNavIndex: number = MOBILE_NAV_INDEX.todo;

  get mobileSectionTitle(): string {
    if (this.showAllFilesSection) return 'All Files';
    if (this.showToDoSection) return 'To Do';
    return 'HR';
  }

  get mobileDetailTitle(): string {
    return (
      this.allFiles.selectedFolderName ||
      this.todo.selectedSubmitter ||
      this.attachments.currentFolderName ||
      this.todo.selectedEFormTitle ||
      'Details'
    );
  }

  /** Leave Comments / Attachments and return to the list the user came from. */
  closeMobileDetail(): void {
    this.attachments.closePreview();
    if (this.showAllFilesSection) {
      this.mobileCardView = 'allFiles';
      this.activeNavIndex = MOBILE_NAV_INDEX.allFiles;
    } else if (this.showToDoSection) {
      this.mobileCardView = 'todo';
      this.activeNavIndex = MOBILE_NAV_INDEX.todo;
    } else {
      this.mobileCardView = 'task';
      this.activeNavIndex = MOBILE_NAV_INDEX.task;
    }
    this.view.refreshView();
  }

  showMobileComments(): void {
    this.attachments.closePreview();
    this.hideComments = false;
    this.mobileCardView = 'comments';
    this.syncMobileSearchInput();
    this.view.refreshView();
  }

  showMobileAttachments(): void {
    this.mobileCardView = 'attachments';
    if (
      this.attachments.currentLibraryDriveId ||
      this.attachments.attachmentBrowsingRoot ||
      this.attachments.currentFolderId ||
      (this.attachments.showUserFile && this.attachments.userFiles.length > 0) ||
      // To Do task already filled the panel (even if empty: "No attachments folder...").
      // Don't fall through to My Files, which would show the user's personal folders.
      (this.showToDoSection && this.attachments.showUserFile)
    ) {
      this.attachments.showUserFile = true;
      this.view.refreshView();
      return;
    }
    if (this.session.currentUser?.email) {
      if (this.attachments.hasLoadedPersonalFiles) {
        this.hr.showMyFilesPanel();
      } else {
        this.hr.openMyFiles();
      }
    }
    this.view.refreshView();
  }

  onNavItemClick(index: number): void {
    const onDetail =
      this.mobileCardView === 'comments' || this.mobileCardView === 'attachments';
    if (onDetail && index === this.activeNavIndex) {
      this.closeMobileDetail();
      return;
    }

    const action = resolveMobileNavClick(
      index,
      this.activeNavIndex,
      !!this.session.currentUser?.email
    );

    switch (action.type) {
      case 'none':
        return;
      case 'goHome':
        this.hr.toggleUserInfo();
        return;
      case 'goTodo':
        this.todo.toggleToDoSection();
        return;
      case 'goAllFiles':
        this.allFiles.toggleAllFilesSection();
        return;
      case 'goComments':
        this.hideComments = false;
        Object.assign(this, getCommentsTabMobileState());
        if (action.syncSearch) {
          this.syncMobileSearchInput();
        }
        this.view.refreshView();
        return;
      case 'goAttachments':
        Object.assign(this, getAttachmentsTabMobileState());
        // Already browsing All Files / HR person attachments - only switch the tab.
        // Never call openMyFiles() here: that clears the folder and reloads tasks.
        if (
          this.attachments.currentLibraryDriveId ||
          this.attachments.attachmentBrowsingRoot ||
          this.attachments.currentFolderId ||
          (this.attachments.showUserFile && this.attachments.userFiles.length > 0)
        ) {
          this.attachments.showUserFile = true;
          this.view.refreshView();
          return;
        }
        if (action.shouldLoadFiles) {
          if (this.attachments.hasLoadedPersonalFiles) {
            this.hr.showMyFilesPanel();
          } else {
            this.hr.openMyFiles();
          }
        }
        return;
    }
  }
}
