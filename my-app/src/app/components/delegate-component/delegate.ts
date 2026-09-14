import { ChangeDetectorRef, Component, EventEmitter, Input, OnDestroy, OnInit, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Subscription } from 'rxjs';
import { DelegateService, DomainUser } from '../../services/delegate.service';
import { UserService } from '../../services/user.service';

// Define these here or import from your shared models
export interface Task {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  size?: number;
  isFolder?: boolean;
  status?: string;
  submittedBy?: string;
  showFullContent?: boolean;
  title?: string;
  assignedTo?: string;
  eFormDetails?: {
    assignedTo?: string;
    status?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

// Type guard to check if object matches Task structure
export function isTask(obj: any): obj is Task {
  return obj && typeof obj.id === 'string' && typeof obj.name === 'string';
}

@Component({
  selector: 'app-delegate-component',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './delegate.html',
  styleUrls: ['./delegate.css'],
})
export class DelegateComponent implements OnInit, OnDestroy {
  @Input() task?: Task;

  // Emits the task + the new assignee so the parent can update SharePoint
  @Output() delegateTask = new EventEmitter<{ task: Task; newAssignee: DomainUser }>();

  canDelegate = false;
  showForm = false;
  private canDelegateSubscription: Subscription | null = null;

  selectedUser: DomainUser | null = null;
  searchQuery = '';
  searchResults: DomainUser[] = [];
  isSearching = false;
  searchError = '';
  private searchTimeout: any = null;
  private searchSeq = 0;// guards against a slower earlier search overwriting a newer one
  private readonly DEBOUNCE_MS = 300;// debounce time for the search input

  constructor(
    private delegateService: DelegateService,
    private userService: UserService,
    private cdr: ChangeDetectorRef,
  ) { }

  ngOnInit(): void {
    this.canDelegateSubscription = this.delegateService.canDelegate$.subscribe(
      canDelegate => (this.canDelegate = canDelegate),
    );

    if (!this.delegateService.isLoaded) {
      void this.delegateService.loadDelegates();
    }
  }

  ngOnDestroy(): void {
    this.canDelegateSubscription?.unsubscribe();
    if (this.searchTimeout) {
      clearTimeout(this.searchTimeout);
    }
  }

  onDelegateClick(): void {
    this.showForm = !this.showForm;
  }

  closePopup(): void {
    this.showForm = false;
    this.resetSearch();
  }

  onSubmit(): void {
    console.log('=== onSubmit called ===', { selectedUser: this.selectedUser, task: this.task });
    if (this.selectedUser && this.task) {
      console.log('Emitting delegateTask event to parent');
      this.delegateTask.emit({ task: this.task, newAssignee: this.selectedUser });
    } else {
      console.warn('onSubmit blocked: missing selectedUser or task', {
        selectedUser: this.selectedUser,
        task: this.task,
      });
    }
    this.showForm = false;
    this.resetSearch();
  }

  canShowDelegateButton(): boolean {
    const status = (this.task?.eFormDetails?.status || this.task?.status || '').trim().toLowerCase();
    const assignedTo =
      this.task?.assignedTo ||
      this.task?.eFormDetails?.assignedTo;

    // Only show if status is pending and NOT approved/completed
    const isPending = status === 'pending';
    const isCompleted = status.includes('approved') || status.includes('complete') || status.includes('rejected');

    // Check if task is assigned to current user
    const isAssignedToCurrentUser = this.userService.matchesAssigneeField(assignedTo);

    return this.canDelegate && isPending && !isCompleted && isAssignedToCurrentUser;
  }

  onSearchInput(): void {
    if (this.searchTimeout) {
      clearTimeout(this.searchTimeout);
    }

    const query = this.searchQuery.trim();// trim the search query to remove whitespace

    if (query.length < 2) {
      this.searchResults = [];
      this.searchError = '';
      this.isSearching = false;
      return;
    }

    this.isSearching = true;
    this.searchError = '';

    const seq = ++this.searchSeq;//this is a sequence number to track the search request

    this.searchTimeout = setTimeout(async () => {
      try {
        const results = await this.delegateService.searchUsers(query);
        if (seq !== this.searchSeq) return;
        this.searchResults = results;
      } catch {
        if (seq !== this.searchSeq) return;
        this.searchError = 'Failed to search users. Please try again.';
        this.searchResults = [];
      } finally {
        if (seq === this.searchSeq) {
          this.isSearching = false;
          // App runs zoneless, so this async result needs an explicit render.
          this.cdr.detectChanges();//cdr means Change Detection Reference,this will trigger a change detection cycle to update the UI
        }
      }
    }, this.DEBOUNCE_MS);
  }

  selectUser(user: DomainUser): void {
    this.selectedUser = user;
    this.searchQuery = user.Title;
    this.searchResults = [];
  }

  private resetSearch(): void {
    this.searchQuery = '';
    this.searchResults = [];
    this.selectedUser = null;
    this.searchError = '';
  }
}