import { ChangeDetectorRef, Component, EventEmitter, Input, OnChanges, OnDestroy, Output, SimpleChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { AppDropdownComponent, AppDropdownOption } from '../app-dropdown/app-dropdown.component';
import { DelegateService, DomainUser } from '../../services/delegate.service';
import { TodoService } from '../../services/todo.service';

@Component({
  selector: 'app-new-comment',
  standalone: true,
  imports: [CommonModule, FormsModule, AppDropdownComponent],
  templateUrl: './new-comment.html',
  styleUrls: ['./new-comment.css'],
})

export class NewCommentComponent implements OnChanges, OnDestroy {
  @Input() open = false;
  @Input() taskId!: number | string; // SharePoint list item id this comment belongs to
  @Input() listName = '';
  @Input() eFormListId = '';
  @Input() eFormTitle = '';
  /** When false, hides "Request for Action" (only SAP workflow rows with SAPOrder). */
  @Input() showRequestForAction = false;
  @Output() closed = new EventEmitter<void>();
  @Output() saved = new EventEmitter<{
    taskId: number | string;
    category: string;
    comment: string;
    fileName?: string;
    attachmentUrl?: string;
    sharePointItemId?: string;
    assignedToName?: string;
    assignedToEmail?: string;
  }>();

  protected selectedCategory = '';
  protected commentText = '';
  protected selectedFile: File | null = null;
  protected saving = false;
  protected saveError = '';
  protected actionSearchQuery = '';
  protected actionRightPanel: 'attachments' | 'assign' = 'attachments';
  protected actionSearchResults: DomainUser[] = [];
  protected selectedActionUser: DomainUser | null = null;
  protected actionSearching = false;
  protected actionSearchError = '';

  private attachmentObjectUrl: string | null = null;
  private actionSearchTimeout: ReturnType<typeof setTimeout> | null = null;
  private readonly ACTION_SEARCH_DEBOUNCE_MS = 300;

  protected get isAttachmentCategory(): boolean {
    return this.selectedCategory === 'attachment';
  }

  protected get isRequestForActionCategory(): boolean {
    return this.selectedCategory === 'action';
  }

  protected get isSplitLayout(): boolean {
    return this.isAttachmentCategory || this.isRequestForActionCategory;
  }

  protected get canSave(): boolean {
    if (this.saving || !this.selectedCategory) {
      return false;
    }

    const hasComment = !!this.commentText.trim();

    if (this.isAttachmentCategory) {
      return !!this.selectedFile;
    }

    if (this.isRequestForActionCategory) {
      return hasComment || !!this.selectedActionUser || !!this.selectedFile;
    }

    return hasComment;
  }

  protected get isActionAttachmentSave(): boolean {
    return this.isRequestForActionCategory && !!this.selectedFile;
  }


  protected get categoryOptions(): AppDropdownOption[] {
    const options: AppDropdownOption[] = [
      { value: 'general', label: 'General Comment' },
      { value: 'attachment', label: 'New Attachment' },
    ];

    if (this.showRequestForAction) {
      options.push({ value: 'action', label: 'Request for Action' });
    }

    return options;
  }

  constructor(
    private todoService: TodoService,
    private delegateService: DelegateService,
    private cdr: ChangeDetectorRef,
  ) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['open']?.currentValue === true) {
      this.resetForm();
    }

    if (changes['showRequestForAction']?.currentValue === false && this.selectedCategory === 'action') {
      this.selectedCategory = '';
      this.resetActionSearch();
      this.actionRightPanel = 'attachments';
    }
  }

  ngOnDestroy(): void {
    this.clearActionSearchTimeout();
    this.revokeAttachmentObjectUrl();
  }

  protected onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.revokeAttachmentObjectUrl();
    this.selectedFile = input.files?.[0] ?? null;
    if (this.selectedFile) {
      this.attachmentObjectUrl = URL.createObjectURL(this.selectedFile);
    }
  }

  close(): void {
    this.closed.emit();
  }

  protected onActionSearchInput(): void {
    this.clearActionSearchTimeout();
    this.selectedActionUser = null;
    this.actionSearchError = '';

    const query = this.actionSearchQuery.trim();
    if (query.length < 2) {
      this.actionSearchResults = [];
      this.actionSearching = false;
      return;
    }

    this.actionSearchTimeout = setTimeout(() => {
      void this.runActionSearch(query);
    }, this.ACTION_SEARCH_DEBOUNCE_MS);
  }

  protected selectActionUser(user: DomainUser): void {
    this.selectedActionUser = user;
    this.actionSearchQuery = user.Title;
    this.actionSearchResults = [];
    this.actionSearchError = '';
  }

  protected showActionAttachments(): void {
    this.actionRightPanel = 'attachments';
  }

  protected showActionAssignTo(): void {
    this.actionRightPanel = 'assign';
  }

  protected async saveComment(): Promise<void> {
    if (!this.taskId) {
      console.error('No task ID provided — cannot save comment');
      return;
    }
    if (!this.listName.trim()) {
      console.error('No list name available — select a Level 2 task group first');
      return;
    }
    if (!this.eFormListId.trim()) {
      console.error('No eForm list ID available — select a Level 2 task group first');
      return;
    }
    if (!this.hasMinimumSaveContent()) {
      return;
    }
    if (this.isAttachmentCategory && !this.selectedFile) {
      return;
    }
    if (this.isRequestForActionCategory && !this.commentText.trim() && !this.selectedActionUser && !this.selectedFile) {
      return;
    }

    const commentToSave = this.buildCommentText();
    const includeAttachment = this.isAttachmentCategory || this.isActionAttachmentSave;

    this.saving = true;
    this.saveError = '';
    try {
      const created = await this.todoService.addComment(this.taskId, {
        comment: commentToSave,
        category: this.selectedCategory,
        listName: this.listName,
        eFormListId: this.eFormListId,
        eFormTitle: this.eFormTitle,
        isAttachment: includeAttachment,
        attachmentFile: includeAttachment ? this.selectedFile : null,
        attachmentFileName: this.selectedFile?.name ?? '',
        assignedToUser: this.isRequestForActionCategory ? this.selectedActionUser : null,
      });
      this.saved.emit({
        taskId: this.taskId,
        category: this.selectedCategory,
        comment: commentToSave,
        fileName: created.attachmentFileName ?? this.selectedFile?.name ?? '',
        attachmentUrl: created.attachmentUrl ?? '',
        sharePointItemId: created.id,
        assignedToName: created.assignedToUser?.Title ?? this.selectedActionUser?.Title ?? '',
        assignedToEmail:
          created.assignedToUser?.Email ??
          created.assignedToUser?.UserPrincipalName ??
          this.selectedActionUser?.Email ??
          this.selectedActionUser?.UserPrincipalName ??
          '',
      });
      this.attachmentObjectUrl = null;
      this.close();
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to save comment';
      this.saveError = message;
      console.error('Failed to save comment', err);
    } finally {
      this.saving = false;
      // App runs zoneless: without this a failed save stays stuck on "saving" with no error shown.
      this.cdr.markForCheck();
    }
  }

  private resetForm(): void {
    this.commentText = '';
    this.selectedCategory = '';
    this.selectedFile = null;
    this.saveError = '';
    this.actionRightPanel = 'attachments';
    this.resetActionSearch();
    this.revokeAttachmentObjectUrl();
  }

  private hasMinimumSaveContent(): boolean {
    const hasComment = !!this.commentText.trim();

    if (this.isAttachmentCategory) {
      return hasComment || !!this.selectedFile;
    }

    if (this.isRequestForActionCategory) {
      return hasComment || !!this.selectedActionUser || !!this.selectedFile;
    }

    return hasComment;
  }

  private buildCommentText(): string {
    const trimmedComment = this.commentText.trim();
    if (!this.isRequestForActionCategory || !this.selectedActionUser) {
      return trimmedComment;
    }

    return trimmedComment;
  }

  private async runActionSearch(query: string): Promise<void> {
    this.actionSearching = true;
    this.actionSearchError = '';

    try {
      this.actionSearchResults = await this.delegateService.searchUsers(query);
    } catch {
      this.actionSearchError = 'Failed to search users. Please try again.';
      this.actionSearchResults = [];
    } finally {
      this.actionSearching = false;
      // App runs zoneless, so this async result needs an explicit render.
      this.cdr.markForCheck();
    }
  }

  private resetActionSearch(): void {
    this.clearActionSearchTimeout();
    this.actionSearchQuery = '';
    this.actionSearchResults = [];
    this.selectedActionUser = null;
    this.actionSearching = false;
    this.actionSearchError = '';
  }

  private clearActionSearchTimeout(): void {
    if (!this.actionSearchTimeout) return;
    clearTimeout(this.actionSearchTimeout);
    this.actionSearchTimeout = null;
  }

  private revokeAttachmentObjectUrl(): void {
    if (!this.attachmentObjectUrl) return;
    URL.revokeObjectURL(this.attachmentObjectUrl);
    this.attachmentObjectUrl = null;
  }

}
