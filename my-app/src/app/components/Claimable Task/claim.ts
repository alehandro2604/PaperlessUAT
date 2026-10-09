/**
 * Claimable Task Component
 * Displays a button to claim a task and a modal to confirm the claim.
 * Claiming makes the current user the only assignee of a task that is
 * shared between several people (same as delegating the task to yourself).
 */

import { Overlay, OverlayModule, OverlayRef } from '@angular/cdk/overlay';
import { PortalModule, TemplatePortal } from '@angular/cdk/portal';
import { CommonModule } from '@angular/common';
import {
  ChangeDetectorRef,
  Component,
  EventEmitter,
  Input,
  OnDestroy,
  Output,
  TemplateRef,
  ViewChild,
  ViewContainerRef,
  ViewEncapsulation,
} from '@angular/core';
import { ToDoTask } from '../../types/todo-task.interface';
import { DelegateService } from '../../services/delegate.service';
import { UserService } from '../../services/user.service';

export interface TaskClaimedEvent {
  task: ToDoTask;
  name: string;
  email: string;
}

@Component({
  selector: 'app-claim',
  standalone: true,
  imports: [CommonModule, OverlayModule, PortalModule],
  templateUrl: './claim.html',
  styleUrls: ['./claim.css'],
  encapsulation: ViewEncapsulation.None,
})
export class ClaimComponent implements OnDestroy {
  @Input() task: ToDoTask | null = null;

  /** Emitted only after SharePoint confirms the current user is the sole assignee. */
  @Output() claimed = new EventEmitter<TaskClaimedEvent>();

  @ViewChild('claimDialog', { static: true }) private claimDialogTpl?: TemplateRef<unknown>;

  isClaiming = false;
  errorMessage = '';
  /** True once SharePoint accepted the claim; the popup then shows the success message. */
  isClaimed = false;

  private overlayRef: OverlayRef | null = null;
  /** Held until the success popup is closed, so the list does not remove this component (and its popup) early. */
  private pendingClaim: TaskClaimedEvent | null = null;

  constructor(
    private readonly overlay: Overlay,
    private readonly viewContainerRef: ViewContainerRef,
    private readonly delegateService: DelegateService,
    private readonly userService: UserService,
    private readonly cdr: ChangeDetectorRef,
  ) {}

  openModal(event?: Event): void {
    event?.preventDefault();
    event?.stopPropagation();
    if (!this.claimDialogTpl) return;

    if (!this.overlayRef) {
      this.overlayRef = this.overlay.create({
        hasBackdrop: true,
        backdropClass: 'claim-modal-backdrop',
        panelClass: 'claim-modal-panel',
        scrollStrategy: this.overlay.scrollStrategies.block(),
        positionStrategy: this.overlay.position().global().centerHorizontally().centerVertically(),
      });

      this.overlayRef.backdropClick().subscribe(() => this.closeModal());
      this.overlayRef.keydownEvents().subscribe((ev: KeyboardEvent) => {
        if (ev.key === 'Escape') this.closeModal();
      });
    }

    this.errorMessage = '';
    if (!this.overlayRef.hasAttached()) {
      this.overlayRef.attach(new TemplatePortal(this.claimDialogTpl, this.viewContainerRef));
    }
  }

  closeModal(event?: Event): void {
    event?.preventDefault();
    event?.stopPropagation();
    // Keep the modal open while the claim is being saved.
    if (this.isClaiming) return;
    this.overlayRef?.detach();

    if (this.pendingClaim) {
      const claim = this.pendingClaim;
      this.pendingClaim = null;
      this.claimed.emit(claim);
    }
  }

  async claimTask(event?: Event): Promise<void> {
    event?.preventDefault();
    event?.stopPropagation();
    if (!this.task || this.isClaiming || this.isClaimed) return;

    const email = this.userService.getCurrentUserEmail().trim();
    const name = this.userService.getCurrentUserName().trim();
    if (!email) {
      this.errorMessage = 'Could not find your email address. Please sign in again.';
      this.cdr.markForCheck();
      return;
    }

    this.isClaiming = true;
    this.errorMessage = '';
    this.cdr.markForCheck();

    try {
      // Claiming = replacing all assignees with the current user. The service
      // reads the item back and throws unless the current user is the only assignee.
      await this.delegateService.updateAssignedTo(
        this.task as any,
        { Title: name, Email: email, UserPrincipalName: email },
        { requireSoleAssignee: true },
      );

      this.isClaiming = false;
      this.isClaimed = true;
      this.pendingClaim = { task: this.task, name, email };
    } catch (err) {
      console.error('Failed to claim task:', err);
      this.isClaiming = false;
      this.errorMessage = err instanceof Error ? err.message : 'Could not claim the task.';
    } finally {
      // The app runs zoneless, so redraw after the await.
      this.cdr.markForCheck();
    }
  }

  ngOnDestroy(): void {
    this.overlayRef?.dispose();
    this.overlayRef = null;
  }
}