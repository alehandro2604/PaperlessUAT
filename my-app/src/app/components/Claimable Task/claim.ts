/**
 * Claimable Task Component
 * Displays a button to claim a task and a modal to confirm the claim.
 */

import { Overlay, OverlayModule, OverlayRef } from '@angular/cdk/overlay';
import { PortalModule, TemplatePortal } from '@angular/cdk/portal';
import { CommonModule } from '@angular/common';
import {
  Component,
  Input,
  OnDestroy,
  TemplateRef,
  ViewChild,
  ViewContainerRef,
  ViewEncapsulation,
} from '@angular/core';
import { ToDoTask } from '../../types/todo-task.interface';

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

  @ViewChild('claimDialog', { static: true }) private claimDialogTpl?: TemplateRef<unknown>;

  private overlayRef: OverlayRef | null = null;

  constructor(
    private readonly overlay: Overlay,
    private readonly viewContainerRef: ViewContainerRef,
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

    if (!this.overlayRef.hasAttached()) {
      this.overlayRef.attach(new TemplatePortal(this.claimDialogTpl, this.viewContainerRef));
    }
  }

  closeModal(event?: Event): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.overlayRef?.detach();
  }

  claimTask(event?: Event): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.closeModal();
  }

  ngOnDestroy(): void {
    this.overlayRef?.dispose();
    this.overlayRef = null;
  }
}
