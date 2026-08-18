import { CommonModule } from '@angular/common';
import { Component, EventEmitter, Input, Output, TemplateRef, ViewChild, ViewContainerRef, ViewEncapsulation } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Overlay, OverlayModule, OverlayRef } from '@angular/cdk/overlay';
import { PortalModule, TemplatePortal } from '@angular/cdk/portal';

@Component({
  selector: 'app-new-comment-complete',
  standalone: true,
  imports: [CommonModule, FormsModule, OverlayModule, PortalModule],
  templateUrl: './complete.html',
  styleUrls: ['./complete.css'],
  encapsulation: ViewEncapsulation.None,
})
export class NewCommentCompleteComponent {
  @Input() task: any;
  @Input() isRequestForActionTask: boolean = false;

  @Output() completeTask = new EventEmitter<{ completionText: string }>();

  @ViewChild('completeDialog', { static: true }) private completeDialogTpl?: TemplateRef<unknown>;

  completeText = '';
  isSubmitting = false;

  private overlayRef: OverlayRef | null = null;

  constructor(
    private readonly overlay: Overlay,
    private readonly viewContainerRef: ViewContainerRef,
  ) { }

  onCompleteClick(event: MouseEvent): void {
    console.log('[NewCommentCompleteComponent] onCompleteClick', {
      taskId: this.task?.id ?? null,
      isRequestForActionTask: this.isRequestForActionTask,
      hasTemplate: !!this.completeDialogTpl,
      overlayExists: !!this.overlayRef,
      overlayHasAttached: this.overlayRef?.hasAttached?.() ?? false,
      eventTarget: (event?.target as HTMLElement | null)?.tagName ?? null,
    });
    event.preventDefault();
    event.stopPropagation();
    this.openModal();
  }

  openModal(): void {
    if (!this.completeDialogTpl) return;
    if (!this.overlayRef) {
      this.overlayRef = this.overlay.create({
        hasBackdrop: true,
        backdropClass: 'complete-modal-backdrop',
        panelClass: 'complete-modal-panel',
        scrollStrategy: this.overlay.scrollStrategies.block(),
        positionStrategy: this.overlay.position().global().centerHorizontally().centerVertically(),
      });

      this.overlayRef.backdropClick().subscribe(() => this.closeModal());
      this.overlayRef.keydownEvents().subscribe((ev: KeyboardEvent) => {
        if (ev.key === 'Escape') this.closeModal();
      });
    }

    this.completeText = '';
    if (!this.overlayRef.hasAttached()) {
      this.overlayRef.attach(new TemplatePortal(this.completeDialogTpl, this.viewContainerRef));
    }
  }

  closeModal(): void {
    this.overlayRef?.detach();
    this.completeText = '';
  }

  onSaveClick(): void {
    if (this.isSubmitting) return;
    this.isSubmitting = true;
    try {
      this.completeTask.emit({ completionText: this.completeText.trim() });
      this.closeModal();
    } finally {
      this.isSubmitting = false;
    }
  }
}
