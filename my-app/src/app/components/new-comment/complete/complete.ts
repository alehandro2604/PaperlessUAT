import { CommonModule } from '@angular/common';
import { ChangeDetectorRef, Component, EventEmitter, Input, OnDestroy, Output, TemplateRef, ViewChild, ViewContainerRef, ViewEncapsulation } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Overlay, OverlayModule, OverlayRef } from '@angular/cdk/overlay';
import { PortalModule, TemplatePortal } from '@angular/cdk/portal';

export interface CompleteTaskEvent {
  completionText: string;
  /** The parent calls this when the SharePoint save ends: no argument on success, an error message on failure. */
  done: (error?: string | null) => void;
}

@Component({
  selector: 'app-new-comment-complete',
  standalone: true,
  imports: [CommonModule, FormsModule, OverlayModule, PortalModule],
  templateUrl: './complete.html',
  styleUrls: ['./complete.css'],
  encapsulation: ViewEncapsulation.None,
})
export class NewCommentCompleteComponent implements OnDestroy {
  @Input() task: any;
  @Input() isRequestForActionTask: boolean = false;

  @Output() completeTask = new EventEmitter<CompleteTaskEvent>();

  @ViewChild('completeDialog', { static: true }) private completeDialogTpl?: TemplateRef<unknown>;

  completeText = '';
  isSubmitting = false;
  errorMessage = '';

  private overlayRef: OverlayRef | null = null;

  constructor(
    private readonly overlay: Overlay,
    private readonly viewContainerRef: ViewContainerRef,
    private readonly cdr: ChangeDetectorRef,
  ) { }

  onCompleteClick(event: MouseEvent): void {
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
    this.errorMessage = '';
    if (!this.overlayRef.hasAttached()) {
      this.overlayRef.attach(new TemplatePortal(this.completeDialogTpl, this.viewContainerRef));
    }
  }

  closeModal(): void {
    // Keep the popup open while SharePoint is saving.
    if (this.isSubmitting) return;
    this.overlayRef?.detach();
    this.completeText = '';
    this.errorMessage = '';
  }

  onSaveClick(): void {
    if (this.isSubmitting) return;
    this.isSubmitting = true;
    this.errorMessage = '';
    this.completeTask.emit({
      completionText: this.completeText.trim(),
      // Close only once SharePoint confirms; on failure keep the text and show why.
      done: (error?: string | null) => {
        this.isSubmitting = false;
        if (error) {
          this.errorMessage = error;
        } else {
          this.closeModal();
        }
        // The app runs zoneless, so redraw after the async save.
        this.cdr.markForCheck();
      },
    });
  }

  ngOnDestroy(): void {
    // A completed task leaves the list, which destroys this component; drop its popup and backdrop too.
    this.overlayRef?.dispose();
    this.overlayRef = null;
  }
}
