import { Injectable } from '@angular/core';

/**
 * Batches screen redraws. The app runs without Zone.js, so Angular does not notice
 * plain field changes made after an await / setTimeout. Code calls refreshView()
 * after changing state; many calls in the same tick become one redraw.
 *
 * AppComponent attaches its ChangeDetectorRef once (attach), because only a
 * component can own one. Signals make this unnecessary - see the cleanup docs.
 */
@Injectable({ providedIn: 'root' })
export class ViewRefreshService {
  private detect: (() => void) | null = null;
  private viewUpdatePending = false;

  /** Called once by AppComponent with () => cdr.detectChanges(). */
  attach(detect: () => void): void {
    this.detect = detect;
  }

  refreshView(): void {
    if (this.viewUpdatePending) return;
    this.viewUpdatePending = true;
    queueMicrotask(() => {
      this.viewUpdatePending = false;
      this.detect?.();
    });
  }
}
