// ============================================================
// MANUAL REFRESH COMPONENT
// ============================================================

import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
    selector: 'app-manual-refresh',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './refresh.html',  
    styleUrls: ['./refresh.css'],
})
export class ManualRefreshComponent {
    @Input() isLoading = false;
    @Output() refreshRequested = new EventEmitter<void>();

    //refresh tasks and attachments
    refreshTasks(): void {
        this.isLoading = true;
        setTimeout(() => {
            this.isLoading = false;
        }, 1000);
        this.refreshRequested.emit();
        this.isLoading = false;
    }

    refreshAttachments(): void {
        this.isLoading = true;
        setTimeout(() => {
            this.isLoading = false;
        }, 1000);
        this.refreshRequested.emit();
    }
}
