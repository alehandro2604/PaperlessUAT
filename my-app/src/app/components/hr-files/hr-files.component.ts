import { CommonModule } from '@angular/common';
import { Component, EventEmitter, Input, OnChanges, Output, SimpleChanges } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { AppDropdownComponent, AppDropdownOption } from '../app-dropdown/app-dropdown.component';
import {
  filterAndSortHrFileItems,
  formatDisplayNameFromEmail,
  getDisplayInitial,
  HrFileListItem,
  HrFilesSortDirection,
} from './hr-files.utils';

@Component({
  selector: 'app-hr-files',
  standalone: true,
  imports: [CommonModule, FormsModule, AppDropdownComponent],
  templateUrl: './hr-files.component.html',
  styleUrls: ['./hr-files.component.css'],
})
export class HrFilesComponent implements OnChanges {
  @Input() items: HrFileListItem[] = [];
  @Input() isLoading = false;
  @Input() progressMessage = '';
  @Input() errorMessage = '';
  @Input() warningMessage = '';

  @Output() itemSelected = new EventEmitter<HrFileListItem>();
  @Output() loadRequested = new EventEmitter<void>();

  protected searchQuery = '';
  protected sortDirection: HrFilesSortDirection = 'ascending';
  protected selectedItemId: string | null = null;
  protected visibleCount = 50;
  protected readonly pageSize = 100;

  protected readonly sortOptions: AppDropdownOption[] = [
    { value: 'ascending', label: 'Ascending' },
    { value: 'descending', label: 'Descending' },
  ];

  protected readonly formatDisplayNameFromEmail = formatDisplayNameFromEmail;
  protected readonly getDisplayInitial = getDisplayInitial;

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['items']) {
      this.resetPagination();
    }
  }

  protected get hasLoaded(): boolean {
    return this.items.length > 0;
  }

  protected get filteredItems(): HrFileListItem[] {
    return filterAndSortHrFileItems(this.items, this.searchQuery, this.sortDirection);
  }

  protected get pagedItems(): HrFileListItem[] {
    return this.filteredItems.slice(0, this.visibleCount);
  }

  protected get hasMore(): boolean {
    return this.filteredItems.length > this.visibleCount;
  }

  protected get remainingCount(): number {
    return Math.max(this.filteredItems.length - this.visibleCount, 0);
  }

  protected onSearchChange(event: Event): void {
    this.searchQuery = (event.target as HTMLInputElement).value;
    this.resetPagination();
  }

  protected onSortChange(value: string): void {
    this.sortDirection = value as HrFilesSortDirection;
    this.resetPagination();
  }

  protected loadMore(): void {
    this.visibleCount += this.pageSize;
  }

  protected onItemClick(item: HrFileListItem): void {
    this.selectedItemId = item.id;
    this.itemSelected.emit(item);
  }

  protected isItemSelected(item: HrFileListItem): boolean {
    return this.selectedItemId === item.id;
  }

  private resetPagination(): void {
    this.visibleCount = this.pageSize;
  }
}
