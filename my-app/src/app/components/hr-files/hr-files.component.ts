import { CommonModule } from '@angular/common';
import {
  ChangeDetectorRef,
  Component,
  DestroyRef,
  EventEmitter,
  Input,
  OnChanges,
  OnInit,
  Output,
  SimpleChanges,
  inject,
} from '@angular/core';
import { FormsModule } from '@angular/forms';
import { DomainUserService } from '../../services/domain-user.service';
import { AppDropdownComponent, AppDropdownOption } from '../app-dropdown/app-dropdown.component';
import {
  enrichHrFileItems,
  filterAndSortHrFileItems,
  formatDisplayName,
  getDisplayInitial,
  HrDomainUserLookup,
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
export class HrFilesComponent implements OnInit, OnChanges {
  @Input() items: HrFileListItem[] = [];
  @Input() isLoading = false;
  @Input() progressMessage = '';
  @Input() errorMessage = '';
  @Input() warningMessage = '';

  @Output() itemSelected = new EventEmitter<HrFileListItem>();
  @Output() itemPrefetchRequested = new EventEmitter<HrFileListItem>();
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

  private readonly domainUserService = inject(DomainUserService);
  private readonly cdr = inject(ChangeDetectorRef);
  private destroyed = false;
  private domainUsers: HrDomainUserLookup = { byEmail: new Map(), byPin: new Map() };
  /** items + FullName / PinNo from AllDomainUsers (the original `items` are never modified). */
  private enrichedItems: HrFileListItem[] = [];

  protected readonly formatDisplayName = formatDisplayName;
  protected readonly getDisplayInitial = getDisplayInitial;

  constructor() {
    inject(DestroyRef).onDestroy(() => (this.destroyed = true));
  }

  ngOnInit(): void {
    this.loadDomainUsers();
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['items']) {
      // The first load can run before sign-in has finished; try again once folders arrive.
      if (!changes['items'].firstChange && this.items.length && !this.hasDomainUsers) {
        this.loadDomainUsers();
      }
      this.enrichedItems = enrichHrFileItems(this.items, this.domainUsers);
      this.resetPagination();
    }
  }

  private get hasDomainUsers(): boolean {
    return this.domainUsers.byEmail.size > 0 || this.domainUsers.byPin.size > 0;
  }

  private loadDomainUsers(): void {
    // Paint names from the last saved copy first, so the list never waits on SharePoint.
    this.domainUsers = this.domainUserService.peekUsers() ?? this.domainUsers;
    this.enrichedItems = enrichHrFileItems(this.items, this.domainUsers);

    void this.domainUserService.getUsers().then(users => {
      if (this.destroyed || (!users.byEmail.size && !users.byPin.size)) return;
      this.domainUsers = users;
      this.enrichedItems = enrichHrFileItems(this.items, users);
      const unmatched = this.enrichedItems.filter(item => !item.fullName && !item.pinNo);
      if (unmatched.length) {
        console.info(
          `[HrFiles] ${unmatched.length} of ${this.items.length} folders not found in AllDomainUsers`,
          unmatched.slice(0, 20).map(item => item.name),
        );
      }
      this.cdr.markForCheck();
      this.cdr.detectChanges();
    });
  }

  protected get hasLoaded(): boolean {
    return this.items.length > 0;
  }

  protected get filteredItems(): HrFileListItem[] {
    return filterAndSortHrFileItems(this.enrichedItems, this.searchQuery, this.sortDirection);
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

  /** Start Comments warm while the user aims / presses — before click navigation. */
  protected onItemPrefetch(item: HrFileListItem): void {
    this.itemPrefetchRequested.emit(item);
  }

  protected isItemSelected(item: HrFileListItem): boolean {
    return this.selectedItemId === item.id;
  }

  private resetPagination(): void {
    this.visibleCount = this.pageSize;
  }
}
