import { CommonModule } from '@angular/common';
import { Component, ElementRef, HostListener, Input, OnDestroy, OnInit, forwardRef } from '@angular/core';
import { ControlValueAccessor, NG_VALUE_ACCESSOR } from '@angular/forms';
import { Subject, Subscription } from 'rxjs';

export type AppDropdownTone =
  | 'neutral'
  | 'complete'
  | 'pending'
  | 'rejected'
  | 'comment'
  | 'attachment'
  | 'action';

export interface AppDropdownOption {
  value: string;
  label: string;
  disabled?: boolean;
  /** Optional status accent (colored dot) for filters like Select status. */
  tone?: AppDropdownTone;
}

@Component({
  selector: 'app-dropdown',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './app-dropdown.component.html',
  styleUrls: ['./app-dropdown.component.css'],
  host: {
    class: 'app-dropdown-host',
    '[class.app-dropdown-host--full]': 'fullWidth',
    '[class.app-dropdown-host--icon]': 'variant === "icon"',
    '[class.app-dropdown-host--placeholder]': 'isPlaceholder',
    '[class.open]': 'open',
  },
  providers: [
    {
      provide: NG_VALUE_ACCESSOR,
      useExisting: forwardRef(() => AppDropdownComponent),
      multi: true,
    },
  ],
})
export class AppDropdownComponent implements ControlValueAccessor, OnInit, OnDestroy {
  /** When any dropdown opens, others subscribe and close. */
  private static readonly opened$ = new Subject<AppDropdownComponent>();

  @Input() options: AppDropdownOption[] = [];
  @Input() placeholder = 'Select...';
  @Input() ariaLabel = 'Select option';
  @Input() triggerClass = '';
  @Input() fullWidth = false;
  @Input() menuAlign: 'left' | 'right' = 'left';
  @Input() variant: 'text' | 'icon' = 'text';
  @Input() iconSrc = 'filter.png';
  @Input() iconAlt = '';
  @Input() triggerId = '';

  open = false;
  value = '';
  disabled = false;
  /** Fixed-position coordinates so the menu is not clipped by overflow:hidden ancestors. */
  menuStyle: Record<string, string> = {};

  private onChange: (value: string) => void = () => {};
  private onTouched: () => void = () => {};
  private openedSub?: Subscription;
  private readonly closeOnScroll = (event: Event): void => {
    const target = event.target as Node | null;
    if (target && this.host.nativeElement.querySelector('.app-dropdown-menu')?.contains(target)) return;
    this.setOpen(false);
  };

  constructor(private readonly host: ElementRef<HTMLElement>) {}

  get selectedOption(): AppDropdownOption | undefined {
    return this.options.find(option => option.value === this.value);
  }

  get displayLabel(): string {
    return this.selectedOption?.label ?? this.placeholder;
  }

  get selectedTone(): AppDropdownTone | undefined {
    return this.selectedOption?.tone;
  }

  get isPlaceholder(): boolean {
    return !this.selectedOption;
  }

  ngOnInit(): void {
    this.openedSub = AppDropdownComponent.opened$.subscribe(source => {
      if (source !== this) this.setOpen(false);
    });
  }

  ngOnDestroy(): void {
    this.openedSub?.unsubscribe();
    window.removeEventListener('scroll', this.closeOnScroll, true);
  }

  /** Single place that opens/closes so the scroll listener only exists while the menu is open. */
  private setOpen(value: boolean): void {
    if (this.open === value) return;
    this.open = value;
    if (value) {
      this.positionMenu();
      window.addEventListener('scroll', this.closeOnScroll, true);
    } else {
      window.removeEventListener('scroll', this.closeOnScroll, true);
    }
  }

  private positionMenu(): void {
    const rect = this.host.nativeElement.getBoundingClientRect();
    const gap = 6;
    const margin = 8;
    const width = Math.max(rect.width, 220);
    // Content-sized variant (see CSS): only set a minimum width and reserve its max width when clamping.
    const contentSized = this.host.nativeElement.classList.contains('assigned-select-dropdown');
    const clampWidth = contentSized ? 320 : width;
    const spaceBelow = window.innerHeight - rect.bottom - gap - margin;
    const spaceAbove = rect.top - gap - margin;
    const openUp = spaceBelow < 180 && spaceAbove > spaceBelow;
    const maxHeight = Math.min(320, openUp ? spaceAbove : spaceBelow);
    const preferredLeft = this.menuAlign === 'right' ? rect.right - width : rect.left;
    const left = Math.max(margin, Math.min(preferredLeft, window.innerWidth - clampWidth - margin));
    this.menuStyle = {
      left: `${left}px`,
      right: 'auto',
      ...(contentSized ? { minWidth: `${width}px` } : { width: `${width}px` }),
      maxHeight: `${Math.max(maxHeight, 80)}px`,
      top: openUp ? 'auto' : `${rect.bottom + gap}px`,
      bottom: openUp ? `${window.innerHeight - rect.top + gap}px` : 'auto',
    };
  }

  @HostListener('window:resize')
  closeOnResize(): void {
    this.setOpen(false);
  }

  writeValue(value: string | null): void {
    this.value = value ?? '';
  }

  registerOnChange(fn: (value: string) => void): void {
    this.onChange = fn;
  }

  registerOnTouched(fn: () => void): void {
    this.onTouched = fn;
  }

  setDisabledState(isDisabled: boolean): void {
    this.disabled = isDisabled;
  }

  toggle(event: Event): void {
    event.stopPropagation();
    if (this.disabled) return;
    this.setOpen(!this.open);
    if (this.open) {
      AppDropdownComponent.opened$.next(this);
      this.onTouched();
    }
  }

  selectOption(option: AppDropdownOption, event: Event): void {
    event.stopPropagation();
    if (option.disabled || this.disabled) return;
    this.value = option.value;
    this.onChange(this.value);
    this.setOpen(false);
  }

  @HostListener('document:click')
  closeOnOutsideClick(): void {
    this.setOpen(false);
  }
}