import { CommonModule } from '@angular/common';
import { Component, HostListener, Input, OnDestroy, OnInit, forwardRef } from '@angular/core';
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

  private onChange: (value: string) => void = () => {};
  private onTouched: () => void = () => {};
  private openedSub?: Subscription;

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
      if (source !== this) this.open = false;
    });
  }

  ngOnDestroy(): void {
    this.openedSub?.unsubscribe();
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
    this.open = !this.open;
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
    this.open = false;
  }

  @HostListener('document:click')
  closeOnOutsideClick(): void {
    this.open = false;
  }
}