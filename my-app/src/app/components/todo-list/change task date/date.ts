import { CommonModule } from '@angular/common';
import {
  Component,
  ElementRef,
  EventEmitter,
  HostListener,
  Input,
  OnChanges,
  Output,
  SimpleChanges,
  ViewEncapsulation,
} from '@angular/core';
import { FormsModule } from '@angular/forms';
import { TodoService } from '../../../services/todo.service';

interface CalendarDay {
  day: number;
  iso: string;
  inMonth: boolean;
  isToday: boolean;
  isSelected: boolean;
}

@Component({
  selector: 'app-change-task-date',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './date.html',
  styleUrls: ['./date.css'],
  encapsulation: ViewEncapsulation.None,
})
export class ChangeTaskDateComponent implements OnChanges {
  @Input() task: any;
  @Output() dateChange = new EventEmitter<{ dueDate: string }>();

  isOpen = false;
  selectedDate = '';
  displayDate = '';
  isSaving = false;
  saveError = '';

  readonly weekdays = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
  viewYear = new Date().getFullYear();
  viewMonth = new Date().getMonth();
  calendarDays: CalendarDay[] = [];
  monthLabel = '';

  constructor(
    private readonly host: ElementRef<HTMLElement>,
    private readonly todoService: TodoService
  ) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['task']) {
      this.setSelectedDate(this.toInputDate(this.readDueDate(this.task)));
      this.syncViewToSelected();
      this.rebuildCalendar();
    }
  }

  onButtonClick(event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.isOpen = !this.isOpen;
    if (this.isOpen) {
      this.setSelectedDate(this.toInputDate(this.readDueDate(this.task)));
      this.saveError = '';
      this.syncViewToSelected();
      this.rebuildCalendar();
    }
  }

  onPanelClick(event: MouseEvent): void {
    event.stopPropagation();
  }

  shiftMonth(delta: number, event?: MouseEvent): void {
    event?.preventDefault();
    event?.stopPropagation();
    const next = new Date(this.viewYear, this.viewMonth + delta, 1);
    this.viewYear = next.getFullYear();
    this.viewMonth = next.getMonth();
    this.rebuildCalendar();
  }

  selectDay(iso: string, event?: MouseEvent): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.setSelectedDate(iso);
    this.syncViewToSelected();
    this.rebuildCalendar();
  }

  selectToday(event?: MouseEvent): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.setSelectedDate(this.formatLocalDate(new Date()));
    this.syncViewToSelected();
    this.rebuildCalendar();
  }

  clearDate(event?: MouseEvent): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.setSelectedDate('');
    this.rebuildCalendar();
  }

  onManualDateChange(value: string): void {
    this.displayDate = value;
    const parsed = this.parseDisplayDate(value);
    if (parsed === null && value.trim() !== '') return;
    this.selectedDate = parsed ?? '';
    if (this.selectedDate) {
      this.syncViewToSelected();
      this.rebuildCalendar();
    } else {
      this.rebuildCalendar();
    }
  }

  onCancel(event?: MouseEvent): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.isOpen = false;
    this.saveError = '';
    this.setSelectedDate(this.toInputDate(this.readDueDate(this.task)));
  }

  async onApply(event: MouseEvent): Promise<void> {
    event.preventDefault();
    event.stopPropagation();
    if (!this.selectedDate || this.isSaving) return;

    this.isSaving = true;
    this.saveError = '';

    try {
      await this.saveToSharepoint(this.task, this.selectedDate);
      this.dateChange.emit({ dueDate: this.selectedDate });
      this.isOpen = false;
    } catch (err) {
      console.error('Failed to save due date to SharePoint', err);
      const detail = err instanceof Error && err.message ? err.message : '';
      this.saveError = detail
        ? `Could not save the due date. ${detail}`
        : 'Could not save the due date. Please try again.';
    } finally {
      this.isSaving = false;
    }
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    if (!this.isOpen) return;
    const target = event.target as Node | null;
    if (target && this.host.nativeElement.contains(target)) return;
    this.onCancel();
  }

  @HostListener('document:keydown', ['$event'])
  onDocumentKeydown(event: KeyboardEvent): void {
    if (!this.isOpen) return;
    if (event.key === 'Escape') this.onCancel();
  }

  private syncViewToSelected(): void {
    const base = this.selectedDate
      ? new Date(`${this.selectedDate}T00:00:00`)
      : new Date();
    if (Number.isNaN(base.getTime())) return;
    this.viewYear = base.getFullYear();
    this.viewMonth = base.getMonth();
  }

  private rebuildCalendar(): void {
    const monthNames = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December',
    ];
    this.monthLabel = `${monthNames[this.viewMonth]} ${this.viewYear}`;

    const first = new Date(this.viewYear, this.viewMonth, 1);
    const startOffset = first.getDay();
    const gridStart = new Date(this.viewYear, this.viewMonth, 1 - startOffset);
    const todayIso = this.formatLocalDate(new Date());

    const days: CalendarDay[] = [];
    for (let i = 0; i < 42; i++) {
      const date = new Date(
        gridStart.getFullYear(),
        gridStart.getMonth(),
        gridStart.getDate() + i
      );
      const iso = this.formatLocalDate(date);
      days.push({
        day: date.getDate(),
        iso,
        inMonth: date.getMonth() === this.viewMonth,
        isToday: iso === todayIso,
        isSelected: !!this.selectedDate && iso === this.selectedDate,
      });
    }
    this.calendarDays = days;
  }

  private setSelectedDate(iso: string): void {
    this.selectedDate = iso;
    this.displayDate = iso ? this.toDisplayDate(iso) : '';
  }

  private toDisplayDate(iso: string): string {
    const match = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return iso;
    return `${match[3]}/${match[2]}/${match[1]}`;
  }

  /** Accepts `DD/MM/YYYY` or `YYYY-MM-DD`. Returns ISO or null if incomplete/invalid. */
  private parseDisplayDate(value: string): string | null {
    const trimmed = value.trim();
    if (!trimmed) return '';
    const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoMatch) {
      return this.isValidYmd(+isoMatch[1], +isoMatch[2], +isoMatch[3])
        ? trimmed
        : null;
    }
    const dmyMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!dmyMatch) return null;
    const day = +dmyMatch[1];
    const month = +dmyMatch[2];
    const year = +dmyMatch[3];
    if (!this.isValidYmd(year, month, day)) return null;
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  private isValidYmd(year: number, month: number, day: number): boolean {
    if (month < 1 || month > 12 || day < 1 || day > 31) return false;
    const dt = new Date(year, month - 1, day);
    return (
      dt.getFullYear() === year &&
      dt.getMonth() === month - 1 &&
      dt.getDate() === day
    );
  }

  private formatLocalDate(date: Date): string {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  private readDueDate(task: any): string {
    const details = task?.eFormDetails ?? {};
    return String(
      details?.dueDate ??
        task?.dueDate ??
        details?.DueDate ??
        details?.rawFields?.DueDate ??
        ''
    ).trim();
  }

  /** Normalize any date-ish value to `YYYY-MM-DD`. */
  private toInputDate(value: string): string {
    if (!value) return '';
    const isoMatch = value.match(/^(\d{4}-\d{2}-\d{2})/);
    if (isoMatch) return isoMatch[1];
    const dt = new Date(value);
    if (Number.isNaN(dt.getTime())) return '';
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  /**
   * Writes the new due date to SharePoint via Graph, then syncs the local task
   * object so the UI reflects it immediately.
   */
  private async saveToSharepoint(task: any, dueDate: string): Promise<void> {
    await this.todoService.updateTaskDueDate(task, dueDate);

    const details = task?.eFormDetails ?? {};
    details.dueDate = dueDate;
    if (details.rawFields && typeof details.rawFields === 'object') {
      for (const key of ['DueDate', 'DueDate0', 'Due', 'Due_x0020_Date'] as const) {
        if (Object.prototype.hasOwnProperty.call(details.rawFields, key)) {
          details.rawFields[key] = dueDate;
          break;
        }
      }
    }
    task.eFormDetails = details;
    task.dueDate = dueDate;
  }
}
