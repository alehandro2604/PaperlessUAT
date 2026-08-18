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
  isSaving = false;
  saveError = '';

  constructor(
    private readonly host: ElementRef<HTMLElement>,
    private readonly todoService: TodoService
  ) {}

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['task']) {
      this.selectedDate = this.toInputDate(this.readDueDate(this.task));
    }
  }

  onButtonClick(event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();
    this.isOpen = !this.isOpen;
    if (this.isOpen) {
      this.selectedDate = this.toInputDate(this.readDueDate(this.task));
      this.saveError = '';
    }
  }

  onPanelClick(event: MouseEvent): void {
    event.stopPropagation();
  }

  onCancel(event?: MouseEvent): void {
    event?.preventDefault();
    event?.stopPropagation();
    this.isOpen = false;
    this.saveError = '';
    this.selectedDate = this.toInputDate(this.readDueDate(this.task));
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

  /** Normalize any date-ish value to `YYYY-MM-DD` for `<input type="date">`. */
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