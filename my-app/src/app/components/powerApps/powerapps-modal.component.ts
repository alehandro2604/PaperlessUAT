import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Observable } from 'rxjs';
import { EForm } from '../../models/form-configuration.model';
import { FormConfigurationService } from '../../services/form-configuration.service';

@Component({
  selector: 'app-powerapps-modal',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './powerapps-modal.component.html',
  styleUrls: ['./powerapps-modal.component.css'],
})
export class PowerappsModalComponent {
  @Input() open = false;
  @Output() closed = new EventEmitter<void>();

  protected selectedFormTab: 'HR' | 'WorkFlows' = 'HR';

  //add power apps links here from the service
  //order by LiveOrder

  protected hrForms$: Observable<EForm[]> = new Observable<EForm[]>();
  protected workflowForms$: Observable<EForm[]> = new Observable<EForm[]>();
  protected loading$: Observable<boolean> = new Observable<boolean>();
  protected error$: Observable<string | null> = new Observable<string | null>();


  constructor(private formConfig: FormConfigurationService) {

    this.hrForms$ = this.formConfig.hrForms$;
    this.workflowForms$ = this.formConfig.workflowForms$;
    this.loading$ = this.formConfig.loading$;
    this.error$ = this.formConfig.error$;
  }

  ngOnInit(): void { }

  protected get activeForms$(): Observable<EForm[]> {
    return this.selectedFormTab === 'WorkFlows'
      ? this.workflowForms$
      : this.hrForms$;
  }

  protected setFormTab(tab: 'HR' | 'WorkFlows'): void {
    this.selectedFormTab = tab;
  }

  protected openForm(form: EForm): void {
    if (form.url) {
      window.open(form.url, '_blank');
    } else {
      console.error('No URL for form:', form.title);
    }
  }


  close(): void {
    this.closed.emit();
  }
}