import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { BehaviorSubject, Observable, from, of } from 'rxjs';
import { catchError, map, shareReplay, switchMap, tap } from 'rxjs/operators';
import { EForm } from '../models/form-configuration.model';
import { AuthService } from './auth.service';
import { sharePointConfig } from '../sharepoint.config';
import { graphGetWithRetry } from '../microsoft-graph';
import { AppConstants } from '../app.constants';

const SITE_ID = sharePointConfig.siteId;
const EFORMS_LIST = 'eForms';



@Injectable({ providedIn: 'root' })
// This service is used to load the form configurations from the eForms list in SharePoint
export class FormConfigurationService {

    //Behaviorsubjects is used to hold the forms, loading and error states
  private _forms$   = new BehaviorSubject<EForm[]>([]);
  private _loading$ = new BehaviorSubject<boolean>(false);
  private _error$   = new BehaviorSubject<string | null>(null);

  //observables is used to emit the forms, loading and error states
  readonly forms$:   Observable<EForm[]>       = this._forms$.asObservable();
  readonly loading$: Observable<boolean>        = this._loading$.asObservable();
  readonly error$:   Observable<string | null>  = this._error$.asObservable();

  readonly hrForms$: Observable<EForm[]> = this.forms$.pipe(
    map(forms => forms.filter(f => f.isHRWorkflow))
  );

  readonly workflowForms$: Observable<EForm[]> = this.forms$.pipe(
    map(forms => forms.filter(f => f.isSAPWorkflow))
  );

  readonly hrTaskListNames$: Observable<string[]> = this.hrForms$.pipe(
    map(forms => forms.map(f => f.taskListName).filter(Boolean))
  );

  readonly documentLibraries$: Observable<{ name: string; displayName: string }[]> =
  this.forms$.pipe(
    map(forms =>
      forms
        .filter(f => !!f.docLibraryName && f.AllFilesCheckBox === true)
        .map(f => ({ name: f.docLibraryName, displayName: f.title, order: f.AllFilesCheckBox }))
        .filter((lib, i, self) => i === self.findIndex(l => l.name === lib.name))
        .sort((a, b) => (a.order ? 1 : -1) - (b.order ? 1 : -1))
        .map(({ name, displayName }) => ({ name, displayName }))
      )
    );

  readonly documentLibraryTaskMap$: Observable<Record<string, string>> =
    this.forms$.pipe(
      map(forms =>
        forms.reduce((acc, f) => {
          if (f.docLibraryName && f.taskListName) {
            acc[f.docLibraryName] = f.taskListName;
          }
          return acc;
        }, {} as Record<string, string>)
      )
    );

  constructor(private http: HttpClient, private authService: AuthService) {}

  /**
   * Shared across concurrent callers. Several components ask for the eForms config while
   * the first request is still in flight; without this each of them fired its own copy
   * (the same config was fetched four times per page load).
   */
  private inFlightLoad$: Observable<void> | null = null;

  load(): Observable<void> {
    if (this.inFlightLoad$) return this.inFlightLoad$;

    this.inFlightLoad$ = this.loadFromSharePoint().pipe(
      tap({
        next: () => { this.inFlightLoad$ = null; },
        error: () => { this.inFlightLoad$ = null; },
      }),
      shareReplay({ bufferSize: 1, refCount: false }),
    );

    return this.inFlightLoad$;
  }

  private loadFromSharePoint(): Observable<void> {
    this._loading$.next(true);
    this._error$.next(null);

    // uses acquireGraphToken() — same token your HR tasks already use successfully.
    // graphGetWithRetry shares the app-wide Graph slot and backs off on 429.
    return from(this.authService.acquireGraphToken()).pipe(
      switchMap(token => {
        const path =
          `/sites/${SITE_ID}/lists/${EFORMS_LIST}/items` +
          `?$expand=fields($select=Title,Prefix,URL,TaskListName,TaskListGuid,DocLibraryName,HRWorkflow,SAPWorkflow,HROrder,SAPOrder,AllFiles,AllFilesCheckBox)` +
          `&$top=500`;

        return from(
          graphGetWithRetry(
            this.http,
            path,
            token,
            AppConstants.graphFileListingTimeoutMs,
          ) as Promise<{ value: any[] }>,
        );
      }),
      map(res => (res?.value ?? []).map(item => this.mapItem(item))
        .filter(form => form.title !== 'Extension of 3-Phase')
        .sort((a, b) => {
        // Sort HR workflows by HROrder, then alphabetically
        if (a.isHRWorkflow && b.isHRWorkflow) {
          if (a.HROrder !== b.HROrder) return a.HROrder - b.HROrder;
          return a.title.localeCompare(b.title);
        }
        // Sort SAP workflows by SAPOrder, then alphabetically
        if (a.isSAPWorkflow && b.isSAPWorkflow) {
          if (a.SAPOrder !== b.SAPOrder) return a.SAPOrder - b.SAPOrder;
          return a.title.localeCompare(b.title);
        }
        // Sort other forms alphabetically
        return a.title.localeCompare(b.title);
      })),
      tap(forms => {
        this._forms$.next(forms);
        this._loading$.next(false);
        console.log(`[FormConfig] Loaded ${forms.length} eForms`);
      }),
      map(() => void 0),
      catchError(err => {
        this._error$.next('Could not load eForms from SharePoint.');
        this._loading$.next(false);
        console.error('[FormConfig] error:', err);
        return of(void 0);
      })
    );
  }

  getFormByTitle(title: string): EForm | undefined {
    return this._forms$.value.find(
      f => f.title.toLowerCase() === title.toLowerCase()
    );
  }

  getFormByDocLibrary(libraryName: string): EForm | undefined {
    const target = String(libraryName ?? '').trim().toLowerCase();
    if (!target) return undefined;
    return this._forms$.value.find(
      (form) => String(form.docLibraryName ?? '').trim().toLowerCase() === target
    );
  }

  /**
   * True when the selected task maps to an eForm row in the SAP workflow list
   * (SharePoint eForms where SAPWorkflow is enabled and SAPOrder is set).
   */
  isSapOrderWorkflow(
    eFormListId?: string,
    eFormTitle?: string,
    taskListName?: string,
  ): boolean {
    const form = this.findMatchingForm(eFormListId, eFormTitle, taskListName);
    return !!form?.isSAPWorkflow && (form.SAPOrder ?? 0) > 0;
  }

  /** Request for Action is only offered for SAP workflow rows (SAPOrder list). */
  allowsRequestForAction(
    eFormListId?: string,
    eFormTitle?: string,
    taskListName?: string,
  ): boolean {
    return this.isSapOrderWorkflow(eFormListId, eFormTitle, taskListName);
  }

  private findMatchingForm(
    eFormListId?: string,
    eFormTitle?: string,
    taskListName?: string,
  ): EForm | undefined {
    const listId = String(eFormListId ?? '').trim().toLowerCase();
    const title = String(eFormTitle ?? '').trim().toLowerCase();
    const listName = String(taskListName ?? '').trim().toLowerCase();

    if (!listId && !title && !listName) {
      return undefined;
    }

    return this._forms$.value.find((form) => {
      const guid = String(form.taskListGuid ?? '').trim().toLowerCase();
      const formTitle = String(form.title ?? '').trim().toLowerCase();
      const formList = String(form.taskListName ?? '').trim().toLowerCase();

      if (listId && guid && guid === listId) return true;
      if (title && formTitle && formTitle === title) return true;
      if (listName && formList && formList === listName) return true;
      return false;
    });
  }

  get snapshot(): EForm[] {
    return this._forms$.value;
  }

  /** Every SharePoint task list referenced by the eForms config (HR + SAP workflows). */
  getAllTaskListNames(): string[] {
    return [...new Set(
      this._forms$.value
        .map((form) => String(form.taskListName ?? '').trim())
        .filter(Boolean),
    )];
  }

  /** Task lists for HR workflows only — used when opening a person in HR Files. */
  getHrTaskListNames(): string[] {
    return [...new Set(
      this._forms$.value
        .filter((form) => form.isHRWorkflow)
        .map((form) => String(form.taskListName ?? '').trim())
        .filter(Boolean),
    )];
  }

  /**
   * Source eForm / task list name from URL (`/Lists/{name}`) or Prefix — same rule as ChildSubject.
   * Example: SickCertificateUploader_Tasks.
   */
  resolveSourceListName(form: EForm): string {
    const fromUrl = this.parseSharePointListNameFromUrl(form.url);
    if (fromUrl) return fromUrl;
    return String(form.prefix ?? '').trim();
  }

  /**
   * HR Files / Comments person history: TaskListName plus each form's source eForm list
   * so submissions (e.g. Sick Certificate Upload) appear alongside workflow tasks.
   */
  getHrQueryListNames(): string[] {
    const names = new Set<string>();
    for (const form of this._forms$.value) {
      if (!form.isHRWorkflow) continue;
      const taskList = String(form.taskListName ?? '').trim();
      if (taskList) names.add(taskList);
      const source = this.resolveSourceListName(form);
      if (source) names.add(source);
    }
    return [...names];
  }

  /**
   * To Do discovery: all workflow TaskListNames plus source eForm lists when
   * TaskListName is empty, or when the source list differs (covers Upload_eForm rows).
   */
  getAllQueryListNames(): string[] {
    const names = new Set<string>();
    for (const form of this._forms$.value) {
      const taskList = String(form.taskListName ?? '').trim();
      if (taskList) names.add(taskList);
      const source = this.resolveSourceListName(form);
      if (source) names.add(source);
    }
    return [...names];
  }

  parseSharePointListNameFromUrl(url: string): string {
    const match = String(url ?? '').match(/\/lists\/([^/?#]+)/i);
    return match?.[1] ? decodeURIComponent(match[1]) : '';
  }

  private mapItem(item: any): EForm {
    const f = item.fields ?? {};

    // URL field via Graph comes back as an object with Url property for hyperlink fields
    const url = f.URL?.Url ?? f.URL ?? '';

    return {
      id:             item.id,
      title:          f.Title          ?? '',
      prefix:         f.Prefix         ?? '',
      url,
      taskListName:   f.TaskListName   ?? '',
      taskListGuid:   f.TaskListGuid   ?? '',
      docLibraryName: f.DocLibraryName ?? '',
      isHRWorkflow:   f.HRWorkflow     === true,
      isSAPWorkflow:  f.SAPWorkflow    === true,
      HROrder:        f.HROrder        ?? 0,
      SAPOrder:       f.SAPOrder       ?? 0,
      AllFiles:       f.AllFiles       ?? 0,
      AllFilesCheckBox: this.isSharePointYes(f.AllFilesCheckBox),
    };
  }

  /** Normalises SharePoint Yes/No / boolean checkbox values. */
  private isSharePointYes(value: unknown): boolean {
    if (value === true || value === 1) return true;
    const text = String(value ?? '').trim().toLowerCase();
    return text === 'yes' || text === 'true' || text === '1';
  }
}