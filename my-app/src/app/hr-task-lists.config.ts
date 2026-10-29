// ============================================================
// KNOWN HR TASK LIST NAMES
// SharePoint list names that cannot be auto-detected from the
// site (auto-detection only picks up lists prefixed "HRTask").
// Usage:  import { HR_TASK_LIST_NAMES } from './hr-task-lists.config';
// ============================================================

export const HR_TASK_LIST_NAMES: readonly string[] = [
  'HRTaskChangeOfShiftIESC', 'HRTaskTeleworkReports', 'HRTaskTelework',
  'HRTaskRest', 'HRTaskSickLeaveByAppointment', 'HRTaskProbation', 'HRTaskMissingPunch',
  'PerformanceReview_Tasks',
];

/**
 * Extra HR lists to always query for person history / To Do discovery.
 * Sick Certificate tasks live on SickCertificateUploader_Tasks (not HRTask*).
 */
export const HR_SOURCE_EFORM_LIST_NAMES: readonly string[] = [
  'SickCertificateUploader_Tasks',
];

/** Procurement / SAP lists — queried for To Do only, not HR Files person folders. */
export const TODO_EXTRA_TASK_LIST_NAMES: readonly string[] = [
  'ProcTasks', 'ProcTasks1', 'ECTasks',
];

/** True for procurement/SAP task lists included in the To Do scope. */
//this is used to determine if a list is a procurement task list
export function isTodoProcurementTaskList(listName: string): boolean {
  const key = String(listName ?? '').trim().toLowerCase();
  if (key === 'proctasksarchive') return true;
  return TODO_EXTRA_TASK_LIST_NAMES.some(name => name.toLowerCase() === key);
}

/** True for Sick Certificate / source eForm submission lists (vs standard HRTask*). */
//this is used to determine if a list is a source eForm list
export function isHrSourceEFormList(listName: string): boolean {
  const key = String(listName ?? '').trim().toLowerCase();
  if (!key) return false;
  if (HR_SOURCE_EFORM_LIST_NAMES.some(name => name.toLowerCase() === key)) return true;
  if (key.includes('sickcertificate')) return true;
  return /_eform$/i.test(key) || /eform$/i.test(key.replace(/[^a-z0-9]/g, ''));//this is checking if the list name ends with _eform or eform
}

/** Increment / probation / performance review are created by HR; Title holds the employee email. */
export function isHrTitleMatchedTaskList(listName: string): boolean {
  const key = String(listName ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  return key.includes('increment') // the key means that the list name contains the word increment
    || key.includes('probation')
    || key.includes('performancereview')
    || key.includes('preformancereview');
}

/** Emails allowed to see the "+" (add comment) button on the All Files tab. */
export const ALL_FILES_PLUS_ALLOWED_EMAILS: readonly string[] = [
  // "name@enemalta.com.mt",
  "joan-anne.zammit@enemalta.com.mt",
];

//allow the following libraries for every users

export const ALLOWED_LIBRARIES_FOR_ALL_USERS: readonly string[] = [
  'CFADocumentLibrary',
  'IllicitConnectionsDocumentLibrary',
  'TDIDocumentLibrary',
];

/** Microsoft 365 group ids used by the All Files "+" rules (same groups as the Power Apps DisplayMode). */
export const ALL_FILES_GROUPS = {
  procurement: '85646fac-3f37-458e-9e36-71c39542599d',
  claimsForDamages: 'b34c0719-4562-4013-9ce0-fcef85928e9f',
  electricitySupply: '1352a0e4-c158-4898-b674-dab944b95b8b',
  pv: '6a5c8870-9faf-45cc-863b-3a0e9cf31516',
  registry: 'cc797395-a2d1-4c64-8ec5-bca630c754a0',
} as const;

/**
 * Per-library rules for the All Files "+". libraryKey is the lower-case, letters-and-digits form of the
 * library name. A rule allows the user when they are in groupId, or (allowOpener) opened the file.
 */
export interface AllFilesLibraryRule { libraryKey: string; groupId?: string; allowOpener?: boolean }
export const ALL_FILES_LIBRARY_RULES: readonly AllFilesLibraryRule[] = [
  { libraryKey: 'procurementdocumentlibrary', groupId: ALL_FILES_GROUPS.procurement, allowOpener: true }, // 2, 3
  { libraryKey: 'ecdocumentlibrary', groupId: ALL_FILES_GROUPS.claimsForDamages },                        // 6
  { libraryKey: 'electricitysupplydocumentlibrary', groupId: ALL_FILES_GROUPS.electricitySupply, allowOpener: true }, // 9, 10
  { libraryKey: 'pvdocumentlibrary', groupId: ALL_FILES_GROUPS.pv },                                      // 11
  { libraryKey: 'noflowdocumentlibrary', groupId: ALL_FILES_GROUPS.registry },                           // 12
];
