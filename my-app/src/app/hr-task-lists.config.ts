// ============================================================
// KNOWN HR TASK LIST NAMES
// SharePoint list names that cannot be auto-detected from the
// site (auto-detection only picks up lists prefixed "HRTask").
// Usage:  import { HR_TASK_LIST_NAMES } from './hr-task-lists.config';
// ============================================================

export const HR_TASK_LIST_NAMES: readonly string[] = [
  'HRTaskChangeOfShiftIESC', 'HRTaskTeleworkReports', 'HRTaskTelework',
  'HRTaskRest', 'HRTaskSickLeaveByAppointment', 'HRTaskProbation', 'HRTaskMissingPunch',
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
  return /_eform$/i.test(key) || /eform$/i.test(key.replace(/[^a-z0-9]/g, ''));//this os checking if the list name ends with _eform or eform
}
