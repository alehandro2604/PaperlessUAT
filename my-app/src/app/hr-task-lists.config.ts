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
