// Pure helpers for the HR task details view, extracted from AppComponent.
// They only turn task data into display rows - no component state, no services.

/** Flat label/value rows for an HR task, with section breaks between groups. */
export function buildHrTaskDetailRows(
  d: any,
  task: any,
): { label: string; value: string; sectionBreak?: boolean }[] {
  const clean = (v: any) =>
    !v || String(v).trim() === '' || String(v).trim() === 'Enter value here'
      ? '' : String(v).trim();

  const rows: { label: string; value: string; sectionBreak?: boolean }[] = [];
  const push = (label: string, value: string) => {
    if (value) rows.push({ label, value });
  };
  const endSection = () => {
    if (rows.length) rows[rows.length - 1].sectionBreak = true;
  };

  const fmtDate = (val: string) => {
    const dt = new Date(val);
    return isNaN(dt.getTime())
      ? val
      : dt.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
      ' ' + dt.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
  };

  const statusLower = clean(d.status).toLowerCase();
  const isApproved = statusLower.includes('approv') || statusLower.includes('complet');
  const isRejected = statusLower.includes('reject') || statusLower.includes('denied');
  const statusSuffix = isApproved ? ' (approved)' : isRejected ? ' (rejected)' : '';

  push('HR Task Type :', d.type || task.name);
  push('eForm List ID :', clean(d.eFormListId));
  push('Task ID :', clean(task.id));
  push('Status :', `${clean(d.status)}${statusSuffix}`);
  endSection();

  push('Requestor :', clean(d.submitter));
  if (clean(d.submittedDate)) push('submitted :', fmtDate(d.submittedDate));
  if (clean(d.section)) push('section :', clean(d.section));
  endSection();

  if (clean(d.missedDate)) push('date :', clean(d.missedDate));
  if (clean(d.timeIn)) push('Time In:', `${clean(d.timeIn)}:00`);
  if (clean(d.timeOut)) push('time out :', `${clean(d.timeOut)}:00`);
  if (clean(d.reason)) push('Reason:', clean(d.reason));
  if (clean(d.fromDate)) push('from date :', clean(d.fromDate));
  if (clean(d.toDate)) push('to date :', clean(d.toDate));
  if (rows.length && rows[rows.length - 1].sectionBreak !== true) endSection();

  if (clean(d.approver1)) {
    push('Approver 1:', clean(d.approver1));
    if (clean(d.approver1Date)) push('Approver 1 Date :', clean(d.approver1Date));
    if (clean(d.approver1Comment)) push('Approver 1 Comment :', clean(d.approver1Comment));
  }
  if (clean(d.approver2)) {
    push('Approver 2:', clean(d.approver2));
    if (clean(d.approver2Date)) push('Approver 2 Date :', clean(d.approver2Date));
    if (clean(d.approver2Comment)) push('Approver 2 Comment :', clean(d.approver2Comment));
  }
  if (clean(d.approver3)) {
    push('Approver 3:', clean(d.approver3));
    if (clean(d.approver3Date)) push('Approver 3 Date :', clean(d.approver3Date));
    if (clean(d.approver3Comment)) push('Approver 3 Comment :', clean(d.approver3Comment));
  }

  if (clean(d.hoursForTransfer)) push('Hours to Transfer :', clean(d.hoursForTransfer));
  if (clean(d.fromYear)) push('From Year :', clean(d.fromYear));
  if (clean(d.toYear)) push('To Year :', clean(d.toYear));
  if (clean(d.needEngineer)) push('Engineer Approval :', clean(d.needEngineer));
  if (clean(d.requestorPin)) push('Requestor PIN :', clean(d.requestorPin));
  if (clean(d.stage)) push('Current Stage :', clean(d.stage));
  if (clean(d.eFormCategory)) push('Category :', clean(d.eFormCategory));
  if (clean(d.eFormProgress)) push('Progress :', clean(d.eFormProgress));
  if (clean(d.taskOutcomeField)) push('Task Outcome :', clean(d.taskOutcomeField));
  if (clean(d.assignedTo2)) push('Assigned To :', clean(d.assignedTo2));
  if (clean(d.assignedToSuperior)) push('Assigned To Superior :', clean(d.assignedToSuperior));
  if (clean(d.assignedToStepNo)) push('Assigned To Step No :', clean(d.assignedToStepNo));
  if (clean(d.assignedToSuperior1)) push('Superior 1 :', clean(d.assignedToSuperior1));
  if (clean(d.assignedToSuperior2field)) push('Superior 2 :', clean(d.assignedToSuperior2field));
  if (clean(d.createdDate)) push('Created :', fmtDate(d.createdDate));
  if (clean(d.modifiedDate)) push('Modified :', fmtDate(d.modifiedDate));

  return rows;
}
