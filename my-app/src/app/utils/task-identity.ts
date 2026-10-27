/**
 * SharePoint item ids are only unique within one list: ShiftDev item 51 and
 * NoFlowTasks item 51 are different tasks. Use these helpers whenever a task
 * is found again by id (complete, delegate, due date, cache updates).
 */

export function getTaskListName(item: { listName?: unknown; eFormDetails?: unknown } | null | undefined): string {
  const details = (item?.eFormDetails ?? {}) as Record<string, unknown>;
  return String(item?.listName ?? details['listName'] ?? '').trim().toLowerCase();
}

/**
 * True when `item` is the task with this id in this list. When either side has
 * no list name, falls back to matching on id alone (the old behaviour).
 */
export function isSameTask(item: any, taskId: string, listName?: string | null): boolean {
  const id = String(taskId ?? '').trim();
  if (!id || String(item?.id ?? item?.Id ?? '').trim() !== id) return false;

  const wantedList = String(listName ?? '').trim().toLowerCase();
  if (!wantedList) return true;
  const itemList = getTaskListName(item);
  return !itemList || itemList === wantedList;
}
