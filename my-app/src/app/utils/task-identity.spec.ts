import { getTaskListName, isSameTask } from './task-identity';

describe('task identity', () => {
  const shiftDev51 = { id: '51', listName: 'SAPTaskShiftingAndDeviation' };
  const noFlow51 = { id: '51', listName: 'NoFlowTasks' };

  it('reads the list name from the task or its eForm details', () => {
    expect(getTaskListName(shiftDev51)).toBe('saptaskshiftinganddeviation');
    expect(getTaskListName({ eFormDetails: { listName: 'NoFlowTasks' } })).toBe('noflowtasks');
    expect(getTaskListName(null)).toBe('');
  });

  it('matches the same id in the same list', () => {
    expect(isSameTask(shiftDev51, '51', 'SAPTaskShiftingAndDeviation')).toBe(true);
    expect(isSameTask(shiftDev51, '51', 'saptaskshiftinganddeviation')).toBe(true);
  });

  it('does not match the same id in another list', () => {
    expect(isSameTask(noFlow51, '51', 'SAPTaskShiftingAndDeviation')).toBe(false);
  });

  it('only removes the completed task when two lists share an id', () => {
    const remaining = [shiftDev51, noFlow51].filter(t => !isSameTask(t, '51', 'SAPTaskShiftingAndDeviation'));
    expect(remaining).toEqual([noFlow51]);
  });

  it('falls back to id only when a list name is missing', () => {
    expect(isSameTask(noFlow51, '51')).toBe(true);
    expect(isSameTask({ id: '51' }, '51', 'NoFlowTasks')).toBe(true);
  });

  it('never matches a different or empty id', () => {
    expect(isSameTask(shiftDev51, '52', 'SAPTaskShiftingAndDeviation')).toBe(false);
    expect(isSameTask(shiftDev51, '', 'SAPTaskShiftingAndDeviation')).toBe(false);
  });
});
