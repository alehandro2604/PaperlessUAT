/** Colour of a status pill in the task details panel. */
export type HrTaskDetailTone = 'success' | 'warning' | 'danger' | 'neutral';

/** Sectioned view of an HR task, rendered by the details panel. */
export interface HrTaskDetailView {
  title: string;
  subtitle: string;
  status: { value: string; tone: HrTaskDetailTone } | null;
  statuses: { label: string; value: string; tone: HrTaskDetailTone }[];
  references: { label: string; value: string }[];
  people: { role: string; name: string; initials: string; isMe: boolean }[];
  details: { label: string; value: string }[];
  approvals: { label: string; name: string; date: string; comment: string }[];
  timeline: { label: string; value: string }[];
}
