export interface ToDoTask {
  listName: any;
  eFormDetails: any;
  id: string;
  name: string;
  description: string;
  status: 'pending' | 'in_progress' | 'completed' | 'rejected';
  assignedTo: string;
  submittedBy: string;
  createdDate: Date;
  updatedDate: Date;
}
