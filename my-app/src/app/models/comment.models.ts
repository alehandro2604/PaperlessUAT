
export interface NewCommentSavedEvent {
    taskId: number | string;
    category: string;
    comment: string;
    fileName?: string;
    attachmentUrl?: string;
    sharePointItemId?: string;
    assignedToName?: string;
    assignedToEmail?: string;
  }
  
  export interface CommentListItem {
    id: string;
    name: string;
    webUrl: string;
    lastModifiedDateTime?: string;
    size?: number;
    isFolder: boolean;
    status?: string;
    submittedBy?: string;
    submittedDate?: string;
    completedBy?: string;
    description?: string;
    contentType?: string;
    downloadUrl?: string;
    extractedContent?: string;
    eFormDetails?: any;
    isContentLoaded?: boolean;
    showFullContent?: boolean;
  }