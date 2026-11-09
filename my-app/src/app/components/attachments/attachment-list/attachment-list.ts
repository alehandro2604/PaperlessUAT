import { DatePipe } from '@angular/common';
import { Component, input, output } from '@angular/core';

//what one row in the list needs.Exported so other files can use the same type
export interface AttachmentFile {
  id: string;
  name: string;
  webUrl: string;
  isFolder: boolean;
  lastModifiedDateTime?: string;
  modifiedBy?: string;
  fileIcon?: string;
  driveId?: string;
}

@Component({
  selector: 'app-attachment-list',
  imports: [DatePipe],
  templateUrl: './attachment-list.html',
  styleUrls: ['../attachments.css', './attachment-list.css'],
})

export class AttachmentList {
  files = input.required<AttachmentFile[]>();//required must pass it
  highlightedId = input<string | null>(null); //row to highlight after a search click
  fallbackAuthor = input(''); // shown when a file has no author
  selectedId = input<string | null>(null); // file currently previewed -> active row

  folderOpened = output<AttachmentFile>();
  fileSelected = output<AttachmentFile>();

  /** Click / tap a file -> preview it. */
  protected onClick(item: AttachmentFile) {
    if (item.isFolder) return;
    this.fileSelected.emit(item);
  }
}