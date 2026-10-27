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

  folderOpened = output<AttachmentFile>();
  fileSelected = output<AttachmentFile>();
  private hovereTime?: ReturnType<typeof setTimeout>;


  /** Mouse rests on a file for 500 ms -> ask the parent to preview it. */
  protected onHover(item: AttachmentFile) {
    if (item.isFolder) return;              // folders have no preview
    clearTimeout(this.hovereTime);          // cancel the previous row's timer
    this.hovereTime = setTimeout(() => {
      this.fileSelected.emit(item);         // "preview this file"
    }, 500);
  }

  protected onHoverLeave() {
    if (this.hovereTime) {
      clearTimeout(this.hovereTime);
      this.hovereTime = undefined;
    }
  }

  /** Click / tap (phones have no hover) -> preview straight away. */
  protected onClick(item: AttachmentFile) {
    if (item.isFolder) return;
    clearTimeout(this.hovereTime);
    this.fileSelected.emit(item);
  }
}