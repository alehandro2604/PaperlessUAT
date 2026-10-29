import { DatePipe } from '@angular/common';
import { Component, HostListener, effect, inject, input, signal } from '@angular/core';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { AttachmentFile } from '../attachments/attachment-list/attachment-list';
import { AuthService } from '../../services/auth.service';
import { DriveFilesService } from '../../services/drive-files.service';
/** Shows a live SharePoint preview of one file. */
@Component({
  selector: 'app-preview-attachments',
  imports: [DatePipe],             // the template uses the "date" pipe
  templateUrl: './preview.html',
  styleUrl: './preview.css',
})
export class PreviewAttachments {
  /** The file to show. null = show nothing. Comes from app.html. */
  file = input<AttachmentFile | null>(null);

  /** What the box shows: waiting, the iframe, or a "no preview" message. */
  protected status = signal<'loading' | 'ready' | 'failed'>('loading');
  /** The preview link, in the "safe" wrapper Angular needs for an iframe src. */
  protected link = signal<SafeResourceUrl | null>(null);

  private auth = inject(AuthService);            // gets the login token
  private driveFiles = inject(DriveFilesService); // calls Graph
  private sanitizer = inject(DomSanitizer);      // lets us put the link in an iframe
  protected expand = signal(false);                       // whether the user clicked "expand" to see the full preview
  protected toggleExpand(): void { this.expand.update(v => !v); }

  // Collapse the expanded preview on Esc, or if the window is resized.
  @HostListener('document:keydown.escape')
  @HostListener('window:resize')
  protected collapse(): void { this.expand.set(false); }

  constructor() {
    // effect() re-runs whenever a signal it reads changes. Here: every time file() changes.
    // Like React's useEffect(() => { ... }, [file]).
    effect(() => {
      const f = this.file();
      if (f) this.load(f);
    });
  }

  private async load(f: AttachmentFile): Promise<void> {
    this.status.set('loading');
    this.link.set(null);
    this.expand.set(false);           // a new file always starts collapsed

    if (!f.driveId) {                 // e.g. an HR task list item, not a SharePoint file
      this.status.set('failed');
      return;
    }

    try {
      const token = await this.auth.acquireSharePointToken();
      const url = await this.driveFiles.getPreviewUrl(f.driveId, f.id, token);

      // While we waited, the user may have hovered another file -> this answer is old, ignore it.
      if (this.file() !== f) return;

      if (url) {
        // Angular blocks unknown URLs in iframes. The service checked it's SharePoint, so allow it.
        this.link.set(this.sanitizer.bypassSecurityTrustResourceUrl(url));
        this.status.set('ready');
      } else {
        this.status.set('failed');
      }
    } catch {
      // Graph refuses file types it can't preview (e.g. .zip), or the network failed.
      if (this.file() === f) this.status.set('failed');
    }
  }
}