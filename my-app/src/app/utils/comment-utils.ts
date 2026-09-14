// Pure comment classification + HTML-building helpers extracted from AppComponent.
// None of these touch component state or injected services - verified via a full
// this-reference dependency trace before extraction.

import { normalizeSharePointFileUrl } from '../file-utils';
import { CommentSearchHit } from '../components/All-files/all-files';

const COMMENT_ENTRY_TITLES = new Set([
  'general comment',
  'new attachment',
  'request for action',
]);

const STATUSLESS_COMMENT_TITLES = new Set([
  'general comment',
  'new attachment',
  'request for action',
]);

export function isSyntheticCommentCard(item: { id?: unknown }): boolean {
  return String(item.id ?? '').startsWith('comment:');
}

export function matchesCommentLabel(value: string, labels: Set<string>): boolean {
  const normalized = value.trim().toLowerCase();
  if (!normalized) return false;
  if (labels.has(normalized)) return true;

  return [...labels].some(entry => {
    if (normalized === entry || normalized.startsWith(`${entry}:`)) return true;
    // "New Attachment/s", "New Attachments", "General Comments", etc.
    if (normalized.startsWith(entry)) return true;
    const withoutSlashS = normalized.replace(/\/s\b/g, 's');
    return (
      withoutSlashS === entry ||
      withoutSlashS === `${entry}s` ||
      withoutSlashS.startsWith(entry)
    );
  });
}

export function isCommentEntryTitle(title: string): boolean {
  return matchesCommentLabel(title, COMMENT_ENTRY_TITLES);
}

export function isStatuslessCommentTitle(title: string): boolean {
  return matchesCommentLabel(title, STATUSLESS_COMMENT_TITLES);
}

export function getCommentEntryKind(item: {
  id?: unknown;
  name?: unknown;
  eFormDetails?: Record<string, unknown>;
}): 'general' | 'attachment' | 'action' | null {
  const details = item.eFormDetails ?? {};
  const commentCategory = String(details['commentCategory'] ?? '').trim().toLowerCase();
  if (commentCategory === 'general') return 'general';
  if (commentCategory === 'attachment') return 'attachment';
  if (commentCategory === 'action') return 'action';

  const candidates = [
    details['category'],
    details['type'],
    item.name,
    (details['rawFields'] as Record<string, unknown> | undefined)?.['Title'],
    (details['rawFields'] as Record<string, unknown> | undefined)?.['Category'],
  ];

  for (const value of candidates) {
    const title = String(value ?? '');
    if (matchesCommentLabel(title, new Set(['general comment']))) return 'general';
    if (matchesCommentLabel(title, new Set(['new attachment']))) return 'attachment';
    if (matchesCommentLabel(title, new Set(['request for action']))) return 'action';
  }

  // Synthetic comment cards without a resolvable category still count as comment entries.
  if (isSyntheticCommentCard(item)) return 'general';
  return null;
}

export function shouldHideCommentStatus(item: {
  id?: unknown;
  name?: unknown;
  eFormDetails?: Record<string, unknown>;
}): boolean {
  return getCommentEntryKind(item) !== null;
}

export function isCommentTypeItem(item: {
  id?: unknown;
  name?: unknown;
  eFormDetails?: Record<string, unknown>;
}): boolean {
  return getCommentEntryKind(item) !== null;
}

export function isSharePointCommentEntry(item: any): boolean {
  const fields = item.fields ?? item;
  return isCommentEntryTitle(String(fields.Title ?? '').trim());
}

export function extractCommentCardTaskId(id: unknown): string {
  const raw = String(id ?? '').trim();
  if (!raw.startsWith('comment:')) return '';
  const parts = raw.split(':');
  return parts.length >= 2 ? parts[1].trim() : '';
}

export function matchesApprovalFilter(
  item: {
    id?: unknown;
    name?: unknown;
    status?: unknown;
    eFormDetails?: Record<string, unknown>;
  },
  filter: string
): boolean {
  const kind = getCommentEntryKind(item);

  switch (filter) {
    case 'approved': {
      if (kind) return false;
      const s = String(item.eFormDetails?.['status'] ?? item.status ?? '').toLowerCase();
      return s.includes('approv') || s.includes('complet');
    }
    case 'pending': {
      if (kind) return false;
      const s = String(item.eFormDetails?.['status'] ?? item.status ?? '').toLowerCase();
      const isApproved = s.includes('approv') || s.includes('complet');
      return !isApproved;
    }
    case 'general-comments':
      return kind === 'general';
    case 'new-attachments':
      return kind === 'attachment';
    case 'rfa':
      return kind === 'action';
    default:
      return true;
  }
}

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function buildCommentHtmlWithAttachment(
  comment: string,
  fileName: string,
  attachmentUrl: string
): string {
  const safeComment = escapeHtml(comment.trim()).replace(/\n/g, '<br>');

  if (!fileName || !attachmentUrl) {
    return safeComment;
  }

  const safeName = escapeHtml(fileName);
  const openUrl = normalizeSharePointFileUrl(attachmentUrl);
  const safeUrl = escapeHtml(openUrl);
  const attachmentLink =
    `<a class="comment-link" href="${safeUrl}" target="_blank" rel="noopener noreferrer">` +
    `${safeName}</a>`;

  return safeComment ? `${safeComment}<br><br>${attachmentLink}` : attachmentLink;
}

export function extractAttachmentFromCommentHtml(html: string): { url: string; fileName: string } {
  const match = html.match(/<a\b[^>]*\bhref=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i);// this is to extract the url and the file name from the comment html
  if (!match) {
    return { url: '', fileName: '' };// this is to return an empty object if the match is not found
  }
  return {
    url: match[1].replace(/&amp;/g, '&').trim(),// this is to replace the &amp; with &
    fileName: match[2].replace(/<[^>]+>/g, '').trim(),// this is to remove the html tags from the file name
  };
}

export function openAttachmentLink(event: Event, url: string): void {
  event.preventDefault();
  event.stopPropagation();
  // Strip stray quote characters left over from encoded hrefs (&quot;...&quot;).
  const cleaned = String(url ?? '').trim().replace(/^["']+|["']+$/g, '');
  const target = normalizeSharePointFileUrl(cleaned);
  if (!target) {
    return;
  }
  // Always open in a new tab - never navigate the app itself.
  window.open(target, '_blank', 'noopener,noreferrer');
}

export function onCommentLinkClick(event: Event): void {
  const anchor = (event.target as HTMLElement | null)?.closest('a');
  if (!anchor || !(anchor instanceof HTMLAnchorElement)) {
    return;
  }
  const href = anchor.getAttribute('href') ?? anchor.href;
  if (!href || href.startsWith('#')) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  openAttachmentLink(event, href);
}

export function buildQuickCommentItemFromSearchHit(hit: CommentSearchHit, listName: string): any {
  const body = String(hit.body ?? '').trim();
  const author = String(hit.author ?? '').trim();
  const date = String(hit.date ?? '').trim();
  return {
    id: hit.taskId,
    name: hit.title || 'Comment',
    webUrl: '',
    lastModifiedDateTime: date,
    isFolder: false,
    status: '',
    submittedBy: author,
    submittedDate: date,
    description: body,
    listName,
    isContentLoaded: true,
    eFormDetails: {
      type: hit.title || 'Comment',
      status: '',
      submitter: author,
      listName,
      submittedDate: date,
      body,
      comment: body,
      commentHtml: body.replace(/\n/g, '<br>'),
      customCreatedDate: date,
      customModifiedDate: date,
      customModifiedBy: author,
    },
  };
}
