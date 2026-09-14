// Pure builders that turn a saved new-comment event into the comment card the
// To Do / Comments lists render. None of these touch component state or injected
// services - the caller supplies everything varying through SavedCommentCardContext.

import { NewCommentSavedEvent, CommentListItem } from '../models/comment.models';
import { buildCommentHtmlWithAttachment } from './comment-utils';

export interface SavedCommentCardContext {
  /** Display name recorded as the author on the generated card. */
  commentAuthor: string;
  /** eFormListId used when the base card carries none. */
  fallbackEFormListId: string;
}

interface SavedCommentDetails {
  status: string;
  commentHtml: string;
  commentCategory: string;
  attachmentFileName: string;
  attachmentUrl?: string;
  assignedTo: string;
  category: string;
}

/** Dispatches to the card shape matching the saved comment's category. */
export function buildCommentCardFromSavedEvent(
  ev: NewCommentSavedEvent,
  base: CommentListItem,
  taskId: string,
  ctx: SavedCommentCardContext,
): CommentListItem {
  switch (ev.category) {
    case 'general':
      return buildGeneralCommentCard(ev, base, taskId, ctx);
    case 'attachment':
      return buildAttachmentCommentCard(ev, base, taskId, ctx);
    case 'action':
      return buildActionCommentCard(ev, base, taskId, ctx);
    default:
      return buildDefaultSavedCommentCard(ev, base, taskId, ctx);
  }
}

export function buildGeneralCommentCard(
  ev: NewCommentSavedEvent,
  base: CommentListItem,
  taskId: string,
  ctx: SavedCommentCardContext,
): CommentListItem {
  return createSavedCommentItem(base, taskId, 'General Comment', {
    commentHtml: ev.comment,
    commentCategory: ev.category,
    attachmentFileName: '',
    assignedTo: '',
    category: '',
    status: ''
  }, ctx, ev.sharePointItemId);
}

export function buildAttachmentCommentCard(
  ev: NewCommentSavedEvent,
  base: CommentListItem,
  taskId: string,
  ctx: SavedCommentCardContext,
): CommentListItem {
  const attachmentFileName = String(ev.fileName ?? '').trim();
  const attachmentUrl = String(ev.attachmentUrl ?? '').trim();
  const commentHtml = buildCommentHtmlWithAttachment(
    ev.comment,
    attachmentFileName,
    attachmentUrl
  );

  return createSavedCommentItem(base, taskId, 'New Attachment', {
    status: '',
    commentHtml,
    commentCategory: ev.category,
    attachmentFileName,
    attachmentUrl,
    assignedTo: '',
    category: '',
  }, ctx, ev.sharePointItemId);
}

export function buildActionCommentCard(
  ev: NewCommentSavedEvent,
  base: CommentListItem,
  taskId: string,
  ctx: SavedCommentCardContext,
): CommentListItem {
  const attachmentFileName = String(ev.fileName ?? '').trim();
  const attachmentUrl = String(ev.attachmentUrl ?? '').trim();
  const commentHtml = attachmentFileName && attachmentUrl
    ? buildCommentHtmlWithAttachment(ev.comment, attachmentFileName, attachmentUrl)
    : ev.comment;
  const assignedTo = String(ev.assignedToName ?? '').trim();

  return createSavedCommentItem(base, taskId, 'Request for Action', {
    status: '',
    commentHtml,
    commentCategory: ev.category,
    attachmentFileName,
    attachmentUrl,
    assignedTo,
    category: '',
  }, ctx, ev.sharePointItemId);
}

export function buildDefaultSavedCommentCard(
  ev: NewCommentSavedEvent,
  base: CommentListItem,
  taskId: string,
  ctx: SavedCommentCardContext,
): CommentListItem {
  return createSavedCommentItem(base, taskId, base.name ?? 'Comment', {
    status: base.eFormDetails?.status ?? '',
    commentHtml: ev.comment,
    commentCategory: ev.category,
    attachmentFileName: '',
    assignedTo: base.eFormDetails?.assignedTo ?? '',
    category: base.eFormDetails?.category ?? '',
  }, ctx, ev.sharePointItemId);
}

export function createSavedCommentItem(
  base: CommentListItem,
  taskId: string,
  name: string,
  details: SavedCommentDetails,
  ctx: SavedCommentCardContext,
  sharePointItemId?: string,
): CommentListItem {
  const commentAuthor = ctx.commentAuthor;
  const baseDetails = base.eFormDetails ?? {};
  const inheritedRawFields = (baseDetails['rawFields'] ?? {}) as Record<string, unknown>;
  return {
    ...base,
    id: sharePointItemId?.trim() || `comment:${taskId}:${Date.now()}`,
    name,
    submittedBy: commentAuthor,
    eFormDetails: {
      ...baseDetails,
      ...details,
      // Do not inherit parent task Status — comment cards hide status (and RFA uses its own).
      rawFields: {
        ...inheritedRawFields,
        Title: name,
        Status: details.status,
        LastState: details.status,
        ApprovalStatus: details.status,
        WorkflowStatus: details.status,
      },
      eFormListId: baseDetails.eFormListId ?? ctx.fallbackEFormListId,
      submitter: commentAuthor,
      commentSubmittedBy: commentAuthor,
      submittedBy: commentAuthor,
      submittedDate: new Date().toISOString(),
    },
    isContentLoaded: true,
  };
}
