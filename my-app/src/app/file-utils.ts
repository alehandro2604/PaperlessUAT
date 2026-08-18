// ============================================================
// FILE UTILITIES
// Pure functions for deriving file extension, category and
// icon path from a filename.  No Angular dependencies.
// Usage:  import { getFileExtension, getFileCategory, getFileIcon, normalizeSharePointFileUrl } from './file-utils';
// ============================================================

import { sharePointConfig } from './sharepoint.config';

export function normalizeSharePointFileUrl(url: string): string {
  let normalized = String(url ?? '').trim().replace(/&amp;/g, '&');
  if (!normalized) return '';

  // Undo double-encoding from encodeURI() on an already-encoded SharePoint URL (%20 -> %2520).
  let previous = '';
  while (previous !== normalized && /%25/i.test(normalized)) {
    previous = normalized;
    try {
      normalized = decodeURIComponent(normalized);
    } catch {
      break;
    }
  }

  // Server-relative paths (/sites/...) must open on SharePoint, not localhost:4200.
  if (normalized.startsWith('/')) {
    normalized = `https://${sharePointConfig.siteHostName}${normalized}`;
  }

  try {
    const parsed = new URL(normalized);
    const path = decodeURIComponent(parsed.pathname).toLowerCase();
    const isOfficeDoc = /\.(docx?|xlsx?|pptx?|odt|ods|odp)$/.test(path);
    if (isOfficeDoc && !parsed.searchParams.has('web')) {
      parsed.searchParams.set('web', '1');
    }
    return parsed.toString();
  } catch {
    return normalized;
  }
}

export function getFileExtension(filename: string): string {
  const lastDot = filename.lastIndexOf('.');
  return lastDot === -1 ? '' : filename.substring(lastDot + 1).toLowerCase();
}

export function getFileCategory(extension: string): string {
  if (['pdf', 'doc', 'docx', 'xml', 'txt'].includes(extension))        return 'document';
  if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg'].includes(extension)) return 'image';
  if (['mp4', 'avi', 'mov', 'wmv', 'flv'].includes(extension))         return 'video';
  if (['mp3', 'wav', 'aac', 'flac', 'ogg'].includes(extension))        return 'audio';
  if (['zip', 'rar', '7z', 'tar', 'gz'].includes(extension))           return 'archive';
  return 'other';
}

export function getFileIcon(category: string, extension: string): string {
  const byExtension: Record<string, string> = {
    pdf:  '/pdf.png',
    doc:  '/doc.png',
    docx: '/doc.png',
    xls:  '/xls.png',
    xlsx: '/xls.png',
  };

  const byCategory: Record<string, string> = {
    document: '/doc.png',
    image:    '/photo.png',
    video:    '/files.png',
    audio:    '/files.png',
    archive:  '/files.png',
    other:    '/files.png',
  };

  return byExtension[extension] ?? byCategory[category] ?? '/files.png';
}