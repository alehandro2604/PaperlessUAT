export interface HrFileListItem {
  id: string;
  name: string;
  webUrl: string;
  isFolder: boolean;
}

export type HrFilesSortDirection = 'ascending' | 'descending';

/** Turn an email-style folder name into a readable display name. */
export function formatDisplayNameFromEmail(name: string): string {
  const value = (name ?? '').trim();
  const toDisplayToken = (token: string): string => {
    const t = (token ?? '').trim();
    if (!t) return '';
    if (/^\d+$/.test(t)) return t;
    if (/^[A-Z0-9]{2,}$/.test(t)) return t;
    if (t.length <= 2) return t.toUpperCase();
    return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
  };

  const splitWhitespace = (chunk: string): string[] =>
    (chunk ?? '')
      .split(/\s+/)
      .map(s => s.trim())
      .filter(Boolean);

  if (!value.includes('@')) {
    return value
      .split(/\s+/)
      .flatMap(splitWhitespace)
      .map(toDisplayToken)
      .filter(Boolean)
      .join(' ');
  }

  const local = value.split('@')[0] ?? value;
  return local
    .split(/[._-]+/)
    .flatMap(splitWhitespace)
    .map(toDisplayToken)
    .join(' ');
}

// make this to display only the initials and not numbers
export function getDisplayInitial(name: string): string {
  const display = formatDisplayNameFromEmail(name);
  return (display.match(/[A-Z]/g)?.join('') || '?').toUpperCase();
}


function getHrFileSearchTokens(name: string): {
  raw: string;
  display: string;
  leadingId: string;
  email: string;
  emailLocal: string;
} {
  const raw = (name ?? '').trim();
  const display = formatDisplayNameFromEmail(raw);
  const leadingId = raw.match(/^(\d+)\b/)?.[1] ?? '';
  const emailMatch = raw.match(/[\w.+-]+@[\w.-]+\.\w+/i);
  const email = emailMatch?.[0] ?? (raw.includes('@') ? raw : '');
  const emailLocal = email ? email.split('@')[0] : '';
  return { raw, display, leadingId, email, emailLocal };
}

function hrFileItemMatchesSearch(name: string, query: string): boolean {
  const q = query.trim().toUpperCase();
  if (!q) return true;

  const { raw, display, leadingId, email, emailLocal } = getHrFileSearchTokens(name);
  const fields = new Set<string>(
    [
      raw,
      display,
      leadingId,
      email,
      emailLocal,
      ...display.split(/\s+/),
      ...emailLocal.split(/[._-]+/),
    ]
      .map(field => field.trim().toUpperCase())
      .filter(Boolean)
  );
  return [...fields].some(field => field.includes(q));
}

function hrFileItemSearchScore(name: string, query: string): number {
  const q = query.trim().toUpperCase();
  if (!q) return 0;

  const { raw, display, leadingId, email, emailLocal } = getHrFileSearchTokens(name);
  const rawUp = raw.toUpperCase();
  const displayUp = display.toUpperCase();
  const emailUp = email.toUpperCase();
  const emailLocalUp = emailLocal.toUpperCase();
  const leadingIdUp = leadingId.toUpperCase();

  if (leadingIdUp && leadingIdUp === q) return 1000;
  if (leadingIdUp && leadingIdUp.startsWith(q)) return 900;
  if (rawUp.startsWith(q)) return 850;
  if (displayUp.startsWith(q)) return 800;
  if (emailLocalUp.startsWith(q)) return 750;
  if (emailUp.startsWith(q)) return 700;
  if (displayUp.split(/\s+/).some(word => word.startsWith(q))) return 650;
  if (emailLocalUp.split(/[._-]+/).some(part => part.startsWith(q))) return 600;
  if (emailUp.includes(q)) return 400;
  if (displayUp.includes(q)) return 350;
  if (leadingIdUp && leadingIdUp.includes(q)) return 200;
  if (rawUp.includes(q)) return 100;
  return 0;
}

export function filterAndSortHrFileItems(
  items: HrFileListItem[],
  query: string,
  sortDirection: HrFilesSortDirection
): HrFileListItem[] {
  const trimmedQuery = query.trim();
  const direction = sortDirection === 'ascending' ? 1 : -1;
  const compareByName = (a: HrFileListItem, b: HrFileListItem) =>
    direction *
    formatDisplayNameFromEmail(a.name).localeCompare(
      formatDisplayNameFromEmail(b.name),
      undefined,
      { sensitivity: 'base' }
    );

  let filtered = items;
  if (trimmedQuery) {
    filtered = items.filter(item => hrFileItemMatchesSearch(item.name, trimmedQuery));
    return [...filtered].sort((a, b) => {
      const scoreDiff = hrFileItemSearchScore(b.name, trimmedQuery) - hrFileItemSearchScore(a.name, trimmedQuery);
      return scoreDiff !== 0 ? scoreDiff : compareByName(a, b);
    });
  }

  return [...filtered].sort(compareByName);
}
