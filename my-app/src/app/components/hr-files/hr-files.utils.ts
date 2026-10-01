export interface HrFileListItem {
  id: string;
  name: string;
  webUrl: string;
  isFolder: boolean;
  fullName?: string; // from AllDomainUsers.FullName
  pinNo?: string; // from AllDomainUsers.PinNo
}

export type HrFilesSortDirection = 'ascending' | 'descending';

export interface HrDomainUser {
  fullName: string;
  pinNo: string;
}

export interface HrDomainUserLookup {
  byEmail: Map<string, HrDomainUser>;
  byPin: Map<string, HrDomainUser>; // keyed by PinNo without leading zeros
}

/** Pulls the email out of a folder name such as "0 melanie.farrugia@domain.com". */
export function extractEmail(name: string): string {
  return (name ?? '').match(/[\w.+'-]+@[\w.-]+\.\w+/)?.[0]?.toLowerCase() ?? '';
}

/** Pulls the leading PIN out of a folder name such as "1000 mark.cauchi@domain.com" ("" for 0). */
export function extractPin(name: string): string {
  return ((name ?? '').trim().match(/^(\d+)\b/)?.[1] ?? '').replace(/^0+/, '');
}

/**
 * Returns copies of the items with fullName / pinNo filled from AllDomainUsers,
 * matched by email first and by the leading PIN when the email is not in the list.
 */
export function enrichHrFileItems(
  items: HrFileListItem[],
  users: HrDomainUserLookup
): HrFileListItem[] {
  if (!users.byEmail.size && !users.byPin.size) return items;
  return items.map(item => {
    const user = users.byEmail.get(extractEmail(item.name)) ?? users.byPin.get(extractPin(item.name));
    if (!user) return item;
    return { ...item, fullName: user.fullName || item.fullName, pinNo: user.pinNo || item.pinNo };
  });
}

/** "1000 mark.cauchi@domain.com" -> "Mark Cauchi"; "" when the folder name has no email. */
function nameFromFolderEmail(name: string): string {
  return extractEmail(name)
    .split('@')[0]
    .split(/[._-]+/)
    .filter(Boolean)
    .map(word => word[0].toUpperCase() + word.slice(1))
    .join(' ');
}

export function formatDisplayName(item: HrFileListItem): string {
  // Not in AllDomainUsers (e.g. a leaver): build the same "Name - PIN" from the folder name.
  const fullName = item.fullName || (item.pinNo ? '' : nameFromFolderEmail(item.name));
  const pinNo = item.pinNo || (fullName ? extractPin(item.name) : '');
  if (fullName && pinNo) return `${fullName} - ${pinNo}`;
  if (fullName) return fullName;
  if (pinNo) return pinNo;
  return item.name;
}

export function getDisplayInitial(item: HrFileListItem): string {
  // Initials come from the name only, so the PIN digits are never included
  const base = item.fullName?.trim() || item.name.split('@')[0].replace(/^\d+\s*/, '').replace(/[._-]+/g, ' ');
  const initials = base.split(/\s+/).filter(Boolean).map(w => w[0]).join('');
  return (initials || '?').slice(0, 2).toUpperCase();
}

function getHrFileSearchTokens(item: HrFileListItem): {
  raw: string;
  display: string;
  leadingId: string;
  email: string;
  emailLocal: string;
} {
  const raw = (item.name ?? '').trim();
  const display = formatDisplayName(item);
  const leadingId = raw.match(/^(\d+)\b/)?.[1] ?? '';
  const emailMatch = raw.match(/[\w.+-]+@[\w.-]+\.\w+/i);
  const email = emailMatch?.[0] ?? (raw.includes('@') ? raw : '');
  const emailLocal = email ? email.split('@')[0] : '';
  return { raw, display, leadingId, email, emailLocal };
}

function hrFileItemMatchesSearch(item: HrFileListItem, query: string): boolean {
  const q = query.trim().toUpperCase();
  if (!q) return true;

  const { raw, display, leadingId, email, emailLocal } = getHrFileSearchTokens(item);
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

function hrFileItemSearchScore(item: HrFileListItem, query: string): number {
  const q = query.trim().toUpperCase();
  if (!q) return 0;

  const { raw, display, leadingId, email, emailLocal } = getHrFileSearchTokens(item);
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
    formatDisplayName(a).localeCompare(
      formatDisplayName(b),
      undefined,
      { sensitivity: 'base' }
    );

  let filtered = items;
  if (trimmedQuery) {
    filtered = items.filter(item => hrFileItemMatchesSearch(item, trimmedQuery));
    return [...filtered].sort((a, b) => {
      const scoreDiff = hrFileItemSearchScore(b, trimmedQuery) - hrFileItemSearchScore(a, trimmedQuery);
      return scoreDiff !== 0 ? scoreDiff : compareByName(a, b);
    });
  }

  return [...filtered].sort(compareByName);
}
