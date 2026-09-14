// Pure HR-task matching/classification helpers extracted from AppComponent.
// None of these touch component state or injected services - verified via a full
// this-reference dependency trace before extraction.

import { sharePointConfig } from '../sharepoint.config';

export function isItemSubmittedBy(
  item: { submittedBy?: string; eFormDetails?: Record<string, unknown> },
  submitterName: string | null
): boolean {
  if (!submitterName) return true;

  const candidates = [
    item.submittedBy,
    item.eFormDetails?.['submitter'],
    item.eFormDetails?.['submittedBy'],
  ]
    .map(value => String(value ?? '').trim().toLowerCase())
    .filter(Boolean);

  const target = submitterName.trim().toLowerCase();
  return candidates.some(candidate =>
    candidate === target || candidate.includes(target) || target.includes(candidate)
  );
}

export function getHrTaskPersonLookupIds(item: any): string[] {
  const f = item?.fields ?? item ?? {};
  const candidates = [
    f.RequestorLookupId, f.AuthorLookupId, f.SubmittedByLookupId,
    f.EmployeeLookupId, f.EmployeeNameLookupId, f.SubmitterLookupId,
    f.AssignedToLookupId, f.AssignedLookupId,
    f.Requestor?.LookupId, f.Author?.LookupId, f.SubmittedBy?.LookupId,
    f.Employee?.LookupId, f.EmployeeName?.LookupId, f.Submitter?.LookupId,
    f.AssignedTo?.LookupId, f.Assigned?.LookupId,
    f.CreatedByLookupId, f.CreatedBy?.LookupId,
    item?.createdBy?.user?.id,
  ];
  return [...new Set(
    candidates
      .map(value => String(value ?? '').trim())
      .filter(Boolean),
  )];
}

export function doesHrTaskMatchPersonLookupId(item: any, lookupId: string | null): boolean {
  if (!lookupId) return false;
  return getHrTaskPersonLookupIds(item).includes(lookupId);
}

export function normalizeHrTaskChainName(value: string): string {
  return String(value ?? '')
    .toLowerCase()
    .replace(/please\s+approve\s+/g, ' ')
    .replace(/\bfinalis(?:ed|e|ation)?\b/g, ' ')
    .replace(/\beforms?\b/g, ' ')
    .replace(/\bfor\b\s*-?\s*.*$/g, ' ')
    .replace(/\bid\s*[=:]\s*\d+\b/g, ' ')
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}

export function getMappedHrTaskChainName(item: any): string {
  const details = item?.eFormDetails ?? {};
  const rawFields = (details.rawFields ?? {}) as Record<string, unknown>;
  const candidates = [
    rawFields['Title'],
    item?.name,
    details.type,
    details.listName,
  ];
  for (const candidate of candidates) {
    const normalized = normalizeHrTaskChainName(String(candidate ?? ''));
    if (normalized.length >= 4) return normalized;
  }
  return normalizeHrTaskChainName(String(item?.name ?? details.type ?? ''));
}

export function getRawHrTaskChainName(item: any): string {
  const f = item?.fields ?? item ?? {};
  return normalizeHrTaskChainName(String(f.Title ?? f.ContentType ?? ''));
}

export function hrTaskChainNamesMatch(seedName: string, candidateName: string): boolean {
  if (!seedName || !candidateName) return false;
  if (seedName === candidateName) return true;
  return seedName.includes(candidateName) || candidateName.includes(seedName);
}

export function doesRawItemReferenceEFormKeys(item: any, eFormKeys: Set<string>): boolean {
  if (!eFormKeys.size) return false;
  const f = item?.fields ?? item ?? {};
  const fieldVals = [f.eFormListId, f.ListId, f.field_11]
    .map(v => String(v ?? '').trim())
    .filter(Boolean);
  if (fieldVals.some(v => eFormKeys.has(v))) return true;

  const text = [f.Title, f.Comment, f.Notes, f.field_10]
    .map(v => String(v ?? ''))
    .join(' ');
  for (const key of eFormKeys) {
    if (new RegExp(`\\bID\\s*[=:]\\s*${key}\\b`, 'i').test(text)) return true;
  }
  return false;
}

export function doesRawItemMatchSeedEFormIdAndTaskName(
  item: any,
  eFormKeyToTaskNames: Map<string, Set<string>>,
): boolean {
  if (!eFormKeyToTaskNames.size) return false;
  const chainName = getRawHrTaskChainName(item);
  if (!chainName) return false;

  for (const [key, names] of eFormKeyToTaskNames) {
    if (!doesRawItemReferenceEFormKeys(item, new Set([key]))) continue;
    if ([...names].some(seedName => hrTaskChainNamesMatch(seedName, chainName))) {
      return true;
    }
  }
  return false;
}

export function extractEFormKeysFromMappedItem(item: any): string[] {
  const keys = new Set<string>();
  const details = item?.eFormDetails ?? {};
  const fieldId = String(details.eFormListId ?? '').trim();
  if (fieldId && fieldId.toLowerCase() !== 'unknown eform') {
    keys.add(fieldId);
  }

  // SharePoint task item id is often the same ID shown in "ID=9699" footers.
  const itemId = String(item?.id ?? '').trim();
  if (/^\d+$/.test(itemId)) {
    keys.add(itemId);
  }

  const text = [
    item?.description,
    details.comment,
    details.commentHtml,
    details.body,
    item?.name,
  ].map(v => String(v ?? '')).join(' ');

  for (const match of text.matchAll(/\bID\s*[=:]\s*(\d+)\b/gi)) {
    if (match[1]) keys.add(match[1]);
  }
  return [...keys];
}

export function doesMappedItemMatchSeedEFormIdAndTaskName(
  item: any,
  eFormKeyToTaskNames: Map<string, Set<string>>,
): boolean {
  if (!eFormKeyToTaskNames.size) return false;
  const chainName = getMappedHrTaskChainName(item);
  if (!chainName) return false;

  for (const key of extractEFormKeysFromMappedItem(item)) {
    const names = eFormKeyToTaskNames.get(key);
    if (!names?.size) continue;
    if ([...names].some(seedName => hrTaskChainNamesMatch(seedName, chainName))) {
      return true;
    }
  }
  return false;
}

export function collectEFormKeyToTaskNames(seedTasks: any[]): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>();
  for (const task of seedTasks) {
    const chainName = getMappedHrTaskChainName(task);
    if (!chainName) continue;
    for (const key of extractEFormKeysFromMappedItem(task)) {
      let names = map.get(key);
      if (!names) {
        names = new Set<string>();
        map.set(key, names);
      }
      names.add(chainName);
    }
  }
  return map;
}

export function buildPersonLookupFilterExpr(
  lookupField: string,
  lookupId: string,
  allowMultiple: boolean,
): string {
  const idLiteral = /^\d+$/.test(lookupId)
    ? lookupId
    : `'${String(lookupId).replace(/'/g, "''")}'`;
  if (allowMultiple) {
    return `fields/${lookupField}/any(a:a eq ${idLiteral})`;
  }
  return `fields/${lookupField} eq ${idLiteral}`;
}

export function isDocumentLibraryTaskList(listName: string): boolean {
  const normalized = listName.toLowerCase();
  return sharePointConfig.documentLibrariesTasks.some(
    (taskList: string) => taskList.toLowerCase() === normalized
  );
}

export function isChildSubjectColumn(columnName: string): boolean {
  const normalized = String(columnName ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (normalized === 'childsubject' || normalized === 'childsubject2') return true;
  // SharePoint encoded internal names, e.g. Child_x0020_Subject / Child_x0020_Subject2
  return /^childx0020subject2?$/.test(normalized);
}

export function getDateSearchTokens(dateStr: string): string {
  if (!dateStr) return '';
  const dt = new Date(dateStr);
  if (isNaN(dt.getTime())) return dateStr.toLowerCase();

  const day = dt.getDate().toString();
  const dayPadded = dt.toLocaleDateString('en-GB', { day: '2-digit' });
  const month2 = dt.toLocaleDateString('en-GB', { month: '2-digit' });
  const monthShort = dt.toLocaleDateString('en-GB', { month: 'short' }).toLowerCase();
  const monthLong = dt.toLocaleDateString('en-GB', { month: 'long' }).toLowerCase();
  const year = dt.getFullYear().toString();

  return [
    `${dayPadded}/${month2}/${year}`,
    `${day}/${month2}/${year}`,
    `${month2}/${year}`,
    `${monthShort} ${year}`,
    `${monthLong} ${year}`,
    `${day} ${monthShort}`,
    `${day} ${monthLong}`,
    monthShort,
    monthLong,
    year,
    day,
  ].join(' ');
}

export function getTaskEFormTitle(item: { eFormDetails?: Record<string, unknown>; name?: unknown }): string {
  const details = item.eFormDetails ?? {};
  const rawFields = (details['rawFields'] ?? {}) as Record<string, unknown>;
  const raw =
    rawFields['Title'] ??
    details['Title'] ??
    details['title'];
  return String(raw ?? '').trim() || 'Unknown Title';
}

export function extractTrailingFolderId(folderName: string): string {
  return String(folderName ?? '').trim().match(/(?:^|[-_\s])(\d+)\s*$/)?.[1] ?? '';
}
