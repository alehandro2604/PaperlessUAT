// Pure person/name-matching helpers extracted from AppComponent.
// None of these touch component state or injected services - verified via a full
// this-reference dependency trace before extraction.

export function cleanParsedPersonName(name: string): string {
  return name
    .replace(/\s+please click\b.*$/i, '')
    .replace(/\s+id=.*$/i, '')
    .replace(/\s+stepno:.*$/i, '')
    .trim();
}

export function extractPersonName(val: any): string {
  if (!val) return '';
  if (typeof val === 'string') return val.trim();
  if (Array.isArray(val) && val.length > 0) {
    return val
      .map((p: any) => extractPersonName(p))
      .filter(Boolean)
      .join(', ');
  }
  if (typeof val === 'object') {
    // Prefer the first *non-empty* identity field. Graph often returns
    // LookupValue: "" while Email/UPN is populated - `??` would stop on "".
    const person = val as Record<string, any>;
    const nestedUser = person['user'] ?? person['User'] ?? {};
    const candidates = [
      person['LookupValue'],
      person['displayName'],
      person['DisplayName'],
      person['Title'],
      person['EMail'],
      person['Email'],
      person['email'],
      person['UserPrincipalName'],
      person['userPrincipalName'],
      person['name'],
      person['Name'],
      nestedUser['displayName'],
      nestedUser['email'],
      nestedUser['userPrincipalName'],
    ];
    for (const candidate of candidates) {
      const text = String(candidate ?? '').trim();
      if (text) return text;
    }
    return '';
  }
  return '';
}

export function getEmailFromHrFolderName(folderName: string): string {
  return String(folderName ?? '').match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i)?.[0]?.toLowerCase() ?? '';
}


/** True when Title contains the folder person's email (increment / probation/performance review). */
export function doesTaskTitleMatchFolderPerson(
  title: unknown,
  folderName: string,
  hints: string[] = [],
): boolean {
  const text = normalizeTaskMatchText(stringifyTaskFieldValue(title));
  if (!text) return false;

  const emails = [
    getEmailFromHrFolderName(folderName),
    ...hints.filter(hint => hint.includes('@')),
  ]
    .map(email => normalizeTaskMatchText(email))
    .filter(email => email.length >= 5);

  if (emails.some(email => text.includes(email))) return true;

  const locals = hints
    .map(hint => (hint.includes('@') ? hint.split('@')[0] : hint))
    .map(hint => normalizeTaskMatchText(hint))
    .filter(hint => hint.length >= 5);

  if (locals.some(local => text.includes(local))) return true;

  // Title is often "{pin} email@domain" — match pin + folder name when email lookup fails.
  const folderPin = String(folderName ?? '').match(/^\d+/)?.[0] ?? '';
  if (folderPin && text.includes(normalizeTaskMatchText(folderPin))) {
    const nameParts = extractPersonNameParts(
      String(folderName ?? '')
        .replace(/^\d+\s+/, '')
        .replace(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/gi, ''),
    );
    if (nameParts.length >= 2) {
      const hits = nameParts.filter(part => part.length >= 4 && text.includes(part));
      if (hits.length >= Math.min(2, nameParts.length)) return true;
    }
  }

  return false;
}

export function normalizeTaskMatchText(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, '');
}

export function fieldMatchesFolderToken(normalizedText: string, token: string): boolean {
  if (!normalizedText || !token) return false;
  if (/^\d+$/.test(token)) {
    if (normalizedText === token) return true;
    if (normalizedText.startsWith(token) && normalizedText.length > token.length) {
      const next = normalizedText.charAt(token.length);
      if (next >= '0' && next <= '9') return false;
      return true;
    }
    return false;
  }
  return normalizedText.includes(token);
}

export function extractPersonNameParts(value: string): string[] {
  return String(value ?? '')
    .toLowerCase()
    .split(/[\s,._@+\-]+/)
    .map(part => normalizeTaskMatchText(part.trim()))
    .filter(part => part.length >= 4 && !/^\d+$/.test(part));
}

export function getHrPersonalFolderMatchTokens(folderName: string): string[] {
  const raw = String(folderName ?? '').trim().toLowerCase();
  if (!raw) return [];

  const withoutLeadingNumber = raw.replace(/^\d+\s+/, '').trim();
  const employeeId = raw.match(/^\d+/)?.[0] ?? '';
  const emailMatch = raw.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i)?.[0] ?? '';
  const emailLocal = emailMatch ? emailMatch.split('@')[0] : '';
  const displayFromEmail = emailLocal.replace(/[._-]+/g, ' ');
  const nameParts = extractPersonNameParts(withoutLeadingNumber);

  return [...new Set(
    [
      withoutLeadingNumber,
      ...nameParts,
      emailMatch,
      emailLocal,
      displayFromEmail,
      employeeId,
    ]
      .map(token => normalizeTaskMatchText(token))
      .filter(token => token.length >= 4 || (/^\d+$/.test(token) && token.length >= 2))
  )];
}

export function candidatesMatchFolderPerson(
  candidates: unknown[],
  folderName: string,
): boolean {
  const folderTokens = getHrPersonalFolderMatchTokens(folderName);
  if (folderTokens.length === 0) return false;

  const nameTokens = folderTokens.filter(token => !/^\d+$/.test(token));
  const idTokens = folderTokens.filter(token => /^\d+$/.test(token));
  const rawCandidates = candidates
    .map(value => stringifyTaskFieldValue(value).trim())
    .filter(Boolean);
  if (rawCandidates.length === 0) return false;

  const normalizedCandidates = rawCandidates
    .map(value => normalizeTaskMatchText(value))
    .filter(Boolean);
  const candidateParts = new Set(rawCandidates.flatMap(value => extractPersonNameParts(value)));

  const matchesLongToken = (token: string) =>
    normalizedCandidates.some(text => fieldMatchesFolderToken(text, token));

  if (nameTokens.length > 0) {
    // Combined first+last / email-local (e.g. "adriankind") — precise enough alone.
    if (nameTokens.some(token => token.length >= 8 && matchesLongToken(token))) {
      return true;
    }

    // Individual parts must match as whole tokens so "adrian" does not hit "adriana".
    const folderParts = extractPersonNameParts(
      String(folderName ?? '').replace(/^\d+\s+/, ''),
    );
    if (folderParts.length === 0) return false;

    const partHits = folderParts.filter(part => candidateParts.has(part));
    const requiredHits = Math.min(2, folderParts.length);
    return partHits.length >= requiredHits;
  }

  return idTokens.some(token =>
    normalizedCandidates.some(text => fieldMatchesFolderToken(text, token)),
  );
}

export function isItemSubmittedByFolderPerson(
  item: { submittedBy?: string; eFormDetails?: Record<string, unknown> },
  folderName: string
): boolean {
  const rawFields = (item.eFormDetails?.['rawFields'] ?? {}) as Record<string, unknown>;
  return candidatesMatchFolderPerson([
    item.submittedBy,
    item.eFormDetails?.['submitter'],
    item.eFormDetails?.['submittedBy'],
    item.eFormDetails?.['commentSubmittedBy'],
    rawFields['Requestor'],
    rawFields['SubmittedBy'],
    rawFields['Submitter'],
    rawFields['Author'],
    rawFields['CreatedBy'],
    rawFields['commentSubmittedBy'],
    rawFields['EmployeeName'],
    rawFields['Employee'],
    rawFields['EmployeeEmail'],
  ], folderName);
}

export function collectTaskAssigneeValues(fields: Record<string, any>): string[] {
  const candidates = [
    fields['AssignedTo'],
    fields['Assigned'],
    fields['AssignedTo0'],
    fields['CurrentAssignee'],
    fields['TaskAssignee'],
  ];

  const field7 = String(fields['field_7'] ?? '').trim();
  if (field7 && !/^\d{1,2}:\d{2}/.test(field7)) {
    candidates.push(fields['field_7']);
  }

  const values = candidates
    .map(value => extractPersonName(value))
    .filter(Boolean);

  return [...new Set(values)];
}

export function stringifyTaskFieldValue(value: unknown): string {
  if (value == null) return '';
  if (Array.isArray(value)) return value.map(v => stringifyTaskFieldValue(v)).filter(Boolean).join(' ');
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>;
    return String(
      record['LookupValue'] ??
      record['Title'] ??
      record['DisplayName'] ??
      record['displayName'] ??
      record['Email'] ??
      record['email'] ??
      record['UserPrincipalName'] ??
      record['userPrincipalName'] ??
      record['name'] ??
      record['value'] ??
      ''
    );
  }
  return String(value);
}
