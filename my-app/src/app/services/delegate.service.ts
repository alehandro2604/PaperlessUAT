import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { BehaviorSubject, Observable } from 'rxjs';
import { firstValueFrom } from 'rxjs';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { graphGet, graphPatch, toGraphPath } from '../microsoft-graph';
import { AuthService } from './auth.service';
import { UserService } from './user.service';
import { SiteMetadataService } from './site-metadata.service';

// Shape of a task passed into updateAssignedTo(). We only need the item id, the
// list it lives in, and the raw SharePoint fields so we can detect the correct
// assignee column and its current value.
export interface DelegationTask {
  id: string;
  listName?: string;
  eFormDetails?: {
    listName?: string;
    rawFields?: Record<string, unknown>;
    assignedTo?: string;
  };
}

// Resolved for the task's person column. `field` is the internal column
// name (e.g. "AssignedTo"), `lookupId` is the API key used to write the user id
// (e.g. "AssignedToLookupId"), and `allowMultiple` says whether the column
// accepts more than one person (changes the payload shape).
interface AssigneeFieldKeys {
  field: string;
  lookupId: string;
  allowMultiple: boolean;
}

@Injectable({
  providedIn: 'root',
})
export class DelegateService {
  //this class is used to manage the delegate list and the users who can delegate tasks
  private allowedDelegates: string[] = [];//allowed delegates is an array of strings that contains the users who can delegate tasks
  private loaded = false;//loaded is a boolean that is used to check if the delegate list is loaded
  private readonly canDelegateSubject = new BehaviorSubject<boolean>(false);//canDelegateSubject is a subject that is used to set the canDelegate value

  constructor(//this constructor is used to inject the http, authService and userService
    private readonly http: HttpClient,
    private readonly authService: AuthService,//authService is used to acquire the SharePoint token
    private readonly userService: UserService,//userService is used to get the current user
    private readonly siteMetadata: SiteMetadataService,//shared site + list metadata (resolved once per session)
  ) {}//this constructor is used to inject the HttpClient, authService and userService

  /** Emits true when the signed-in user appears in the SharePoint delegate list. */
  get canDelegate$(): Observable<boolean> {
    return this.canDelegateSubject.asObservable();//this method is used to return the canDelegateSubject as an observable
  }

  get canDelegate(): boolean {
    return this.canDelegateSubject.value;//this method is used to return the canDelegate value
  }

  get isLoaded(): boolean {
    return this.loaded;//this method is used to return the loaded value
  }

  /** Load UsersWhoCanDelegate list items from SharePoint (safe to call multiple times). */
  async loadDelegates(): Promise<void> {
    if (!this.userService.getCurrentUser()) {//this method is used to check if the current user is logged in
      this.reset();
      return;
    }

    try {
      const token = await this.authService.acquireSharePointToken();//this method is used to acquire the SharePoint token
      const listName = sharePointConfig.delegateListDisplayName;//this method is used to get the list name from the sharePointConfig

      const site = await this.siteMetadata.resolve(token);//shared resolver — site + lists are fetched once per session
      if (!site?.siteId) {//this method is used to check if the site is loaded
        this.reset();
        return;
      }

      const targetList = await this.siteMetadata.findList(token, listName);
      if (!targetList?.id) {//this method is used to check if the target list is loaded
        console.warn(`Delegate list "${listName}" was not found on the SharePoint site.`);
        this.reset();
        return;
      }

      const items = await this.fetchAllListItems(site.siteId, targetList.id, token);//this method is used to fetch all the items from the target list
      this.allowedDelegates = items.flatMap(item =>
        this.extractDelegateIdentifiers(item.fields ?? {}),
      );
      this.loaded = true;
      this.canDelegateSubject.next(this.isCurrentUserInDelegateList());//this method is used to set the canDelegate value
    } catch (error) {
      console.warn('Failed to load delegate permissions from SharePoint.', error);
      this.reset();
    }
  }

  private reset(): void {//this method is used to reset the delegate list
    this.allowedDelegates = [];
    this.loaded = false;
    this.canDelegateSubject.next(false);
  }

  private async fetchAllListItems(
    siteId: string,
    listId: string,
    token: string,
  ): Promise<any[]> {
    const collected: any[] = [];
    let nextPath: string | null =
      `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=10000`;

    while (nextPath) {
      const page: any = await firstValueFrom(
        graphGet(this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs),
      );
      collected.push(...(page?.value ?? []));
      nextPath = toGraphPath(page?.['@odata.nextLink']);
    }

    return collected;
  }

  private extractDelegateIdentifiers(fields: Record<string, unknown>): string[] {
    const identifiers = new Set<string>();
    const directKeys = [//this array is used to store the direct keys from the fields
      //these are the keys that are used to store the delegate identifiers
      'Title',
      'Email',
      'UserEmail',
      'UserName',
      'Name',
      'User',
      'Delegate',
      'DelegateUser',
      'Person',
      'Employee',
    ];

    for (const key of directKeys) {//this method is used to add the identifier values to the identifiers set
      this.addIdentifierValues(identifiers, fields[key]);
    }

    for (const value of Object.values(fields)) {//this method is used to add the identifier values to the identifiers set
      if (value && typeof value === 'object') {
        this.addIdentifierValues(identifiers, value);
      }
    }

    return [...identifiers];
  }

  private addIdentifierValues(target: Set<string>, value: unknown): void {
    if (!value) return;

    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (trimmed) target.add(trimmed);
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(entry => this.addIdentifierValues(target, entry));
      return;
    }

    if (typeof value === 'object') {
      const person = value as Record<string, unknown>;
      for (const key of ['LookupValue', 'displayName', 'Email', 'email', 'name', 'title']) {
        const entry = person[key];
        if (typeof entry === 'string' && entry.trim()) {
          target.add(entry.trim());
        }
      }
    }
  }

  private isCurrentUserInDelegateList(): boolean {
    if (!this.loaded || this.allowedDelegates.length === 0) {
      return false;
    }

    const currentUserName = this.userService.getCurrentUserName();
    const currentUserEmail = this.userService.getCurrentUserEmail();
    const currentUserUpn = this.userService.getCurrentUser()?.userPrincipalName ?? '';

    return this.allowedDelegates.some(delegate =>
      this.userService.matchesAssigneeField(delegate, currentUserName, currentUserEmail) ||
      (currentUserUpn && delegate.toLowerCase() === currentUserUpn.toLowerCase()),
    );
  }

  //this method is used to load the domain users
  private cachedUsers: DomainUser[] | null = null;
  private usersLoadPromise: Promise<DomainUser[]> | null = null;
  
  //this method is used to load the domain users
  private async loadDomainUsers(): Promise<DomainUser[]> {
    if (this.cachedUsers) return this.cachedUsers;
    // de-dupe concurrent loads (e.g. fast typing before first load resolves)
    if (this.usersLoadPromise) return this.usersLoadPromise;
  
    this.usersLoadPromise = (async () => {
      const token = await this.authService.acquireSharePointToken();
      const siteId = await this.resolveSiteId(token);            // reuse existing helper
      const listName = sharePointConfig.allDomainUsersListDisplayName;
      const listId = await this.resolveListId(siteId, listName, token); // reuse existing helper
      const items = await this.fetchAllListItems(siteId, listId, token);
      const users = items.flatMap(item => this.extractDomainUsers(item.fields ?? {}));
      this.cachedUsers = this.dedupeUsers(users);
      return this.cachedUsers;
    })();
  
    try {
      return await this.usersLoadPromise;
    } finally {
      this.usersLoadPromise = null;
    }
  }
  
  async searchUsers(query: string): Promise<DomainUser[]> {
    const q = query?.trim().toLowerCase() ?? '';
    if (q.length < 2) return [];

    const users = await this.loadDomainUsers(); // fetched once, cached thereafter
    return users.filter(
      u =>
        (u.Title?.toLowerCase() ?? '').includes(q) ||
        (u.Email?.toLowerCase() ?? '').includes(q) ||
        (u.UserPrincipalName?.toLowerCase() ?? '').includes(q),
    );
  }

  /**
   * Maps one AllDomainUsers row → DomainUser.
   * SharePoint Title on this list is often a dept code (e.g. "EPD"), not a person
   * name — prefer FullName / ADmail like SubordinaryTaskService does.
   */
  private extractDomainUsers(fields: Record<string, unknown>): DomainUser[] {
    // Prefer real person columns; keep field_* as fallbacks for older list schemas.
    const title = this.getStringValue(fields, [
      'FullName',
      'DisplayName',
      'Name',
      'field_2',
      'field_6',
      'field_9',
      'field_13',
      'field_16',
    ]);
    const email = this.getStringValue(fields, [
      'ADmail',
      'Email',
      'EMail',
      'UserEmail',
      'field_3',
      'field_8',
      'field_11',
    ]);
    const upn = this.getStringValue(fields, [
      'UserPrincipalName',
      'UPN',
      'field_23',
      'ADmail',
      'Email',
    ]);

    if (!title && !email) return [];

    return [
      {
        Title: title || email,
        Email: email,
        UserPrincipalName: upn || email,
      },
    ];
  }

  /**
   * Collapses duplicate copies of the same person into a single record.
   * The source list stores each person across several columns (name-only and
   * name/email pairs), so extraction emits the same person multiple times.
   * We group by name, gather the distinct emails + UPN, and emit one record
   * per (name, email) — dropping redundant name-only copies when an emailed
   * copy for the same name exists.
   */
  private dedupeUsers(users: DomainUser[]): DomainUser[] {
    const byName = new Map<string, { name: string; emails: Map<string, string>; upn: string }>();

    for (const user of users) {
      const name = (user.Title ?? '').trim();
      if (!name) continue;

      const key = name.toLowerCase();
      let entry = byName.get(key);
      if (!entry) {
        entry = { name, emails: new Map(), upn: '' };
        byName.set(key, entry);
      }

      const email = (user.Email ?? '').trim();
      if (email) entry.emails.set(email.toLowerCase(), email);
      if (!entry.upn && user.UserPrincipalName) entry.upn = user.UserPrincipalName.trim();
    }

    const result: DomainUser[] = [];
    for (const entry of byName.values()) {
      if (entry.emails.size === 0) {
        result.push({ Title: entry.name, Email: '', UserPrincipalName: entry.upn });
      } else {
        for (const email of entry.emails.values()) {
          result.push({ Title: entry.name, Email: email, UserPrincipalName: entry.upn });
        }
      }
    }

    return result;
  }

  /** PATCH the SharePoint list item assignee field. Throws when the write cannot be verified. */
  /**
   * Writes the task assignee. With `requireSoleAssignee` (Claim), success also
   * needs every other assignee gone, not just the new one present.
   */
  async updateAssignedTo(
    task: DelegationTask,
    selectedUser: DomainUser,
    options: { requireSoleAssignee?: boolean } = {},
  ): Promise<void> {
    const itemId = task?.id;
    const listName = task?.listName || task?.eFormDetails?.listName || '';
    const newAssigneeName = selectedUser.Title?.trim() ?? '';
    const newAssigneeEmail = (selectedUser.Email || selectedUser.UserPrincipalName || '').trim();

    if (!itemId) throw new Error('Cannot delegate: task has no id');
    if (!listName) throw new Error('Cannot delegate: task has no listName');
    if (!newAssigneeEmail) {
      throw new Error(`No email found for ${newAssigneeName || 'selected user'}`);
    }

    const graphToken = await this.authService.acquireSharePointWriteToken();
    if (!graphToken) throw new Error('Unable to acquire SharePoint write token');

    // Resolve the site, the target list, and which person column we are writing.
    const siteId = await this.resolveSiteId(graphToken);
    const listId = await this.resolveListId(siteId, listName, graphToken);
    const assigneeKeys = await this.resolveAssigneeFieldKeys(
      siteId,
      listId,
      task.eFormDetails?.rawFields,
      task.eFormDetails?.assignedTo,
      graphToken,
    );

    // Claims-format login names used by ensureuser to resolve/provision the user.
    const loginNames = [
      `i:0#.f|membership|${newAssigneeEmail}`,//membership is the format of the email for the user
      selectedUser.UserPrincipalName
        ? `i:0#.f|membership|${selectedUser.UserPrincipalName}`// Format for the name
        : null,
    ].filter(Boolean) as string[];

    // A person column stores a numeric site user id, not a name/email. ensureuser
    // both provisions the user into the site (if new) and returns that id.
    let spUserId: number | null = null;

    for (const loginName of loginNames) {
      try {
        const restToken = await this.authService.acquireSharePointRestWriteToken();
        spUserId = await this.ensureSharePointSiteUser(restToken, loginName);
        break;
      } catch (error) {
        console.warn('SharePoint ensureUser failed for login name:', loginName, error);
      }
    }

    // Fallback: if ensureuser failed (e.g. REST permission), scan the hidden
    // User Information List for an existing row.
    if (!spUserId) {
      spUserId = await this.resolveSharePointUserId(
        siteId,
        graphToken,
        newAssigneeEmail,
        newAssigneeName,
      );
    }

    if (!spUserId) {
      throw new Error(
        `Could not resolve SharePoint user id for ${newAssigneeEmail}. The user may need site access first.`,
      );
    }

    let lastError: unknown = null;

    // Primary path: write via Microsoft Graph. We try each valid payload shape
    // and stop as soon as a read-back confirms the assignee actually changed.
    const patchAttempts = this.buildGraphPatchAttempts(assigneeKeys, spUserId);

    for (const body of patchAttempts) {
      try {
        await firstValueFrom(
          graphPatch(
            this.http,
            `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
            body,
            graphToken,
          ),
        );

        // Graph can return success without applying the change, so always verify.
        if (
          await this.verifyAssignedToUpdate(
            siteId,
            listId,
            itemId,
            assigneeKeys,
            graphToken,
            newAssigneeEmail,
            newAssigneeName,
            spUserId,
            task.eFormDetails?.assignedTo,
            options.requireSoleAssignee,
          )
        ) {
          return;
        }

        console.warn('Graph PATCH accepted but AssignedTo did not change.', body);
      } catch (error) {
        lastError = error;
        console.warn('Graph PATCH failed:', this.formatGraphError(error), body);
      }
    }

    // Fallback path: write via SharePoint REST (_api) with a MERGE, then verify.
    try {
      const restToken = await this.authService.acquireSharePointRestWriteToken();
      await this.updateAssignedToViaRest(restToken, listName, itemId, assigneeKeys, spUserId);

      if (
        await this.verifyAssignedToUpdate(
          siteId,
          listId,
          itemId,
          assigneeKeys,
          graphToken,
          newAssigneeEmail,
          newAssigneeName,
          spUserId,
          task.eFormDetails?.assignedTo,
          options.requireSoleAssignee,
        )
      ) {
        return;
      }
    } catch (error) {
      lastError = error;
      console.warn('SharePoint REST update failed:', error);
    }

    // Nothing applied verifiably — surface a failure so the UI does not show a
    // fake success (which would later revert on refresh).
    throw new Error(
      `Could not update AssignedTo for ${newAssigneeEmail} on list "${listName}". ${this.formatGraphError(lastError)}`.trim(),
    );
  }

  /** Absolute site URL used for SharePoint REST (_api) calls. */
  private getSiteWebUrl(): string {
    return `https://${sharePointConfig.siteHostName}/${sharePointConfig.sitePath}`;
  }

  /**
   * Builds the candidate Graph PATCH bodies for a person column. Multi-value
   * columns expect an Int32 collection on the LookupId key; single-value columns
   * expect a plain integer. We try the array shapes first, then the scalar.
   */
  private buildGraphPatchAttempts(
    assigneeKeys: AssigneeFieldKeys,
    spUserId: number,
  ): Record<string, unknown>[] {
    const { field, lookupId, allowMultiple } = assigneeKeys;
    const attempts: Record<string, unknown>[] = [];

    if (allowMultiple) {
      attempts.push({
        [`${field}@odata.type`]: 'Collection(Edm.Int32)',
        [lookupId]: [spUserId],
      });
      attempts.push({
        [`${lookupId}@odata.type`]: 'Collection(Edm.Int32)',
        [lookupId]: [spUserId],
      });
    }

    attempts.push({ [lookupId]: spUserId });
    return attempts;
  }

  /** Re-reads the item and confirms the assignee actually became the new user. */
  private async verifyAssignedToUpdate(
    siteId: string,
    listId: string,
    itemId: string,
    assigneeKeys: AssigneeFieldKeys,
    token: string,
    email: string,
    name: string,
    spUserId: number,
    previousAssignee?: string,
    requireSoleAssignee = false,
  ): Promise<boolean> {
    const readBack: any = await firstValueFrom(
      graphGet(
        this.http,
        `/sites/${siteId}/lists/${listId}/items/${itemId}/fields?$select=${assigneeKeys.field},${assigneeKeys.lookupId}`,
        token,
      ),
    );

    if (requireSoleAssignee) {
      return this.readBackIsSoleAssignee(readBack ?? {}, assigneeKeys, email, name, spUserId);
    }

    return this.readBackMatchesAssignee(
      readBack ?? {},
      assigneeKeys,
      email,
      name,
      spUserId,
      previousAssignee,
    );
  }

  /**
   * Calls SharePoint REST `ensureuser`, which guarantees the user exists in the
   * site's User Information List (creating the row if needed) and returns their
   * numeric site user id — the value a person column actually stores.
   */
  private async ensureSharePointSiteUser(token: string, logonName: string): Promise<number> {
    const resp: any = await firstValueFrom(
      this.http.post(
        `${this.getSiteWebUrl()}/_api/web/ensureuser`,
        JSON.stringify({ logonName }),
        {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
          },
        },
      ),
    );

    const id = resp?.d?.Id ?? resp?.d?.id;
    if (id == null || id === '') {
      throw new Error(`ensureUser did not return an id for ${logonName}`);
    }

    return parseInt(String(id), 10);
  }

  /**
   * Fallback write through SharePoint REST. Uses the `<field>Id` convention
   * (e.g. AssignedToId) with a MERGE so it updates the existing item in place.
   * Multi-value columns take a `{ results: [...] }` wrapper.
   */
  private async updateAssignedToViaRest(
    token: string,
    listName: string,
    itemId: string,
    assigneeKeys: AssigneeFieldKeys,
    spUserId: number,
  ): Promise<void> {
    const escapedListName = listName.replace(/'/g, "''");
    const fieldIdKey = `${assigneeKeys.field}Id`;
    const body = assigneeKeys.allowMultiple
      ? { [fieldIdKey]: { results: [spUserId] } }
      : { [fieldIdKey]: spUserId };

    await firstValueFrom(
      this.http.post(
        `${this.getSiteWebUrl()}/_api/web/lists/getbytitle('${escapedListName}')/items(${itemId})`,
        JSON.stringify(body),
        {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
        },
      ),
    );
  }

  /** Resolves the Graph site id from the configured host name and site path. */
  private async resolveSiteId(token: string): Promise<string> {
    const { siteId } = await this.siteMetadata.resolve(token);
    if (!siteId) throw new Error('Could not resolve SharePoint site ID');
    return siteId;
  }

  /** Finds the list id by matching the task's list name or display name. */
  private async resolveListId(_siteId: string, listName: string, token: string): Promise<string> {
    const targetList = await this.siteMetadata.findList(token, listName);
    if (!targetList?.id) throw new Error(`List "${listName}" not found on site`);
    return targetList.id;
  }

  /**
   * Works out which column holds the assignee and how to write it. We first try
   * to match a candidate field whose current value equals the known assignee,
   * then confirm against the live column metadata (real internal name + whether
   * it allows multiple people).
   */
  private async resolveAssigneeFieldKeys(
    siteId: string,
    listId: string,
    rawFields?: Record<string, unknown>,
    currentAssignee?: string,
    token?: string,
  ): Promise<AssigneeFieldKeys> {
    const candidates = ['AssignedTo', 'Assigned', 'field_7'];
    let field = 'AssignedTo';

    // Prefer the candidate whose stored value matches the current assignee.
    if (rawFields && currentAssignee?.trim()) {
      for (const candidate of candidates) {
        if (!Object.prototype.hasOwnProperty.call(rawFields, candidate)) continue;
        const value = this.extractPersonFieldValue(rawFields[candidate]);
        if (value && this.userService.matchesAssigneeField(value, currentAssignee, currentAssignee)) {
          field = candidate;
          break;
        }
      }
    } else {
      for (const candidate of candidates) {
        if (rawFields && Object.prototype.hasOwnProperty.call(rawFields, candidate)) {
          field = candidate;
          break;
        }
      }
    }

    // Confirm the internal name and multi-select flag from column metadata.
    let allowMultiple = true;
    if (token) {
      try {
        const columnsResp: any = await firstValueFrom(
          graphGet(
            this.http,
            `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,personOrGroup`,
            token,
          ),
        );
        const column = (columnsResp?.value ?? []).find(
          (col: any) =>
            (col.name ?? '').toLowerCase() === field.toLowerCase() ||
            (col.displayName ?? '').toLowerCase() === 'assigned to',
        );
        if (column?.personOrGroup) {
          field = column.name ?? field;
          allowMultiple = column.personOrGroup.allowMultipleSelection !== false;
        }
      } catch (error) {
        console.warn('Could not read assignee column metadata:', error);
      }
    }

    return { field, lookupId: `${field}LookupId`, allowMultiple };
  }

  /** Looks up the GUID of the hidden User Information List (avoids spaced URLs). */
  private async resolveUserInformationListId(siteId: string, token: string): Promise<string | null> {
    const listsResp: any = await firstValueFrom(
      graphGet(this.http, `/sites/${siteId}/lists?$select=id,name,displayName`, token),
    );
    const userInfoList = (listsResp?.value ?? []).find(
      (list: any) =>
        (list.displayName ?? '').toLowerCase() === 'user information list' ||
        (list.name ?? '').toLowerCase() === 'user information list',
    );
    return userInfoList?.id ?? null;
  }

  /**
   * Fallback user-id resolver used when ensureuser is unavailable. Pages through
   * the User Information List by GUID and matches on email or display name.
   */
  private async resolveSharePointUserId(
    siteId: string,
    token: string,
    email: string,
    displayName: string,
  ): Promise<number | null> {
    try {
      const userInfoListId = await this.resolveUserInformationListId(siteId, token);
      if (!userInfoListId) {
        console.warn('User Information List was not found on the SharePoint site.');
        return null;
      }

      const emailLower = email.toLowerCase();
      let nextPath: string | null =
        `/sites/${siteId}/lists/${userInfoListId}/items?$expand=fields($select=Id,EMail,Title)&$top=500`;

      // Walk every page until we find a matching user or run out of pages.
      while (nextPath) {
        const page: any = await firstValueFrom(graphGet(this.http, nextPath, token));
        const match = (page?.value ?? []).find((item: any) => {
          const itemEmail = (item.fields?.EMail ?? item.fields?.Email ?? '').toLowerCase();
          const itemTitle = (item.fields?.Title ?? '').trim();
          return (
            itemEmail === emailLower ||
            this.userService.matchesAssigneeField(itemTitle, displayName, email)
          );
        });
        const rawId = match?.fields?.Id;
        if (rawId != null && rawId !== '') {
          return parseInt(String(rawId), 10);
        }
        nextPath = toGraphPath(page?.['@odata.nextLink']);
      }
    } catch (error) {
      console.warn('Could not query User Information List:', error);
    }

    return null;
  }

  /**
   * Normalises a person field (string, array, or object) into a readable name or
   * email string, used both for matching and for comparison during verification.
   */
  private extractPersonFieldValue(value: unknown): string {
    if (!value) return '';
    if (typeof value === 'string') return value.trim();
    if (Array.isArray(value)) {
      return value
        .map(entry => this.extractPersonFieldValue(entry))
        .filter(Boolean)
        .join(', ');
    }
    if (typeof value === 'object') {
      const person = value as Record<string, unknown>;
      return String(
        person['LookupValue'] ??
          person['displayName'] ??
          person['DisplayName'] ??
          person['Email'] ??
          person['email'] ??
          person['Title'] ??
          '',
      ).trim();
    }
    return '';
  }

  /**
   * Decides whether a read-back proves the write succeeded. Crucially, it rejects
   * the case where the column still holds the previous assignee (the old bug that
   * reported false success), and matches by name, email, or lookup id.
   */
  private readBackMatchesAssignee(
    fields: Record<string, unknown>,
    assigneeKeys: AssigneeFieldKeys,
    email: string,
    name: string,
    spUserId: number | null,
    previousAssignee?: string,
  ): boolean {
    const assignedText = this.extractPersonFieldValue(fields[assigneeKeys.field]);
    const matchesNewAssignee = this.userService.matchesAssigneeField(assignedText, name, email);
    const stillPreviousAssignee =
      !!previousAssignee?.trim() &&
      this.userService.matchesAssigneeField(assignedText, previousAssignee);

    // Unchanged value (still the old assignee) means the write did not take.
    if (stillPreviousAssignee && !matchesNewAssignee) {
      return false;
    }

    if (matchesNewAssignee) {
      return true;
    }

    // Fall back to inspecting each person entry by email, name, or lookup id.
    const assignedEntries = fields[assigneeKeys.field];
    if (Array.isArray(assignedEntries)) {
      for (const entry of assignedEntries) {
        if (entry && typeof entry === 'object' && this.personEntryMatches(entry, email, name, spUserId)) {
          return true;
        }
      }
    }

    const lookupId = fields[assigneeKeys.lookupId];
    if (spUserId) {
      if (Number(lookupId) === spUserId) return true;
      if (Array.isArray(lookupId) && lookupId.some(id => Number(id) === spUserId)) return true;
    }

    return false;
  }

  /**
   * Claim check: the read-back must hold the new assignee and nobody else. The
   * claimer was already one of the assignees, so "is among them" proves nothing.
   */
  private readBackIsSoleAssignee(
    fields: Record<string, unknown>,
    assigneeKeys: AssigneeFieldKeys,
    email: string,
    name: string,
    spUserId: number | null,
  ): boolean {
    const entries = fields[assigneeKeys.field];
    if (Array.isArray(entries) && entries.length > 0) {
      return entries.every(entry =>
        entry && typeof entry === 'object'
          ? this.personEntryMatches(entry, email, name, spUserId)
          : this.userService.matchesAssigneeField(String(entry ?? ''), name, email),
      );
    }

    const lookupId = fields[assigneeKeys.lookupId];
    if (Array.isArray(lookupId) && lookupId.length > 0) {
      return !!spUserId && lookupId.every(id => Number(id) === spUserId);
    }
    if (lookupId != null && lookupId !== '') {
      return !!spUserId && Number(lookupId) === spUserId;
    }

    // A single-person column can only ever hold one person.
    if (!assigneeKeys.allowMultiple) {
      return this.readBackMatchesAssignee(fields, assigneeKeys, email, name, spUserId);
    }
    return false;
  }

  /** True when one person entry from a read-back is the given user (email, name or lookup id). */
  private personEntryMatches(entry: object, email: string, name: string, spUserId: number | null): boolean {
    const person = entry as Record<string, unknown>;
    const personEmail = String(person['Email'] ?? person['email'] ?? '').toLowerCase();
    const personName = String(
      person['LookupValue'] ?? person['displayName'] ?? person['DisplayName'] ?? person['Title'] ?? '',
    ).trim();
    const lookupId = Number(person['LookupId'] ?? person['Id'] ?? person['id']);

    if (personEmail && personEmail === email.toLowerCase()) return true;
    if (personName && this.userService.matchesAssigneeField(personName, name, email)) return true;
    return !!spUserId && lookupId === spUserId;
  }

  /** Extracts a human-readable message from Graph or SharePoint REST errors. */
  private formatGraphError(error: unknown): string {
    const err = error as any;
    return (
      err?.error?.error?.message ||
      err?.error?.['odata.error']?.message?.value ||
      err?.message ||
      ''
    );
  }

  private getStringValue(fields: Record<string, unknown>, keys: string[]): string {
    for (const key of keys) {
      const value = fields[key];
      if (typeof value === 'string' && value.trim()) {
        return value.trim();
      }
      if (value && typeof value === 'object') {
        const obj = value as Record<string, unknown>;
        for (const nestedKey of ['LookupValue', 'displayName', 'Email', 'email', 'value']) {
          const nestedValue = obj[nestedKey];
          if (typeof nestedValue === 'string' && nestedValue.trim()) {
            return nestedValue.trim();
          }
        }
      }
    }
    return '';
  }
}

export interface DomainUser {
  Title: string;
  Email: string;
  UserPrincipalName: string;
}

