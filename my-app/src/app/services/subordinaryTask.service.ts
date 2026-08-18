import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { async, firstValueFrom, from, Observable, of } from 'rxjs';
import { catchError } from 'rxjs/operators';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { graphGet, toGraphPath } from '../microsoft-graph';
import { AuthService } from './auth.service';
import { UserService } from './user.service';
import { SiteMetadataService } from './site-metadata.service';
import { ToDoTask } from '../types/todo-task.interface';

export interface DomainUser {
    PinNo: string;
    FullName: string;
    ADmail: string;
    Manager: string;
    ManagerEmail: string;
}

interface DomainUserColumnMap {
    PinNo: string[];
    FullName: string[];
    ADmail: string[];
    Manager: string[];
    ManagerEmail: string[];
}

interface SubordinateFolder {
    employee: DomainUser;
    tasks: ToDoTask[];
    pendingCount: number;
}

@Injectable({
    providedIn: 'root'
})

// This service is used to get the direct reports of a manager
export class SubordinaryTaskService {
    searchTerm = '';
    tasks: ToDoTask[] = [];
    private cachedUsers: DomainUser[] | null = null;
    private usersLoadPromise: Promise<DomainUser[]> | null = null;
    private cachedSubordinates: DomainUser[] | null = null;
    private subordinatesLoadPromise: Promise<DomainUser[]> | null = null;
    subordanteEmployees: any;
    expandedEmployeeFolders: any;

    // This method is used to get the direct reports of a manager
    constructor(
        private readonly http: HttpClient,
        private readonly authService: AuthService,
        private readonly userService: UserService,
        private readonly siteMetadata: SiteMetadataService,
        ) { }

    getCachedSubordinates(): DomainUser[] {
        return this.cachedSubordinates ?? [];
    }

    clearSubordinatesCache(): void {
        this.cachedSubordinates = null;
        this.subordinatesLoadPromise = null;
    }

    /** Build all manager lookup keys from login email/UPN/name (includes email local part, e.g. alehandro). */
    async buildManagerSearchKeys(rawIdentifiers: string[]): Promise<string[]> {
        const keys = new Set<string>();

        for (const raw of rawIdentifiers) {
            const trimmed = (raw ?? '').trim();
            if (!trimmed) continue;
            keys.add(trimmed);
            keys.add(trimmed.toLowerCase());

            if (trimmed.includes('@')) {
                const localPart = trimmed.split('@')[0]?.trim();
                if (localPart) {
                    keys.add(localPart);
                    keys.add(localPart.toLowerCase());
                }
            }
        }

        const users = await this.loadDomainUsers();
        for (const raw of rawIdentifiers) {
            const normalized = this.normalize(raw);
            if (!normalized) continue;

            // Get the local part of the email address ,local meaning the part before the @ symbol
            const localPart = normalized.includes('@') ? normalized.split('@')[0] : normalized;
            const self = users.find(user => {// Find the user in the list of usersLoadPromise: Promise<DomainUser[]> | null = null;
                const admail = this.normalize(user.ADmail);
                const admailLocal = admail.includes('@') ? admail.split('@')[0] : admail;
                return admail === normalized || admailLocal === localPart || this.normalize(user.FullName) === normalized;
            });// Find the user in the list of users
        }
        return [...keys].map(key => key.trim()).filter(Boolean);
    }// Build all manager lookup keys from login email/UPN/name (includes email local part, e.g. alehandro).

    async ensureSubordinatesForManager(identifiers: string[], forceRefresh = false): Promise<DomainUser[]> {
        if (forceRefresh) {
            this.clearSubordinatesCache();
            this.cachedUsers = null;
            this.usersLoadPromise = null;
        }
        if (this.cachedSubordinates !== null) {
            return this.cachedSubordinates;
        }
        if (this.subordinatesLoadPromise) {
            return this.subordinatesLoadPromise;
        }

        this.subordinatesLoadPromise = (async () => {
            const searchKeys = await this.buildManagerSearchKeys(identifiers);
            if (!searchKeys.length) {
                this.cachedSubordinates = [];
                return this.cachedSubordinates;
            }

            const results = await Promise.all(
                searchKeys
                    .filter(key => this.isStrongManagerSearchKey(key))
                    .map(key => this.getAllSubordinates(key).catch(() => []))
            );
            const byKey = new Map<string, DomainUser>();
            for (const employee of results.flat()) {
                const key = (employee.ADmail || employee.FullName).toLowerCase();
                if (key && !byKey.has(key)) {
                    byKey.set(key, employee);
                }
            }
            this.cachedSubordinates = [...byKey.values()];
            return this.cachedSubordinates;
        })();

        try {
            return await this.subordinatesLoadPromise;
        } finally {
            this.subordinatesLoadPromise = null;
        }
    }

    isTaskAssignedToSubordinate(assignedTo: string | undefined | null): boolean {
        return this.isTaskAssignedToAnyAssigneeValue(
            assignedTo ? [assignedTo] : []
        );
    }

    isTaskAssignedToEmployeeForDomainUser(assignedTo: string | undefined | null, employee: DomainUser): boolean {
        return this.isTaskAssignedToEmployee(assignedTo, employee);
    }

    isTaskAssignedToAnyAssigneeValue(assigneeValues: string[]): boolean {
        if (!assigneeValues.length || !this.cachedSubordinates?.length) {
            return false;
        }
        return assigneeValues.some(value =>
            this.cachedSubordinates!.some(employee => this.isTaskAssignedToEmployee(value, employee))
        );
    }


    getDirectReports(managerEmail: string): Observable<DomainUser[]> {
        return from(this.getDirectReportsFromGraph(managerEmail)).pipe(
            catchError(err => {
                console.error('SubordinateTaskService: getDirectReports failed', err);
                return of([]);
            })
        );
    }
   
      
      toggleEmployeeFolder(key: string): void {
        this.expandedEmployeeFolders.has(key)
          ? this.expandedEmployeeFolders.delete(key)
          : this.expandedEmployeeFolders.add(key);
      }
      
      isEmployeeFolderExpanded(key: string): boolean {
        return this.expandedEmployeeFolders.has(key);
      }

    // This method is used to get the tasks for a subordinate
    async getAllSubordinateEmails(managerEmail: string, visited: Set<string> = new Set()): Promise<string[]> {
        const visitKey = this.normalize(managerEmail);
        if (!visitKey || visited.has(visitKey)) return [];
        visited.add(visitKey);

        const directReports = await this.getDirectReports(managerEmail).toPromise();
        if (!directReports || directReports.length === 0) return [];

        const emails = directReports.map(u => u.ADmail);

        const nested = await Promise.all(
            emails.map(email => this.getAllSubordinateEmails(email, visited))
        );

        return [...emails, ...nested.flat()];
    }

    async getAllSubordinates(managerIdentifier: string, visited: Set<string> = new Set()): Promise<DomainUser[]> {
        const visitKey = this.normalize(managerIdentifier);
        if (!visitKey || visited.has(visitKey)) return [];
        visited.add(visitKey);

        const directReports = await this.getDirectReports(managerIdentifier).toPromise() ?? [];
        if (!directReports.length) return [];
    
        const nested = await Promise.all(
          directReports.map(u => this.getAllSubordinates(u.ADmail || u.FullName || u.PinNo, visited))
        );

        return [...directReports, ...nested.flat()];
    }

    async isSubordinate(managerEmail: string, targetEmail: string): Promise<boolean> {
        const allEmails = await this.getAllSubordinateEmails(managerEmail);
        return allEmails.includes(targetEmail.toLowerCase());
      }

      
    private async getDirectReportsFromGraph(managerIdentifier: string): Promise<DomainUser[]> {
        const manager = this.normalize(managerIdentifier);
        if (!manager) return [];

        const managerLocal = manager.includes('@') ? manager.split('@')[0] : manager;
        const users = await this.loadDomainUsers();
        return users.filter(user => this.managerMatchesIdentifier(managerIdentifier, user));
    }

    private managerMatchesIdentifier(managerIdentifier: string, user: DomainUser): boolean {
        const manager = this.normalize(managerIdentifier);
        if (!manager) return false;

        const managerLocal = manager.includes('@') ? manager.split('@')[0] : manager;
        const managerField = this.normalize(user.Manager);
        const managerEmail = this.normalize(user.ManagerEmail);
        const managerEmailLocal = managerEmail.includes('@') ? managerEmail.split('@')[0] : managerEmail;

        if (managerEmail === manager || managerEmailLocal === manager || managerEmailLocal === managerLocal) {
            return true;
        }

        if (managerField === manager) return true;
        if (managerLocal.length >= 4 && managerField === managerLocal) return true;

        // Avoid substring includes on short keys — they match almost everyone.
        if (manager.length >= 4 && (managerField.includes(manager) || managerField.includes(managerLocal))) {
            return true;
        }

        return false;
    }

    /** Keys safe for org-tree walks: emails, full names, pins — not short name fragments. */
    private isStrongManagerSearchKey(key: string): boolean {
        const trimmed = (key ?? '').trim();
        if (!trimmed) return false;
        if (trimmed.includes('@')) return true;
        if (/^\d+$/.test(trimmed)) return true;
        return trimmed.length >= 6;
    }

    private async loadDomainUsers(): Promise<DomainUser[]> {
        if (this.cachedUsers) return this.cachedUsers;
        if (this.usersLoadPromise) return this.usersLoadPromise;

        this.usersLoadPromise = (async () => {
            const token = await this.authService.acquireSharePointToken();
            const siteId = await this.resolveSiteId(token);
            const listId = await this.resolveListId(siteId, sharePointConfig.allDomainUsersListDisplayName, token);
            const columnMap = await this.resolveDomainUserColumnMap(siteId, listId, token);
            const items = await this.fetchAllListItems(siteId, listId, token);
            this.cachedUsers = items.map(item => this.mapDomainUser(item.fields ?? {}, columnMap));
            return this.cachedUsers;
        })();

        try {
            return await this.usersLoadPromise;
        } finally {
            this.usersLoadPromise = null;
        }
    }

    private async resolveSiteId(token: string): Promise<string> {
        const { siteId } = await this.siteMetadata.resolve(token);
        if (!siteId) throw new Error('Could not resolve SharePoint site ID');
        return siteId;
    }

    private async resolveListId(_siteId: string, listName: string, token: string): Promise<string> {
        const targetList = await this.siteMetadata.findList(token, listName);
        if (!targetList?.id) throw new Error(`List "${listName}" not found on site`);
        return targetList.id;
    }

    private async fetchAllListItems(siteId: string, listId: string, token: string): Promise<any[]> {
        const collected: any[] = [];
        let nextPath: string | null =
            `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;

        while (nextPath) {
            const page: any = await firstValueFrom(
                graphGet(this.http, nextPath, token, AppConstants.graphFileListingTimeoutMs)
            );
            collected.push(...(page?.value ?? []));
            nextPath = toGraphPath(page?.['@odata.nextLink']);
        }

        return collected;
    }

    private async resolveDomainUserColumnMap(siteId: string, listId: string, token: string): Promise<DomainUserColumnMap> {
        const columnsResp: any = await firstValueFrom(
            graphGet(
                this.http,
                `/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`,
                token
            )
        );
        const columns: Array<{ name?: string; displayName?: string }> = columnsResp?.value ?? [];

        return {
            PinNo: this.resolveColumnNames(columns, ['PinNo', 'PIN']),
            FullName: this.resolveColumnNames(columns, ['FullName', 'DisplayName']),
            ADmail: this.resolveColumnNames(columns, ['ADmail', 'Email', 'EMail']),
            Manager: this.resolveColumnNames(columns, ['Manager']),
            ManagerEmail: this.resolveColumnNames(columns, ['ManagerEmail', 'Manager Email']),
        };
    }

    private resolveColumnNames(
        columns: Array<{ name?: string; displayName?: string }>,
        displayNames: string[]
    ): string[] {
        const normalizedTargets = displayNames.map(name => this.normalizeFieldKey(name));
        const resolved = columns
            .filter(column => {
                const displayName = this.normalizeFieldKey(column.displayName ?? '');
                const name = this.normalizeFieldKey(column.name ?? '');
                return normalizedTargets.some(target => displayName === target || name === target);
            })
            .flatMap(column => [column.name, column.displayName])
            .filter((name): name is string => !!name);

        return [...new Set([...resolved, ...displayNames])];
    }

    private getFullNameValue(fields: Record<string, any>, columnMap: DomainUserColumnMap): string {
        // Read FullName directly — SharePoint Title holds dept codes (e.g. "EPD"), not person names.
        if (fields['FullName'] != null && fields['FullName'] !== '') {
            return this.stringifyFieldValue(fields['FullName']);
        }

        const candidates = columnMap.FullName.filter(
            name => this.normalizeFieldKey(name) !== 'title'
        );
        return this.getFieldValue(fields, candidates.length ? candidates : ['DisplayName']);
    }

    private mapDomainUser(fields: Record<string, any>, columnMap: DomainUserColumnMap): DomainUser {
        return {
            PinNo: this.getFieldValue(fields, columnMap.PinNo),
            FullName: this.getFullNameValue(fields, columnMap),
            ADmail: this.getFieldValue(fields, columnMap.ADmail),
            Manager: this.getFieldValue(fields, columnMap.Manager),
            ManagerEmail: this.getFieldValue(fields, columnMap.ManagerEmail),
        };
    }

    private getFieldValue(fields: Record<string, any>, candidates: string[]): string {
        for (const candidate of candidates) {
            const value = fields[candidate];
            if (value != null && value !== '') return this.stringifyFieldValue(value);
        }

        const candidateKeys = candidates.map(candidate => this.normalizeFieldKey(candidate));
        const matchingKey = Object.keys(fields).find(key => {
            const normalizedKey = this.normalizeFieldKey(key);
            return candidateKeys.some(candidateKey =>
                normalizedKey === candidateKey ||
                normalizedKey.startsWith(candidateKey)
            ) && !normalizedKey.includes('lookupid');
        });

        return matchingKey ? this.stringifyFieldValue(fields[matchingKey]) : '';
    }

    private stringifyFieldValue(value: unknown): string {
        if (value == null) return '';
        if (Array.isArray(value)) {
            return value.map(item => this.stringifyFieldValue(item)).filter(Boolean).join(', ');
        }
        if (typeof value === 'object') {
            const record = value as Record<string, unknown>;
            return String(
                record['LookupValue'] ??
                record['displayName'] ??
                record['email'] ??
                record['value'] ??
                ''
            );
        }
        return String(value);
    }

    private normalizeFieldKey(value: string): string {
        return value.replace(/[^a-z0-9]/gi, '').toLowerCase();
    }

    private normalize(value: string | undefined | null): string {
        return String(value ?? '').trim().toLowerCase();
    }

    getTasksForEmployee(employee: DomainUser): ToDoTask[] {
        return this.tasks.filter((task: ToDoTask & { eFormDetails?: { assignedTo?: string; status?: string }; assignedTo?: string }) => {
            const status = String(task.eFormDetails?.status ?? task.status ?? '').toLowerCase();
            if (status.includes('approv') || status.includes('complet') || task.status === 'completed') {
                return false;
            }
            if (task.status !== 'pending' && task.status !== 'rejected') {
                return false;
            }
            const assignedTo = task.eFormDetails?.assignedTo ?? task.assignedTo;
            return this.isTaskAssignedToEmployee(assignedTo, employee);
        });
    }

    getEmployeeDisplayName(employee: DomainUser): string {
        return employee.FullName || employee.ADmail || 'Unknown User';
    }

    /**
     * Resolve an HRPersonal folder name (e.g. "1004 Jason Saliba" or
     * "1004 jason.saliba@…") to an AllDomainUsers row so HR Files can use
     * SharePoint LookupId filtering even when the folder has no email.
     */
    async findUserForHrFolder(folderName: string): Promise<DomainUser | null> {
        const raw = String(folderName ?? '').trim();
        if (!raw) return null;

        let users: DomainUser[];
        try {
            users = await this.loadDomainUsers();
        } catch {
            return null;
        }
        if (!users.length) return null;

        const email =
            raw.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i)?.[0]?.toLowerCase() ?? '';
        const pin = raw.match(/^\d+/)?.[0] ?? '';
        const withoutPin = raw
            .replace(/^\d+\s+/, '')
            .replace(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i, '')
            .trim();
        const nameNorm = this.normalizePersonKey(withoutPin);

        if (email) {
            const byEmail = users.find(u => this.normalize(u.ADmail) === email);
            if (byEmail) return byEmail;
        }

        if (pin) {
            const pinNorm = pin.replace(/^0+/, '') || pin;
            const byPin = users.filter(u => {
                const userPin = String(u.PinNo ?? '').trim();
                if (!userPin) return false;
                if (userPin === pin) return true;
                return userPin.replace(/^0+/, '') === pinNorm;
            });
            if (byPin.length === 1) return byPin[0];
            if (byPin.length > 1 && nameNorm) {
                const named = byPin.find(u => this.namesLooselyMatch(u.FullName, nameNorm));
                if (named) return named;
                // Prefer any pin hit whose surname/token overlaps the folder name.
                const loose = byPin.find(u => {
                    const person = this.normalizePersonKey(u.FullName);
                    return nameNorm.split(/\s+/).some(part => part.length >= 4 && person.includes(part));
                });
                if (loose) return loose;
            }
            if (byPin.length >= 1 && nameNorm) {
                // Still return best pin row rather than giving up — email unlocks LookupId.
                return byPin[0];
            }
            if (byPin.length > 0 && !nameNorm) return byPin[0];
        }

        if (nameNorm) {
            const byName = users.filter(u => this.namesLooselyMatch(u.FullName, nameNorm));
            if (byName.length === 1) return byName[0];
            if (byName.length > 1) {
                // Prefer the row that has an email address.
                const withEmail = byName.find(u => String(u.ADmail ?? '').includes('@'));
                if (withEmail) return withEmail;
                return byName[0];
            }
        }

        return null;
    }

    private namesLooselyMatch(fullName: string, folderNameNorm: string): boolean {
        const person = this.normalizePersonKey(fullName);
        if (!person || !folderNameNorm) return false;
        if (person === folderNameNorm) return true;
        if (person.includes(folderNameNorm) || folderNameNorm.includes(person)) return true;

        const personParts = person.split(/\s+/).filter(p => p.length >= 3);
        const folderParts = folderNameNorm.split(/\s+/).filter(p => p.length >= 3);
        if (personParts.length === 0 || folderParts.length === 0) return false;
        const hits = folderParts.filter(fp => personParts.some(pp => pp === fp || pp.includes(fp) || fp.includes(pp)));
        return hits.length >= Math.min(2, folderParts.length);
    }

    private normalizePersonKey(value: string): string {
        return String(value ?? '')
            .toLowerCase()
            .replace(/[^a-z0-9\s]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    private taskMatchesSearch(task: ToDoTask, query: string): boolean {
        const extended = task as ToDoTask & {
            listName?: string;
            submittedDate?: string;
            eFormDetails?: Record<string, unknown>;
        };
        const searchableValues = [
            extended.name,
            extended.description,
            extended.status,
            extended.assignedTo,
            extended.listName,
            extended.submittedBy,
            extended.submittedDate,
            ...Object.values(extended.eFormDetails ?? {}),
        ];
        return searchableValues.some(value => String(value ?? '').toLowerCase().includes(query));
    }

    private isTaskAssignedToEmployee(assignedTo: string | undefined | null, employee: DomainUser): boolean {
        if (!assignedTo?.trim()) return false;

        const assignees = this.userService.parseAssignees(assignedTo);
        const candidates = assignees.length ? assignees : [assignedTo.trim()];

        return candidates.some(assignee =>
            this.userService.matchesAssigneeField(assignee, employee.FullName, employee.ADmail) ||
            this.userService.matchesAssigneeField(assignee, employee.ADmail, employee.ADmail) ||
            (employee.PinNo
                ? this.userService.matchesAssigneeField(assignee, employee.PinNo, employee.ADmail)
                : false)
        );
    }

    private employeeMatchTokens(employee: DomainUser): string[] {
        const tokens = new Set<string>();
        const add = (value: string | undefined | null) => {
            const normalized = this.normalizeMatchText(value ?? '');
            if (normalized.length >= 4) tokens.add(normalized);
        };

        add(employee.FullName);
        add(employee.ADmail);
        add(employee.PinNo);
        if (employee.ADmail.includes('@')) {
            add(employee.ADmail.split('@')[0]);
        }
        employee.FullName.split(/\s+/).filter(part => part.length > 2).forEach(part => add(part));

        return [...tokens];
    }

    private normalizeMatchText(value: string): string {
        return value.toLowerCase().replace(/[^a-z0-9]/g, '');
    }

}
