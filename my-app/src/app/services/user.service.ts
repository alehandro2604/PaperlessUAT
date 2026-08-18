import { Injectable } from '@angular/core';

export interface CurrentUser {
  email: string;
  username: string;
  userPrincipalName: string;
  employeeId?: string;
}

@Injectable({
  providedIn: 'root'
})
export class UserService {
  geteformId(): string {
    return '';
  }
  private currentUser: CurrentUser | null = null;

  setCurrentUser(user: CurrentUser | null): void {
    this.currentUser = user;
  }

  getCurrentUser(): CurrentUser | null {
    return this.currentUser;
  }

  getCurrentUserName(): string {
    return this.currentUser?.username || this.currentUser?.userPrincipalName || 'Unknown User';
  }

  getCurrentUserEmail(): string {
    return this.currentUser?.email || '';
  }

  parseAssignees(assignedTo: string | undefined | null): string[] {
    if (!assignedTo?.trim()) return [];
    return assignedTo.split(/[,;]+/).map(a => a.trim()).filter(Boolean);
  }

  /** True when the current user appears in an Assigned To value (supports multiple comma/semicolon-separated names). */
  matchesAssigneeField(assignedTo: string | undefined | null, name?: string, email?: string): boolean {
    if (!assignedTo?.trim()) return false;

    const currentUserName = (name ?? this.getCurrentUserName()).toLowerCase().trim();
    const currentUserEmail = (email ?? this.getCurrentUserEmail()).toLowerCase().trim();
    const assignees = this.parseAssignees(assignedTo);

    return assignees.some(assignee =>
      this.matchesSingleAssignee(assignee, currentUserName, currentUserEmail)
    );
  }

  private matchesSingleAssignee(assignedTo: string, name: string, email: string): boolean {
    const assignedToLower = assignedTo.toLowerCase().trim();
    if (!assignedToLower) return false;
    if (assignedToLower === name) return true;
    if (email && assignedToLower.includes(email)) return true;

    const emailLocal = email.includes('@') ? email.split('@')[0].toLowerCase() : email.toLowerCase();
    if (emailLocal.length > 2 && assignedToLower.includes(emailLocal)) return true;

    const normalizedAssignee = assignedToLower.replace(/[^a-z0-9]/g, '');
    const normalizedEmailLocal = emailLocal.replace(/[^a-z0-9]/g, '');
    if (normalizedEmailLocal.length > 3 && normalizedAssignee.includes(normalizedEmailLocal)) return true;

    if (this.namesReferToSamePerson(assignedToLower, name)) return true;

    const nameParts = name.split(/[\s,-]+/).filter(p => p.length > 2);
    if (nameParts.length === 0) return false;

    const matchedParts = nameParts.filter(part => assignedToLower.includes(part));
    return matchedParts.length >= Math.min(2, nameParts.length);
  }

  /** Treat "Last, First" and "First Last" as the same person when all parts match. */
  private namesReferToSamePerson(assignee: string, displayName: string): boolean {
    const normalize = (value: string): string[] =>
      value
        .toLowerCase()
        .split(/[\s,;]+/)
        .map(part => part.trim())
        .filter(part => part.length > 1);

    const assigneeParts = normalize(assignee);
    const nameParts = normalize(displayName);
    if (assigneeParts.length === 0 || nameParts.length === 0) return false;
    if (assigneeParts.length !== nameParts.length) return false;

    const assigneeSet = new Set(assigneeParts);
    return nameParts.every(part => assigneeSet.has(part));
  }
}