/** The signed-in user, as loaded from Microsoft Graph /me. */
export interface AppUser {
  email: string;
  username: string;
  userPrincipalName: string;
  employeeId?: string;
}
