import { Injectable } from '@angular/core';
import {PublicClientApplication,InteractionRequiredAuthError,AuthenticationResult,AccountInfo} from '@azure/msal-browser';
import { sharePointConfig } from '../sharepoint.config';
import { AppConstants } from '../app.constants';
import { withTimeout } from '../microsoft-graph';

export interface UserProfile {
  email: string;
  username: string;
  userPrincipalName: string;
  employeeId?: string;
}

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  getSharePointListUrl() {
      throw new Error('Method not implemented.');
  }
  getSharePointSiteUrl() {
      throw new Error('Method not implemented.');
  }
  getSharePointToken() {
      throw new Error('Method not implemented.');
  }
  getAccessToken() {
    throw new Error('Method not implemented.');
  }
  private readonly graphScopes = ['User.Read'];
  private readonly sharePointReadScopes = ['Sites.Selected'];
  //to write data into sharepoint directly(might not work due to permissions not granted yet)
  private readonly sharePointWriteScopes = ['Sites.Selected'];
  // SharePoint REST (_api) needs a resource-specific token, not a Graph token.
  // The ".default" scope on the tenant host requests the app's granted SP perms.
  private readonly sharePointRestWriteScopes = [
    `https://${sharePointConfig.siteHostName}/.default`,
  ];
  private initialized = false;
  private hasShownNotLoggedInAlert = false;

  private readonly msal = new PublicClientApplication({
    auth: {
      clientId: sharePointConfig.clientId,
      authority: `https://login.microsoftonline.com/${sharePointConfig.tenantId}`,
      redirectUri: sharePointConfig.redirectUri,
    },
  });

  constructor() {}

  async initialize(): Promise<void> {
    if (this.initialized) return;
    await this.msal.initialize();
    this.initialized = true;
  }

  async handleRedirect(): Promise<AuthenticationResult | null> {
    await this.initialize();
    return await this.msal.handleRedirectPromise();
  }

  getActiveAccount(): AccountInfo | null {
    return this.msal.getActiveAccount();
  }

  setActiveAccount(account: AccountInfo): void {
    this.msal.setActiveAccount(account);
  }

  getAllAccounts(): AccountInfo[] {
    return this.msal.getAllAccounts();
  }

  async acquireTokenSilent(scopes: string[]): Promise<string> {
    const account = this.getActiveAccount();
    if (!account) {
      throw new Error('No signed-in account found. Please sign in again.');
    }

    const result = await withTimeout(
      this.msal.acquireTokenSilent({ account, scopes }),
      15_000,
      'Timed out while requesting access token.'
    );
    return result.accessToken;
  }

  async acquireGraphToken(): Promise<string> {
    return this.acquireTokenSilent(this.graphScopes);
  }

  async acquireSharePointToken(): Promise<string> {
    const account = this.getActiveAccount();
    if (!account) {
      throw new Error('No signed-in account found. Please sign in again.');
    }

    try {
      return await this.acquireTokenSilent(this.sharePointReadScopes);
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const popupResult = await withTimeout(
          this.msal.acquireTokenPopup({ account, scopes: this.sharePointReadScopes }),
          30_000,
          'Timed out while waiting for SharePoint permission consent.'
        );
        return popupResult.accessToken;
      }
      throw new Error(
        'SharePoint read permission is missing. Grant delegated "Sites.Read.All" consent, then sign in again.'
      );
    }
  }
  //newly added code
  async acquireSharePointWriteToken(): Promise<string> {
    const account = this.getActiveAccount();
    if (!account) {
      throw new Error('No signed-in account found. Please sign in again.');
    }

    try {
      return await this.acquireTokenSilent(this.sharePointWriteScopes);
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const popupResult = await withTimeout(
          this.msal.acquireTokenPopup({ account, scopes: this.sharePointWriteScopes }),
          30_000,
          'Timed out while waiting for SharePoint write permission consent.'
        );
        return popupResult.accessToken;
      }
      throw new Error(
        'SharePoint write permission is missing. Grant delegated "Sites.Selected" consent and ensure site access is configured, then sign in again.'
      );
    }
  }

  /** Token for SharePoint REST (_api) calls such as ensureuser and list item MERGE. */
  async acquireSharePointRestWriteToken(): Promise<string> {
    const account = this.getActiveAccount();
    if (!account) {
      throw new Error('No signed-in account found. Please sign in again.');
    }

    try {
      return await this.acquireTokenSilent(this.sharePointRestWriteScopes);
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const popupResult = await withTimeout(
          this.msal.acquireTokenPopup({ account, scopes: this.sharePointRestWriteScopes }),
          30_000,
          'Timed out while waiting for SharePoint REST permission consent.',
        );
        return popupResult.accessToken;
      }
      throw new Error(
        'SharePoint REST write permission is missing. Grant delegated SharePoint access, then sign in again.',
      );
    }
  }
  //end of newly added code

  async loginRedirect(prompt: string): Promise<void> {
    const currentPath = `${window.location.pathname}${window.location.search}${window.location.hash}`;
    sessionStorage.setItem(AppConstants.postLoginPathKey, currentPath);
    await this.msal.loginRedirect({
      scopes: this.graphScopes,
      prompt: 'select_account',
      redirectStartPage: window.location.href,
    });
  }

  logoutRedirect(): void {
    const account = this.getActiveAccount();
    this.msal.logoutRedirect({ account: account ?? undefined });
  }

  restorePostLoginPathIfNeeded(): void {
    const targetPath = sessionStorage.getItem(AppConstants.postLoginPathKey);
    if (!targetPath) return;
    sessionStorage.removeItem(AppConstants.postLoginPathKey);
    const currentPath = `${window.location.pathname}${window.location.search}${window.location.hash}`;
    if (targetPath !== currentPath) {
      window.location.replace(targetPath);
    }
  }

  shouldShowNotLoggedInAlert(): boolean {
    if (!this.hasShownNotLoggedInAlert) {
      this.hasShownNotLoggedInAlert = true;
      return true;
    }
    return false;
  }

  resetNotLoggedInAlert(): void {
    this.hasShownNotLoggedInAlert = false;
  }
}
