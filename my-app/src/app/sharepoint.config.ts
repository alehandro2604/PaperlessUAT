import { Capacitor } from '@capacitor/core';

// Single source of truth for environment config. Target: LIVE (sites/PaperlessLive).
// There is deliberately no runtime config.json — everything reads from here.
//
// LOCAL vs UAT: flip the redirectUri / backendUrl pair together.
// On UAT the Angular app is HTTPS, so backendUrl MUST be HTTPS (or same-origin).
// Browsers block http://localhost from an https:// page (mixed content), and
// localhost is the *user's* PC — not the server. Prefer reverse-proxying /api
// to Node on the same host so backendUrl can be '' (same origin).
export const sharePointConfig = {
  tenantId: '331b05a3-d03f-421a-9573-f2db41268c2e',
  clientId: '45a7a07b-7d2b-4240-98cf-2847ae6245f9',

  // --- LOCAL (browser) / Android WebView ---
  // The APK origin is https://localhost. That URI must also be a Single-page application redirect in Entra.
  redirectUri: Capacitor.getPlatform() === 'android'
    ? 'https://localhost'
    : 'http://localhost:4200',
  backendUrl: 'http://localhost:3000',

  // --- UAT (https://swiftpaperlessuat.enemalta.lan) ---
  // Same-origin: IIS/ARR (or nginx) must proxy /api → http://127.0.0.1:3000
  // redirectUri: 'https://swiftpaperlessuat.enemalta.lan',
  // backendUrl: '',

  siteHostName: 'emoffice365.sharepoint.com',
  sitePath: 'sites/PaperlessLive',
  siteId:'emoffice365.sharepoint.com,6f7b5d38-686a-40c8-8f83-5eb68880ab9b,20bef680-9015-46b5-9bfa-6d00b61c7fa9',
  documentLibrariesTasks: [] as string[],
  hrPersonalListDisplayName: 'HRPersonal',
  // Users listed here are allowed by the app UI to browse all top-level folders
  // in HRPersonal. Users not listed here only see their own personal folder.
  //no hard coded users here
  hrPersonalAllAccessUsers: [] as string[],

  delegateListDisplayName: 'UsersWhoCanDelegate',
  allDomainUsersListDisplayName: 'AllDomainUsers',

  // Requires "Expose an API" scope on the app registration (see backend README).
  backendApiScope: 'api://45a7a07b-7d2b-4240-98cf-2847ae6245f9/access_as_user'
};