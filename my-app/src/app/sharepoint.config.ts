import { Capacitor } from '@capacitor/core';
import { sharePointTarget } from './sharepoint.target';

// Single source of truth for environment config. There is deliberately no
// runtime config.json — everything reads from here.
//
// LIVE vs UAT: don't edit this file to switch. Pick the target when you start
// or build (see sharepoint.target.ts). Both can run side by side locally:
//   npm run start:live  → http://localhost:4200, backend :3000, sites/PaperlessLive
//   npm run start:uat   → http://localhost:4201, backend :3001, sites/PaperlessUAT
//
// LOCAL vs DEPLOYED: flip the redirectUri / backendUrl pair in the target below.
// When deployed the Angular app is HTTPS, so backendUrl MUST be HTTPS (or same-origin).
// Browsers block http://localhost from an https:// page (mixed content), and
// localhost is the *user's* PC — not the server. Prefer reverse-proxying /api
// to Node on the same host so backendUrl can be '' (same origin).
const targets = {
  live: {
    label: 'Paperless System',
    localPort: 4200,
    backendUrl: 'http://localhost:3000',
    sitePath: 'sites/PaperlessLive',
    siteId: 'emoffice365.sharepoint.com,6f7b5d38-686a-40c8-8f83-5eb68880ab9b,20bef680-9015-46b5-9bfa-6d00b61c7fa9',
  },
  uat: {
    label: 'Paperless System UAT',
    localPort: 4201,
    backendUrl: 'http://localhost:3001',
    // --- Deployed UAT (https://swiftpaperlessuat.enemalta.lan) ---
    // Same-origin: IIS/ARR (or nginx) must proxy /api → http://127.0.0.1:3000
    // redirectUri: 'https://swiftpaperlessuat.enemalta.lan',
    // backendUrl: '',
    sitePath: 'sites/PaperlessUAT',
    siteId: 'emoffice365.sharepoint.com,d07b0fc2-c9d8-4793-8f51-0afb2e7d0fb7,e85f6d4d-da1c-4d28-b615-d9f7f44656c4',
  },
};

const target = targets[sharePointTarget];

export const sharePointConfig = {
  target: sharePointTarget,
  label: target.label,

  tenantId: '331b05a3-d03f-421a-9573-f2db41268c2e',
  clientId: '45a7a07b-7d2b-4240-98cf-2847ae6245f9',

  // --- LOCAL (browser) / Android WebView ---
  // The APK origin is https://localhost. Every URI here must be a Single-page application redirect in Entra.
  redirectUri: Capacitor.getPlatform() === 'android'
    ? 'https://localhost'
    : `http://localhost:${target.localPort}`,
  backendUrl: target.backendUrl,

  siteHostName: 'emoffice365.sharepoint.com',
  sitePath: target.sitePath,
  siteId: target.siteId,
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
