import { isDevMode } from '@angular/core';
import { Capacitor } from '@capacitor/core';

// Single source of truth for environment config. There is deliberately no
// runtime config.json — everything reads from here.
//
// This repo is UAT only: http://localhost:4201, backend :3001, sites/PaperlessUAT.
//
// LOCAL vs DEPLOYED: picked automatically, nothing to flip. `ng serve` (dev mode) uses
// localhost and the local backend; a production build (`npm run build:uat`) uses the target's deployedUrl with a same-origin backend ('').
// When deployed the Angular app is HTTPS, so backendUrl MUST be HTTPS (or same-origin).
// Browsers block http://localhost from an https:// page (mixed content), and
// localhost is the *user's* PC — not the server. IIS/ARR (or nginx) on each server
// must proxy /api to that environment's Node backend.
const target = {
  label: 'Paperless System UAT',
  localPort: 4201,
  deployedUrl: 'https://swiftpaperlessuat.enemalta.lan',
  backendUrl: 'http://localhost:3001',
  sitePath: 'sites/PaperlessUAT',
  siteId: 'emoffice365.sharepoint.com,d07b0fc2-c9d8-4793-8f51-0afb2e7d0fb7,e85f6d4d-da1c-4d28-b615-d9f7f44656c4',
};

export const sharePointConfig = {
  target: 'uat' as const,
  label: target.label,

  tenantId: '331b05a3-d03f-421a-9573-f2db41268c2e',
  clientId: '45a7a07b-7d2b-4240-98cf-2847ae6245f9',

  // Android WebView (APK origin https://localhost) / local `ng serve` / deployed build.
  // Every URI here must be a Single-page application redirect in Entra.
  redirectUri: Capacitor.getPlatform() === 'android'
    ? 'https://localhost'
    : isDevMode()
      ? `http://localhost:${target.localPort}`
      : target.deployedUrl,
  backendUrl: Capacitor.getPlatform() === 'android' || isDevMode() ? target.backendUrl : '',

  siteHostName: 'emoffice365.sharepoint.com',
  sitePath: target.sitePath,
  siteId: target.siteId,
  documentLibrariesTasks: [] as string[],
  hrPersonalListDisplayName: 'HRPersonal',

  delegateListDisplayName: 'UsersWhoCanDelegate',
  allDomainUsersListDisplayName: 'AllDomainUsers',

  // Requires "Expose an API" scope on the app registration (see backend README).
  backendApiScope: 'api://45a7a07b-7d2b-4240-98cf-2847ae6245f9/access_as_user'
};
