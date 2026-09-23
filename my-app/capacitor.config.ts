import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'mt.com.enemalta.paperless',
  appName: 'Paperless',
  webDir: 'dist/my-app/browser',
  server: {
    // Stay inside the app for Microsoft sign-in. Any other host opens Chrome,
    // and Chrome cannot return to the app at https://localhost.
    allowNavigation: [
      'login.microsoftonline.com',
      '*.login.microsoftonline.com',
      '*.microsoftonline.com',
      'login.microsoft.com',
      '*.microsoft.com',
      'login.windows.net',
      '*.msauth.net',
      '*.msftauth.net',
      'login.live.com',
      'account.live.com',
    ],
  },
};

export default config;
