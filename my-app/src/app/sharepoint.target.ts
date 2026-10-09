// Which SharePoint site this build talks to. Do not edit by hand to switch:
// the "uat" build configuration swaps this file for sharepoint.target.uat.ts.
//   npm run start:live  → http://localhost:4200 (sites/PaperlessLive)
//   npm run start:uat   → http://localhost:4201 (sites/PaperlessUAT)
export const sharePointTarget: 'live' | 'uat' = 'live';
