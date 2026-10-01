// Paperless desktop shell (Electron).
// Serves the built Angular site from http://localhost:4200 (the redirect URI already
// registered in Entra for local dev) and opens it in a window. Microsoft sign-in
// therefore works exactly as it does in the browser - no Entra changes needed.
const { app, BrowserWindow, shell, dialog } = require('electron');
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 4200;
const ORIGIN = `http://localhost:${PORT}`;

// Packaged: <install>/resources/app  |  Dev: ../dist/my-app/browser
const WEB_ROOT = app.isPackaged
  ? path.join(process.resourcesPath, 'app')
  : path.join(__dirname, '..', 'dist', 'my-app', 'browser');

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.webmanifest': 'application/manifest+json',
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.woff': 'font/woff', '.woff2': 'font/woff2',
  '.ttf': 'font/ttf', '.txt': 'text/plain; charset=utf-8', '.map': 'application/json',
};

// Hosts allowed to open inside the app (Microsoft sign-in / consent).
const MS_HOST = /(^|\.)(microsoftonline\.com|microsoft\.com|windows\.net|msauth\.net|msftauth\.net|live\.com)$/i;
const isMicrosoft = (u) => { try { return MS_HOST.test(new URL(u).hostname); } catch { return false; } };
const isLocal = (u) => { try { return new URL(u).origin === ORIGIN; } catch { return false; } };

function startServer() {
  const server = http.createServer((req, res) => {
    let rel = decodeURIComponent((req.url || '/').split('?')[0]);
    let file = path.normalize(path.join(WEB_ROOT, rel));
    if (!file.startsWith(WEB_ROOT)) { res.writeHead(403); return res.end(); }
    // SPA fallback: unknown paths (no extension) -> index.html
    if (!fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      file = path.extname(rel) ? null : path.join(WEB_ROOT, 'index.html');
    }
    if (!file || !fs.existsSync(file)) { res.writeHead(404); return res.end('Not found'); }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream' });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(PORT, '127.0.0.1', () => resolve(server));
  });
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1400, height: 900, minWidth: 900, minHeight: 600,
    title: 'Paperless', autoHideMenuBar: true,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true },
  });

  // window.open (MSAL popups for consent): keep Microsoft in-app, everything else -> real browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (isMicrosoft(url) || isLocal(url)) {
      return { action: 'allow', overrideBrowserWindowOptions: { width: 520, height: 720, autoHideMenuBar: true } };
    }
    shell.openExternal(url);
    return { action: 'deny' };
  });
  // Same for normal navigation (sign-in redirect stays in the window).
  win.webContents.on('will-navigate', (e, url) => {
    if (!isLocal(url) && !isMicrosoft(url)) { e.preventDefault(); shell.openExternal(url); }
  });

  win.loadURL(ORIGIN);
}

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    const w = BrowserWindow.getAllWindows()[0];
    if (w) { if (w.isMinimized()) w.restore(); w.focus(); }
  });

  app.whenReady().then(async () => {
    try {
      await startServer();
    } catch (err) {
      dialog.showErrorBox(
        'Paperless cannot start',
        err && err.code === 'EADDRINUSE'
          ? `Port ${PORT} is already in use. Close "ng serve" or any other program using it, then start Paperless again.`
          : String(err)
      );
      return app.quit();
    }
    if (!fs.existsSync(path.join(WEB_ROOT, 'index.html'))) {
      dialog.showErrorBox('Paperless cannot start', `Built site not found at:\n${WEB_ROOT}\n\nRun "ng build" first.`);
      return app.quit();
    }
    createWindow();
    app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
  });

  app.on('window-all-closed', () => app.quit());
}
