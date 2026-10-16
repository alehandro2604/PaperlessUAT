// Ties the user's personal Redis entries to this browser tab (see /session/start and
// /session/end in paperless-backend/cacheRoutes.js). A refresh keeps the tab id (sessionStorage),
// so the cache survives it; closing the tab or browser lets the backend expire the entries.
import { Injectable } from '@angular/core';
import { AuthService } from './auth.service';
import { sharePointConfig } from '../sharepoint.config';

const TAB_ID_KEY = 'paperless.cacheTabId';

@Injectable({ providedIn: 'root' })
export class CacheSessionService {
  private readonly base = `${sharePointConfig.backendUrl}/api/cache`;
  private readonly tabId = this.loadTabId();
  private started: Promise<void> | null = null;
  /** Last backend token seen; pagehide cannot wait for a fresh one. */
  private lastToken = '';

  constructor(private readonly auth: AuthService) {
    if (typeof window === 'undefined') return;
    // `persisted` = page is going into the back/forward cache and may come back as-is.
    window.addEventListener('pagehide', (event) => {
      if (!(event as PageTransitionEvent).persisted) this.end();
    });
    // A page restored from the back/forward cache never reloads, so re-announce it.
    window.addEventListener('pageshow', (event) => {
      if ((event as PageTransitionEvent).persisted) {
        this.started = null;
        void this.ensure();
      }
    });
  }

  noteToken(token: string): void {
    if (token) this.lastToken = token;
  }

  /** Tells the backend this tab is open. Once per page load; never throws. */
  ensure(): Promise<void> {
    this.started ??= this.start().catch((err) => {
      console.warn('[cache-session] start failed; entries fall back to the default expiry', err);
    });
    return this.started;
  }

  private async start(): Promise<void> {
    const token = await this.auth.acquireTokenSilent([sharePointConfig.backendApiScope]);
    this.noteToken(token);
    await this.post('session/start', token, false);
  }

  /** Best effort while the page is going away: keepalive lets the request outlive it. */
  private end(): void {
    if (!this.lastToken || !this.started) return;
    void this.post('session/end', this.lastToken, true).catch(() => undefined);
  }

  private async post(path: string, token: string, keepalive: boolean): Promise<void> {
    const res = await fetch(`${this.base}/${path}`, {
      method: 'POST',
      keepalive,
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabId: this.tabId }),
    });
    if (!res.ok) throw new Error(`${path} responded ${res.status}`);
  }

  private loadTabId(): string {
    try {
      let id = sessionStorage.getItem(TAB_ID_KEY);
      if (!id) {
        id = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
        sessionStorage.setItem(TAB_ID_KEY, id);
      }
      return id;
    } catch {
      return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    }
  }
}
