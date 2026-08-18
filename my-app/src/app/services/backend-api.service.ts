// Calls the Node backend (paperless-backend) which caches Graph results in Redis.
// The backend uses the on-behalf-of flow, so the token we send must be issued for
// the backend API itself (backendApiScope), NOT a Graph token.
import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { AuthService } from './auth.service';
import { sharePointConfig } from '../sharepoint.config';

@Injectable({ providedIn: 'root' })
export class BackendApiService {
  constructor(private http: HttpClient, private auth: AuthService) {}

  /**
   * Cached per-person task lookup. First call for a person hits Graph
   * (slower); repeat calls within the cache TTL are served from Redis.
   * @param list SharePoint list name/id; omit to use the backend's default.
   */
  async getTasksForPerson(personEmail: string, list?: string): Promise<any> {
    const token = await this.auth.acquireTokenSilent([sharePointConfig.backendApiScope]);
    const url = `${sharePointConfig.backendUrl}/api/tasks/${encodeURIComponent(personEmail)}`;
    return firstValueFrom(
      this.http.get(url, {
        headers: { Authorization: `Bearer ${token}` },
        params: list ? { list } : {},
      })
    );
  }

  /** Drop the Redis read-through entry for a person/list so the next GET refetches Graph. */
  async invalidateTasksForPerson(personEmail: string, list?: string): Promise<void> {
    const token = await this.auth.acquireTokenSilent([sharePointConfig.backendApiScope]);
    const url = `${sharePointConfig.backendUrl}/api/tasks/${encodeURIComponent(personEmail)}`;
    await firstValueFrom(
      this.http.delete(url, {
        headers: { Authorization: `Bearer ${token}` },
        params: list ? { list } : {},
      })
    );
  }
}
