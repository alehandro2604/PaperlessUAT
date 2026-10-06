import { TestBed } from '@angular/core/testing';
import { provideHttpClient } from '@angular/common/http';
import { HttpTestingController, provideHttpClientTesting } from '@angular/common/http/testing';

import { CacheService } from './cache';
import { AuthService } from './auth.service';
import { sharePointConfig } from '../sharepoint.config';

describe('CacheService', () => {
  let service: CacheService;
  let http: HttpTestingController;
  const base = `${sharePointConfig.backendUrl}/api/cache`;

  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [
        provideHttpClient(),
        provideHttpClientTesting(),
        { provide: AuthService, useValue: { acquireTokenSilent: async () => 'test-token' } },
      ],
    });
    service = TestBed.inject(CacheService);
    http = TestBed.inject(HttpTestingController);
  });

  afterEach(() => http.verify());

  it('sends the backend token and returns undefined for a missing key', async () => {
    const pending = service.get('fc:missing');
    await new Promise(r => setTimeout(r));
    const req = http.expectOne(r => r.url === `${base}/entry` && r.params.get('key') === 'fc:missing');
    expect(req.request.headers.get('Authorization')).toBe('Bearer test-token');
    req.flush(null);
    expect(await pending).toBeUndefined();
  });

  it('skips oversized entries instead of throwing', async () => {
    const pending = service.set('fc:big', { rows: [] });
    await new Promise(r => setTimeout(r));
    http.expectOne(`${base}/entry`).flush('too large', { status: 413, statusText: 'Payload Too Large' });
    await expect(pending).resolves.toBeUndefined();
  });
});
