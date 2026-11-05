/**
 * `SaasService` over the real `ApiService`, against a fake transport answering the JSON
 * the `api/saas` controller answers.
 *
 * What matters here is invisible from the outside: that a failure body's `code` and
 * `correlationId` survive into the `ApiError`, that concurrent gates share one request,
 * and that a failure is never cached - a cached failure would keep a paying tenant's
 * features hidden until the cache expired.
 */

// `ApiService` and `SaasService` pull `@microsoft/sp-core-library` in for `Log`; plain
// jsdom has no SPFx runtime. Must precede the imports (no Babel hoisting).
const logged: { warn: string[]; error: string[] } = { warn: [], error: [] };
jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    verbose: (): void => undefined,
    info: (): void => undefined,
    warn: (_source: string, message: string): void => {
      logged.warn.push(message);
    },
    error: (_source: string, error: Error): void => {
      logged.error.push(error.message);
    }
  }
}));

import type { HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import { ApiError } from '../models/ApiError';
import { ISaasEntitlements } from '../models/Saas';
import { ApiService } from './ApiService';
import { ApiTokenStore } from './ApiTokenStore';
import { ISaasCacheStorage, SaasService } from './SaasService';

/** The shape `ApiService` needs from a transport; see `ProjectListingService.test.ts`. */
interface IFakeHttpClient {
  fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

interface IReply {
  status: number;
  body: unknown;
  /** Holds this one reply back until it resolves. */
  gate?: Promise<void>;
}

interface IRecordedRequest {
  url: string;
  method: string | undefined;
  headers: { [name: string]: string };
  body: unknown;
}

const BASE_URL: string = 'https://api.example.com/api';

const ENTITLEMENTS = {
  tenantKnown: true,
  entraTenantId: '11111111-2222-3333-4444-555555555555',
  isTenantAdmin: true,
  isBillingEnforced: true,
  services: ['TENDER'],
  features: ['TENDER.CORE'],
  details: [
    {
      serviceKey: 'TENDER',
      name: 'Tender',
      description: 'Tenders',
      status: 'Active',
      source: 'Marketplace',
      validTo: null,
      termEnd: '2026-12-04T00:00:00Z',
      isFreeTrial: false,
      autoRenew: true,
      canPurchase: false,
      marketplaceSubscriptionId: 'sub-1'
    }
  ],
  cacheSeconds: 60,
  evaluatedOn: '2026-11-04T12:00:00Z'
};

const ERROR_BODY = { code: 'SAAS_NOT_TENANT_ADMIN', message: 'Only tenant admins can buy.', correlationId: '0HN7:00000001' };

/** In-memory Web Storage. */
class MemoryStorage implements ISaasCacheStorage {
  public readonly items: { [key: string]: string } = {};
  // eslint-disable-next-line @rushstack/no-new-null -- Web Storage API signature
  public getItem(key: string): string | null {
    return Object.prototype.hasOwnProperty.call(this.items, key) ? this.items[key] : null;
  }
  public setItem(key: string, value: string): void {
    this.items[key] = value;
  }
  public removeItem(key: string): void {
    delete this.items[key];
  }
}

/** Storage that throws on every call, like a blocked or full sessionStorage. */
const BROKEN_STORAGE: ISaasCacheStorage = {
  getItem: (): never => {
    throw new Error('SecurityError');
  },
  setItem: (): never => {
    throw new Error('QuotaExceededError');
  },
  removeItem: (): never => {
    throw new Error('SecurityError');
  }
};

/**
 * A transport that answers each request with the next reply. A reply can be held back
 * (`gate`) to test requests that are in flight together.
 */
function fakeTransport(replies: IReply[], requests: IRecordedRequest[], gate?: Promise<void>): IFakeHttpClient {
  return {
    fetch: async (url: string, options: IHttpClientOptions): Promise<HttpClientResponse> => {
      requests.push({
        url,
        method: options.method,
        headers: options.headers as { [name: string]: string },
        body: options.body ? JSON.parse(String(options.body)) : undefined
      });
      const reply: IReply = replies.length > 1 ? (replies.shift() as IReply) : replies[0];
      if (gate) {
        await gate;
      }
      if (reply.gate) {
        await reply.gate;
      }
      return {
        ok: reply.status >= 200 && reply.status < 300,
        status: reply.status,
        statusText: reply.status === 200 ? 'OK' : 'Error',
        text: async (): Promise<string> => (reply.body === undefined ? '' : JSON.stringify(reply.body))
      } as unknown as HttpClientResponse;
    }
  };
}

interface IHarness {
  service: SaasService;
  requests: IRecordedRequest[];
  storage: MemoryStorage;
  clock: { now: number };
}

function build(
  replies: IReply[],
  // eslint-disable-next-line @rushstack/no-new-null -- `null` is SaasService's "no storage"
  options?: { gate?: Promise<void>; storage?: ISaasCacheStorage | null; scope?: string; token?: string }
): IHarness {
  const requests: IRecordedRequest[] = [];
  const storage: MemoryStorage = new MemoryStorage();
  const clock = { now: Date.parse('2026-11-04T12:00:00Z') };
  const tokens: ApiTokenStore = new ApiTokenStore();
  if (options?.token) {
    tokens.setToken(options.token);
  }

  const service: SaasService = new SaasService(
    new ApiService(fakeTransport(replies, requests, options?.gate), { baseUrl: BASE_URL, timeoutMs: 0 }, tokens),
    'saas',
    {
      cacheScope: options?.scope || 'tenant-a|user@contoso.com',
      storage: options && options.storage !== undefined ? options.storage : storage,
      now: () => clock.now
    }
  );

  return { service, requests, storage, clock };
}

beforeEach(() => {
  logged.warn.length = 0;
  logged.error.length = 0;
});

describe('SaasService.getEntitlements', () => {
  it('GETs saas/entitlements and returns the DTO, which is not enveloped', async () => {
    const { service, requests } = build([{ status: 200, body: ENTITLEMENTS }]);

    const result: ISaasEntitlements = await service.getEntitlements();

    expect(requests[0].url).toBe(`${BASE_URL}/saas/entitlements`);
    expect(requests[0].method).toBe('GET');
    expect(result.isTenantAdmin).toBe(true);
    expect(result.features).toEqual(['TENDER.CORE']);
    expect(result.details[0].termEnd).toBe('2026-12-04T00:00:00Z');
  });

  it('serves a second call from the cache until cacheSeconds have passed', async () => {
    const { service, requests, clock } = build([{ status: 200, body: ENTITLEMENTS }]);

    await service.getEntitlements();
    clock.now += 59 * 1000;
    await service.getEntitlements();
    expect(requests).toHaveLength(1);

    clock.now += 2 * 1000;
    await service.getEntitlements();
    expect(requests).toHaveLength(2);
  });

  it('defaults the cache lifetime to 300 seconds when the server omits it', async () => {
    const { service, requests, clock } = build([{ status: 200, body: { ...ENTITLEMENTS, cacheSeconds: undefined } }]);

    await service.getEntitlements();
    clock.now += 299 * 1000;
    await service.getEntitlements();
    expect(requests).toHaveLength(1);

    clock.now += 2 * 1000;
    await service.getEntitlements();
    expect(requests).toHaveLength(2);
  });

  it('caches in sessionStorage under the API base URL and tenant, so a new instance reuses it', async () => {
    const first = build([{ status: 200, body: ENTITLEMENTS }]);
    await first.service.getEntitlements();

    const keys: string[] = Object.keys(first.storage.items);
    expect(keys).toHaveLength(1);
    expect(keys[0]).toContain(BASE_URL);
    expect(keys[0]).toContain('tenant-a|user@contoso.com');

    // A page reload: a fresh service over the same storage.
    const second = build([{ status: 500, body: undefined }], { storage: first.storage });
    second.clock.now = first.clock.now;
    const result = await second.service.getEntitlements();

    expect(second.requests).toHaveLength(0);
    expect(result.features).toEqual(['TENDER.CORE']);
  });

  it("does not read another tenant's cached answer", async () => {
    const first = build([{ status: 200, body: ENTITLEMENTS }]);
    await first.service.getEntitlements();

    const other = build([{ status: 200, body: { ...ENTITLEMENTS, features: [] } }], {
      storage: first.storage,
      scope: 'tenant-b|user@fabrikam.com'
    });
    const result = await other.service.getEntitlements();

    expect(other.requests).toHaveLength(1);
    expect(result.features).toEqual([]);
  });

  it('never stores the bearer token', async () => {
    const { service, storage, requests } = build([{ status: 200, body: ENTITLEMENTS }], { token: 'secret-jwt' });

    await service.getEntitlements();

    expect(requests[0].headers.Authorization).toBe('Bearer secret-jwt');
    expect(JSON.stringify(storage.items)).not.toContain('secret-jwt');
  });

  it('shares one request between concurrent callers', async () => {
    let release: () => void = (): void => undefined;
    const gate: Promise<void> = new Promise<void>((resolve) => {
      release = resolve;
    });
    const { service, requests } = build([{ status: 200, body: ENTITLEMENTS }], { gate });

    const calls: Promise<ISaasEntitlements>[] = [service.getEntitlements(), service.getEntitlements(), service.getEntitlements()];
    release();
    const results: ISaasEntitlements[] = await Promise.all(calls);

    expect(requests).toHaveLength(1);
    expect(results[0]).toBe(results[1]);
    expect(results[1]).toBe(results[2]);
  });

  it('bypasses the cache on refresh and stores the fresh answer', async () => {
    const { service, requests } = build([
      { status: 200, body: { ...ENTITLEMENTS, features: [] } },
      { status: 200, body: ENTITLEMENTS }
    ]);

    expect((await service.getEntitlements()).features).toEqual([]);
    expect((await service.getEntitlements({ refresh: true })).features).toEqual(['TENDER.CORE']);
    expect((await service.getEntitlements()).features).toEqual(['TENDER.CORE']);
    expect(requests).toHaveLength(2);
  });

  it('does not cache a failure: the next call asks again', async () => {
    const { service, requests, storage } = build([
      { status: 503, body: { code: 'SAAS_UNAVAILABLE', message: 'Down', correlationId: 'c-503' } },
      { status: 200, body: ENTITLEMENTS }
    ]);

    await expect(service.getEntitlements()).rejects.toBeInstanceOf(ApiError);
    expect(Object.keys(storage.items)).toHaveLength(0);

    const result = await service.getEntitlements();
    expect(result.features).toEqual(['TENDER.CORE']);
    expect(requests).toHaveLength(2);
  });

  it('rejects a 2xx body that is not the DTO, and caches nothing', async () => {
    const { service, storage } = build([{ status: 200, body: 'maintenance page' }]);

    await expect(service.getEntitlements()).rejects.toThrow('unexpected format');
    expect(Object.keys(storage.items)).toHaveLength(0);
  });

  it('still works, from memory, when sessionStorage throws', async () => {
    const { service, requests } = build([{ status: 200, body: ENTITLEMENTS }], { storage: BROKEN_STORAGE });

    await service.getEntitlements();
    await service.getEntitlements();
    service.invalidate();

    expect(requests).toHaveLength(1);
  });

  it('asks again after invalidate', async () => {
    const { service, requests, storage } = build([{ status: 200, body: ENTITLEMENTS }]);

    await service.getEntitlements();
    service.invalidate();
    expect(Object.keys(storage.items)).toHaveLength(0);

    await service.getEntitlements();
    expect(requests).toHaveLength(2);
  });
});

describe('SaasService errors', () => {
  it('carries code and correlationId from the failure body into the ApiError', async () => {
    const { service } = build([{ status: 403, body: ERROR_BODY }]);

    const error: ApiError = await service.createPurchaseIntent('TENDER', 'https://contoso.sharepoint.com/#/subscription').then(
      () => {
        throw new Error('expected a rejection');
      },
      (caught: ApiError) => caught
    );

    expect(error).toBeInstanceOf(ApiError);
    expect(error.status).toBe(403);
    expect(error.code).toBe('SAAS_NOT_TENANT_ADMIN');
    expect(error.correlationId).toBe('0HN7:00000001');
    expect(error.message).toBe('Only tenant admins can buy.');
  });

  it('logs a 4xx as a warning and a 5xx as an error, with code and correlationId', async () => {
    const refused = build([{ status: 403, body: ERROR_BODY }]);
    await expect(refused.service.getCatalog()).rejects.toThrow();
    expect(logged.warn[0]).toContain('SAAS_NOT_TENANT_ADMIN');
    expect(logged.warn[0]).toContain('0HN7:00000001');

    const broken = build([{ status: 503, body: { code: 'SAAS_UNAVAILABLE', message: 'Down', correlationId: 'c-503' } }]);
    await expect(broken.service.getEntitlements()).rejects.toThrow();
    expect(logged.error.some((line: string) => line.indexOf('SAAS_UNAVAILABLE') >= 0 && line.indexOf('c-503') >= 0)).toBe(true);
  });

  it('never logs the bearer token', async () => {
    const { service } = build([{ status: 401, body: undefined }], { token: 'secret-jwt' });

    await expect(service.getEntitlements()).rejects.toThrow();

    expect(JSON.stringify(logged)).not.toContain('secret-jwt');
  });

  it('leaves code and correlationId empty for a body without them', async () => {
    const { service } = build([{ status: 401, body: undefined }]);

    const error: ApiError = await service.getEntitlements().catch((caught: ApiError) => caught) as ApiError;

    expect(error.status).toBe(401);
    expect(error.code).toBeUndefined();
    expect(error.correlationId).toBeUndefined();
  });
});

describe('SaasService other calls', () => {
  it('GETs the catalog', async () => {
    const { service, requests } = build([
      { status: 200, body: { services: [{ serviceKey: 'TENDER', name: 'Tender', displayOrder: 1, isAvailable: true, hasOffer: true, features: [], plans: [] }] } }
    ]);

    const catalog = await service.getCatalog();

    expect(requests[0].url).toBe(`${BASE_URL}/saas/catalog`);
    expect(catalog.services[0].serviceKey).toBe('TENDER');
  });

  it('GETs the subscriptions', async () => {
    const { service, requests } = build([{ status: 200, body: [{ marketplaceSubscriptionId: 'sub-1', status: 'Subscribed' }] }]);

    const subscriptions = await service.getSubscriptions();

    expect(requests[0].url).toBe(`${BASE_URL}/saas/subscriptions`);
    expect(subscriptions[0].marketplaceSubscriptionId).toBe('sub-1');
  });

  it('POSTs a purchase intent with the service and return URL', async () => {
    const { service, requests } = build([
      {
        status: 200,
        body: { intentId: 'i-1', marketplaceOfferUrl: 'https://appsource.microsoft.com/offer', expiresOn: '2026-11-05T12:00:00Z' }
      }
    ]);

    const intent = await service.createPurchaseIntent('TENDER', 'https://contoso.sharepoint.com/#/subscription?saas=returned&service=TENDER');

    expect(requests[0].url).toBe(`${BASE_URL}/saas/purchase-intents`);
    expect(requests[0].method).toBe('POST');
    expect(requests[0].body).toEqual({
      serviceKey: 'TENDER',
      returnUrl: 'https://contoso.sharepoint.com/#/subscription?saas=returned&service=TENDER'
    });
    expect(intent.marketplaceOfferUrl).toBe('https://appsource.microsoft.com/offer');
  });

  it('rejects a purchase intent without an offer URL', async () => {
    const { service } = build([{ status: 200, body: { intentId: 'i-1' } }]);

    await expect(service.createPurchaseIntent('TENDER', 'https://x')).rejects.toThrow('unexpected format');
  });

  it('normalises the subscriptions and rejects a body that is not a list', async () => {
    const listed = build([{ status: 200, body: [{ marketplaceSubscriptionId: 'sub-1' }, { status: 'Subscribed' }] }]);
    expect(await listed.service.getSubscriptions()).toHaveLength(1);

    const broken = build([{ status: 200, body: { items: [] } }]);
    await expect(broken.service.getSubscriptions()).rejects.toThrow('unexpected format');
  });
});

describe('SaasService.activateSubscription', () => {
  const SUBSCRIPTION_ID: string = '0f8fad5b-d9cb-469f-a165-70867728950e';
  const ACTIVATED = {
    state: 'Activated',
    subscription: {
      marketplaceSubscriptionId: SUBSCRIPTION_ID,
      subscriptionName: 'XSProject Tender',
      offerId: 'xs-tender',
      planId: 'monthly',
      planName: 'Monthly',
      quantity: null,
      status: 'Subscribed',
      services: ['TENDER'],
      termStart: '2026-11-05T00:00:00Z',
      termEnd: '2026-12-05T00:00:00Z',
      termUnit: 'P1M',
      autoRenew: true,
      isFreeTrial: false
    },
    services: ['TENDER'],
    entitlements: ENTITLEMENTS,
    message: 'Tender is active.',
    correlationId: '0HN7:00000042'
  };
  const LOCKED = { ...ENTITLEMENTS, services: [], features: [] };

  it('POSTs to saas/subscriptions/{id}/activate with the JWT and returns the answer', async () => {
    const { service, requests } = build([{ status: 200, body: ACTIVATED }], { token: 'secret-jwt' });

    const result = await service.activateSubscription(SUBSCRIPTION_ID);

    expect(requests[0].url).toBe(`${BASE_URL}/saas/subscriptions/${SUBSCRIPTION_ID}/activate`);
    expect(requests[0].method).toBe('POST');
    expect(requests[0].headers.Authorization).toBe('Bearer secret-jwt');
    expect(requests[0].body).toEqual({});
    expect(result.state).toBe('Activated');
    expect(result.subscription.status).toBe('Subscribed');
    expect(result.services).toEqual(['TENDER']);
    expect(result.entitlements.features).toEqual(['TENDER.CORE']);
  });

  it('writes the answered entitlements into the cache, so every gate sees them at once', async () => {
    const { service, requests, storage } = build([
      { status: 200, body: LOCKED },
      { status: 200, body: ACTIVATED }
    ]);

    expect((await service.getEntitlements()).features).toEqual([]);
    await service.activateSubscription(SUBSCRIPTION_ID);

    expect((await service.getEntitlements()).features).toEqual(['TENDER.CORE']);
    expect(requests).toHaveLength(2);
    // Under the same key as a read, so the next page load reuses it too.
    const keys: string[] = Object.keys(storage.items);
    expect(keys).toHaveLength(1);
    expect(JSON.parse(storage.items[keys[0]]).value.features).toEqual(['TENDER.CORE']);
  });

  it('does not let an entitlements read that was in flight overwrite the answer', async () => {
    let release: () => void = (): void => undefined;
    const held: Promise<void> = new Promise<void>((resolve) => {
      release = resolve;
    });
    const { service, requests } = build([
      { status: 200, body: LOCKED, gate: held },
      { status: 200, body: ACTIVATED }
    ]);

    const stale: Promise<ISaasEntitlements> = service.getEntitlements({ refresh: true });
    await service.activateSubscription(SUBSCRIPTION_ID);
    release();
    expect((await stale).features).toEqual([]);

    expect((await service.getEntitlements()).features).toEqual(['TENDER.CORE']);
    expect(requests).toHaveLength(2);
  });

  it('returns a Suspended answer as it is, and caches its entitlements', async () => {
    const { service } = build([
      { status: 200, body: { ...ACTIVATED, state: 'Suspended', entitlements: LOCKED, message: 'Suspended by Microsoft.' } }
    ]);

    const result = await service.activateSubscription(SUBSCRIPTION_ID);

    expect(result.state).toBe('Suspended');
    expect(result.message).toBe('Suspended by Microsoft.');
    expect((await service.getEntitlements()).features).toEqual([]);
  });

  it('rejects an answer that is not the promised shape, and caches nothing', async () => {
    const { service, storage } = build([{ status: 200, body: { state: 'Activated' } }]);

    await expect(service.activateSubscription(SUBSCRIPTION_ID)).rejects.toThrow('unexpected format');
    expect(Object.keys(storage.items)).toHaveLength(0);
  });

  it('carries a refusal through with its code and reference, logged without the token', async () => {
    const { service, storage } = build(
      [{ status: 403, body: { code: 'SAAS_NOT_TENANT_ADMIN', message: 'Admins only.', correlationId: 'c-403' } }],
      { token: 'secret-jwt' }
    );

    const error: ApiError = (await service.activateSubscription(SUBSCRIPTION_ID).catch((caught: ApiError) => caught)) as ApiError;

    expect(error.status).toBe(403);
    expect(error.code).toBe('SAAS_NOT_TENANT_ADMIN');
    expect(error.correlationId).toBe('c-403');
    expect(logged.warn[0]).toContain(`activate subscription ${SUBSCRIPTION_ID}`);
    expect(logged.warn[0]).toContain('c-403');
    expect(JSON.stringify(logged)).not.toContain('secret-jwt');
    expect(Object.keys(storage.items)).toHaveLength(0);
  });

  it('logs a failed activation at Microsoft as an error', async () => {
    const { service } = build([
      { status: 502, body: { code: 'SAAS_ACTIVATION_FAILED', message: 'Marketplace refused.', correlationId: 'c-502' } }
    ]);

    await expect(service.activateSubscription(SUBSCRIPTION_ID)).rejects.toBeInstanceOf(ApiError);
    expect(logged.error.some((line: string) => line.indexOf('SAAS_ACTIVATION_FAILED') >= 0 && line.indexOf('c-502') >= 0)).toBe(true);
  });
});
