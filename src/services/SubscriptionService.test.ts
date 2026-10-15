/**
 * Binding tests for `SubscriptionService` over the real `ApiService`, against a fake
 * transport answering the JSON the Web API answers.
 *
 * The failures worth catching are the ones that would switch a paid feature on by
 * mistake: a refusal inside an HTTP 200 read as "active", one feature's answer reused for
 * another, or a failed read cached so that "Try again" never asks again.
 */

// `@microsoft/sp-core-library` needs the SPFx page runtime, which plain jsdom has not got;
// `ApiService` pulls it in only for `Log`. Must precede the imports: the TypeScript build
// emits CommonJS without Babel's `jest.mock` hoisting.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    verbose: (): void => undefined,
    info: (): void => undefined,
    warn: (): void => undefined,
    error: (): void => undefined
  }
}));

import type { HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import { ApiError } from '../models/ApiError';
import { FeatureCode, IFeatureSubscription, ISubscriptionLinks } from '../models/FeatureSubscription';
import { ApiService } from './ApiService';
import { ApiTokenStore } from './ApiTokenStore';
import { SubscriptionService } from './SubscriptionService';

/**
 * The shape `ApiService` needs from a transport.
 *
 * Declared here rather than imported from `./ApiHttpClient`, whose module pulls
 * `@microsoft/sp-http` in as a value and so cannot load under jsdom.
 */
interface IFakeHttpClient {
  fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

interface IRecordedRequest {
  url: string;
  options: IHttpClientOptions;
}

/** What the fake API answers, per request URL - a body, or a status to fail with. */
type Responder = (url: string) => { body: unknown; status?: number };

function fakeResponse(body: unknown, status: number): HttpClientResponse {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? 'OK' : 'Error',
    text: async (): Promise<string> => (typeof body === 'string' ? body : JSON.stringify(body))
  } as unknown as HttpClientResponse;
}

function fakeTransport(respond: Responder, requests: IRecordedRequest[]): IFakeHttpClient {
  return {
    fetch: async (url: string, options: IHttpClientOptions): Promise<HttpClientResponse> => {
      requests.push({ url, options });
      const { body, status } = respond(url);
      return fakeResponse(body, status || 200);
    }
  };
}

/** The `ResponseDetail` envelope every BMDesk action wraps its result in. */
function envelope(data: unknown, success: boolean = true, message?: string): unknown {
  return { data, success, message, messageType: success ? 0 : 1 };
}

const LINKS: ISubscriptionLinks = {
  marketplaceOfferUrl: 'https://marketplace.microsoft.com/product/saas/contoso.xsproject',
  manageSubscriptionUrl: ''
};

/** ED Control bought, Minutes not - the two answers the multi-feature tests tell apart. */
const BY_FEATURE: { [feature: string]: unknown } = {
  ED_CONTROL: { feature: 'ED_CONTROL', isActive: true, subscriptionStatus: 'Active', planId: 'EDCONTROL_BASIC' },
  MINUTES: { feature: 'MINUTES', isActive: false, subscriptionStatus: 'Inactive', requiresSubscription: true }
};

function byFeature(url: string): { body: unknown } {
  const feature: string = url.substring(url.lastIndexOf('/') + 1);
  return { body: envelope(BY_FEATURE[feature]) };
}

function build(
  respond: Responder,
  requests: IRecordedRequest[],
  endpoint: string = 'Subscription'
): SubscriptionService {
  const store: ApiTokenStore = new ApiTokenStore();
  store.setToken('app-jwt');

  return new SubscriptionService(
    new ApiService(fakeTransport(respond, requests), { baseUrl: 'https://api.example.com/api' }, store),
    LINKS,
    endpoint
  );
}

async function rejection(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (error) {
    return error;
  }

  throw new Error('Expected the call to fail.');
}

describe('SubscriptionService.getFeatureSubscription', () => {
  it('reads one feature from the configured route, as the signed-in session', async () => {
    const requests: IRecordedRequest[] = [];
    const subscription: IFeatureSubscription = await build(byFeature, requests, 'RenamedSubscription')
      .getFeatureSubscription('ED_CONTROL');

    expect(requests[0].url).toBe('https://api.example.com/api/RenamedSubscription/features/ED_CONTROL');
    expect(requests[0].options.method).toBe('GET');
    // The tenant is the API's to derive from the JWT; the browser sends nothing about it.
    expect((requests[0].options.headers as Record<string, string>).Authorization).toBe('Bearer app-jwt');
    expect(subscription).toEqual({
      feature: 'ED_CONTROL',
      state: 'active',
      isActive: true,
      isEnabled: true,
      subscriptionStatus: 'Active',
      planId: 'EDCONTROL_BASIC',
      planName: undefined
    });
  });

  it('accepts a bare body as well as an enveloped one', async () => {
    const subscription: IFeatureSubscription = await build(
      () => ({ body: { feature: 'ED_CONTROL', isActive: false, subscriptionStatus: 'Pending' } }),
      []
    ).getFeatureSubscription('ED_CONTROL');

    expect(subscription.state).toBe('pending');
  });

  it('keeps one feature from unlocking another', async () => {
    const service: SubscriptionService = build(byFeature, []);

    const edControl: IFeatureSubscription = await service.getFeatureSubscription('ED_CONTROL');
    const minutes: IFeatureSubscription = await service.getFeatureSubscription('MINUTES');

    expect(edControl.isActive).toBe(true);
    expect(minutes.isActive).toBe(false);
    expect(minutes.feature).toBe('MINUTES');
  });

  it('fails, rather than answering, when the API reports a different feature', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: envelope(BY_FEATURE.ED_CONTROL) }), []).getFeatureSubscription('MINUTES')
    );

    expect(error).toBeInstanceOf(ApiError);
    expect((error as ApiError).status).toBe(200);
  });

  // HTTP 200 does not mean the API answered: a refusal inside the envelope is a failure
  // to verify, never "active" and never "inactive".
  // The API always names the feature; a body that does not - a proxy's `{}` - is not an
  // answer, and read as "not bought" it would invite a second purchase.
  it('fails, rather than reading "not bought", when the answer does not name the feature', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: envelope({ isActive: false, subscriptionStatus: 'Inactive' }) }), []).getFeatureSubscription('KYP')
    );

    expect(error).toBeInstanceOf(ApiError);
    expect((error as ApiError).status).toBe(200);
  });

  it('fails on a refusal inside an HTTP 200', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: envelope(null, false, 'Subscription lookup failed.') }), []).getFeatureSubscription('KYP')
    );

    expect((error as ApiError).message).toBe('Subscription lookup failed.');
  });

  it('fails when a successful envelope carries no subscription', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: envelope(null) }), []).getFeatureSubscription('KYP')
    );

    expect(error).toBeInstanceOf(ApiError);
  });

  it('fails on a failed status', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: 'Unauthorized', status: 401 }), []).getFeatureSubscription('KYP')
    );

    expect((error as ApiError).status).toBe(401);
  });
});

describe('SubscriptionService - caching', () => {
  it('asks once per feature, however many screens ask', async () => {
    const requests: IRecordedRequest[] = [];
    const service: SubscriptionService = build(byFeature, requests);

    await Promise.all([
      service.getFeatureSubscription('ED_CONTROL'),
      service.getFeatureSubscription('ED_CONTROL')
    ]);
    await service.getFeatureSubscription('ED_CONTROL');
    await service.getFeatureSubscription('MINUTES');

    expect(requests.map((request: IRecordedRequest) => request.url)).toEqual([
      'https://api.example.com/api/Subscription/features/ED_CONTROL',
      'https://api.example.com/api/Subscription/features/MINUTES'
    ]);
  });

  it('asks again on refresh, and keeps the newer answer', async () => {
    const requests: IRecordedRequest[] = [];
    let isActive: boolean = false;
    const service: SubscriptionService = build(
      () => ({ body: envelope({ feature: 'ED_CONTROL', isActive, subscriptionStatus: isActive ? 'Active' : 'Pending' }) }),
      requests
    );

    expect((await service.getFeatureSubscription('ED_CONTROL')).state).toBe('pending');

    // Marketplace has told the API meanwhile.
    isActive = true;
    expect((await service.getFeatureSubscription('ED_CONTROL')).state).toBe('pending');
    expect((await service.getFeatureSubscription('ED_CONTROL', { refresh: true })).state).toBe('active');
    expect((await service.getFeatureSubscription('ED_CONTROL')).state).toBe('active');
    expect(requests.length).toBe(2);
  });

  it('does not cache a failure, so a retry really retries', async () => {
    const requests: IRecordedRequest[] = [];
    let attempts: number = 0;
    const service: SubscriptionService = build((url: string) => {
      attempts++;
      return attempts === 1 ? { body: 'Server error', status: 500 } : byFeature(url);
    }, requests);

    await rejection(service.getFeatureSubscription('ED_CONTROL'));
    const subscription: IFeatureSubscription = await service.getFeatureSubscription('ED_CONTROL');

    expect(subscription.isActive).toBe(true);
    expect(requests.length).toBe(2);
  });

  // Marketplace's webhook can suspend or cancel a subscription mid-session; a cached answer
  // that never expired would hide that until the page was reloaded.
  describe('time to live', () => {
    const START: number = Date.UTC(2026, 9, 15, 9, 0, 0);
    const FIVE_MINUTES: number = 5 * 60 * 1000;
    let now: number;
    let nowSpy: jest.SpyInstance;

    beforeEach(() => {
      now = START;
      nowSpy = jest.spyOn(Date, 'now').mockImplementation(() => now);
    });

    afterEach(() => {
      nowSpy.mockRestore();
    });

    it('does not ask again within five minutes', async () => {
      const requests: IRecordedRequest[] = [];
      const service: SubscriptionService = build(byFeature, requests);

      await service.getFeatureSubscription('ED_CONTROL');
      now = START + FIVE_MINUTES - 1;
      await service.getFeatureSubscription('ED_CONTROL');

      expect(requests.length).toBe(1);
    });

    it('asks again once five minutes have passed, and keeps the newer answer', async () => {
      const requests: IRecordedRequest[] = [];
      let status: string = 'Active';
      const service: SubscriptionService = build(
        () => ({ body: envelope({ feature: 'ED_CONTROL', isActive: status === 'Active', subscriptionStatus: status }) }),
        requests
      );

      expect((await service.getFeatureSubscription('ED_CONTROL')).state).toBe('active');

      // Suspended by Marketplace meanwhile.
      status = 'Suspended';
      now = START + FIVE_MINUTES;
      expect((await service.getFeatureSubscription('ED_CONTROL')).state).toBe('suspended');
      expect((await service.getFeatureSubscription('ED_CONTROL')).state).toBe('suspended');
      expect(requests.length).toBe(2);
    });

    it('lets a written answer expire as well', async () => {
      const requests: IRecordedRequest[] = [];
      const service: SubscriptionService = build(
        () => ({ body: envelope({ feature: 'ED_CONTROL', isActive: false, isEnabled: true, subscriptionStatus: 'Inactive' }) }),
        requests
      );

      await service.setFeatureEnabled('ED_CONTROL', true);
      await service.getFeatureSubscription('ED_CONTROL');
      expect(requests.length).toBe(1);

      now = START + FIVE_MINUTES;
      await service.getFeatureSubscription('ED_CONTROL');
      expect(requests.length).toBe(2);
    });
  });

  it('keeps each instance its own cache', async () => {
    const requests: IRecordedRequest[] = [];

    await build(byFeature, requests).getFeatureSubscription('ED_CONTROL');
    await build(byFeature, requests).getFeatureSubscription('ED_CONTROL');

    expect(requests.length).toBe(2);
  });
});

describe('SubscriptionService.setFeatureEnabled', () => {
  /** What `POST …/enable` answers for a feature switched on but not bought yet. */
  const SWITCHED_ON_UNPAID: unknown = envelope({
    feature: 'ED_CONTROL',
    isActive: false,
    isEnabled: true,
    subscriptionStatus: 'Inactive',
    requiresSubscription: true
  });

  it('switches a feature on through its enable action, as the signed-in session', async () => {
    const requests: IRecordedRequest[] = [];
    const subscription: IFeatureSubscription = await build(() => ({ body: SWITCHED_ON_UNPAID }), requests)
      .setFeatureEnabled('ED_CONTROL', true);

    expect(requests[0].url).toBe('https://api.example.com/api/Subscription/features/ED_CONTROL/enable');
    expect(requests[0].options.method).toBe('POST');
    expect((requests[0].options.headers as Record<string, string>).Authorization).toBe('Bearer app-jwt');
    // Switched on is a choice, not a purchase.
    expect(subscription.isEnabled).toBe(true);
    expect(subscription.isActive).toBe(false);
  });

  it('switches a feature off through its disable action', async () => {
    const requests: IRecordedRequest[] = [];
    await build(
      () => ({ body: envelope({ feature: 'ED_CONTROL', isActive: false, isEnabled: false, subscriptionStatus: 'Active' }) }),
      requests
    ).setFeatureEnabled('ED_CONTROL', false);

    expect(requests[0].url).toBe('https://api.example.com/api/Subscription/features/ED_CONTROL/disable');
  });

  it('replaces the cached answer with the one the write returned', async () => {
    const requests: IRecordedRequest[] = [];
    const service: SubscriptionService = build((url: string) => (
      url.indexOf('/enable') > -1
        ? { body: SWITCHED_ON_UNPAID }
        : { body: envelope({ feature: 'ED_CONTROL', isActive: false, isEnabled: false, subscriptionStatus: 'Inactive' }) }
    ), requests);

    expect((await service.getFeatureSubscription('ED_CONTROL')).isEnabled).toBe(false);
    await service.setFeatureEnabled('ED_CONTROL', true);

    expect((await service.getFeatureSubscription('ED_CONTROL')).isEnabled).toBe(true);
    expect(requests.length).toBe(2);
  });

  it('fails on a refusal, e.g. a role that may not change features', async () => {
    const error: unknown = await rejection(
      build(
        () => ({ body: envelope(null, false, 'You do not have permission to change this feature.') }),
        []
      ).setFeatureEnabled('ED_CONTROL', true)
    );

    expect((error as ApiError).message).toBe('You do not have permission to change this feature.');
  });

  it('fails when the API answers about a different feature', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: SWITCHED_ON_UNPAID }), []).setFeatureEnabled('MINUTES', true)
    );

    expect(error).toBeInstanceOf(ApiError);
  });
});

describe('SubscriptionService.setFeatureEnabled - return URL', () => {
  const ANSWER: unknown = envelope({ feature: 'ED_CONTROL', isActive: false, isEnabled: true, subscriptionStatus: 'Inactive' });
  const RETURN_URL: string = 'https://contoso.sharepoint.com/sites/xs/SitePages/App.aspx?debug=true';

  it('posts the return URL when switching on', async () => {
    const requests: IRecordedRequest[] = [];
    await build(() => ({ body: ANSWER }), requests).setFeatureEnabled('ED_CONTROL', true, RETURN_URL);

    expect(JSON.parse(requests[0].options.body as string)).toEqual({ returnUrl: RETURN_URL });
  });

  it('posts an empty object, not an undefined property, without a return URL', async () => {
    const requests: IRecordedRequest[] = [];
    await build(() => ({ body: ANSWER }), requests).setFeatureEnabled('ED_CONTROL', true);

    expect(requests[0].options.body).toBe('{}');
  });

  it('does not send the return URL when switching off', async () => {
    const requests: IRecordedRequest[] = [];
    await build(() => ({ body: ANSWER }), requests).setFeatureEnabled('ED_CONTROL', false, RETURN_URL);

    expect(requests[0].options.body).toBe('{}');
  });
});

describe('SubscriptionService.activateMarketplacePurchase', () => {
  const TOKEN: string = 'abc+def/ghi==';
  const ACTIVATED: unknown = envelope({
    feature: 'KYP',
    isActive: true,
    isEnabled: true,
    subscriptionStatus: 'Subscribed',
    planId: 'kyp-plan',
    planName: 'KYP'
  });

  it('posts the token in the body to marketplace/activate, never in the URL', async () => {
    const requests: IRecordedRequest[] = [];
    const subscription: IFeatureSubscription = await build(() => ({ body: ACTIVATED }), requests)
      .activateMarketplacePurchase(TOKEN);

    expect(requests[0].url).toBe('https://api.example.com/api/Subscription/marketplace/activate');
    expect(requests[0].url).not.toContain('abc');
    expect(requests[0].options.method).toBe('POST');
    expect(JSON.parse(requests[0].options.body as string)).toEqual({ token: TOKEN });
    expect(subscription.feature).toBe('KYP');
    expect(subscription.isActive).toBe(true);
  });

  it('seeds the cache for the activated feature, so a read asks nothing', async () => {
    const requests: IRecordedRequest[] = [];
    const service: SubscriptionService = build(() => ({ body: ACTIVATED }), requests);

    await service.activateMarketplacePurchase(TOKEN);
    const read: IFeatureSubscription = await service.getFeatureSubscription('KYP');

    expect(read.isActive).toBe(true);
    expect(requests.length).toBe(1);
  });

  it('fails on a refusal inside an HTTP 200, and caches nothing', async () => {
    const requests: IRecordedRequest[] = [];
    const service: SubscriptionService = build(
      (url: string) => (
        url.indexOf('/marketplace/activate') > -1
          ? { body: envelope(null, false, 'This purchase belongs to another account.') }
          : { body: envelope({ feature: 'KYP', isActive: false, subscriptionStatus: 'Inactive' }) }
      ),
      requests
    );

    const error: unknown = await rejection(service.activateMarketplacePurchase(TOKEN));
    expect((error as ApiError).message).toBe('This purchase belongs to another account.');

    expect((await service.getFeatureSubscription('KYP')).isActive).toBe(false);
    expect(requests.length).toBe(2);
  });

  it('fails when the answer names an unknown feature', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: envelope({ feature: 'BIM', isActive: true }) }), []).activateMarketplacePurchase(TOKEN)
    );

    expect(error).toBeInstanceOf(ApiError);
    expect((error as ApiError).status).toBe(200);
  });

  it('fails when the answer carries no subscription', async () => {
    const error: unknown = await rejection(
      build(() => ({ body: envelope(null) }), []).activateMarketplacePurchase(TOKEN)
    );

    expect(error).toBeInstanceOf(ApiError);
  });
});

describe('SubscriptionService.getSubscriptionStatus', () => {
  it('reads every feature, each with its own answer', async () => {
    const requests: IRecordedRequest[] = [];
    const subscriptions: IFeatureSubscription[] = await build(
      () => ({ body: envelope([BY_FEATURE.ED_CONTROL, BY_FEATURE.MINUTES]) }),
      requests
    ).getSubscriptionStatus();

    expect(requests[0].url).toBe('https://api.example.com/api/Subscription/features');
    expect(
      subscriptions.map((item: IFeatureSubscription): [FeatureCode, boolean] => [item.feature, item.isActive])
    ).toEqual([
      ['ED_CONTROL', true],
      ['MINUTES', false]
    ]);
  });

  it('exposes the configured Marketplace links', () => {
    expect(build(byFeature, []).links).toBe(LINKS);
  });
});
