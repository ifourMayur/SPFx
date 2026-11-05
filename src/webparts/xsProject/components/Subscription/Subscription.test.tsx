/**
 * Render tests for the Subscription page: the list, the Enable round trip, activating the
 * purchase on the return from the Marketplace, and the per-subscription Activate button.
 *
 * The page's two exits to the browser - reading the current URL and leaving for the
 * Marketplace - are injected props, because jsdom cannot spy on `window.location`.
 *
 * Every `ReactDOM.render` here is unmounted in `afterEach`.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { MemoryRouter, Route, Routes, useLocation } from 'react-router-dom';

import Subscription from './Subscription';
import SubscriptionRoute from './SubscriptionRoute';
import { ApiError } from '../../../../models/ApiError';
import {
  ISaasActivationResult,
  ISaasCatalog,
  ISaasEntitlements,
  ISaasPurchaseIntent,
  ISaasServiceStatus,
  ISaasSubscription,
  SAAS_MANAGE_BILLING_URL
} from '../../../../models/Saas';
import { IGetEntitlementsOptions, ISaasService } from '../../../../services/SaasService';

const PAGE_URL: string = 'https://contoso.sharepoint.com/sites/xs/SitePages/App.aspx?mode=1#/subscription';
const OFFER_URL: string = 'https://appsource.microsoft.com/product/web-apps/xsproject.tender';
const SUBSCRIPTION_ID: string = '0f8fad5b-d9cb-469f-a165-70867728950e';
const PENDING_ID: string = '11111111-2222-3333-4444-555555555555';
const RETURN_QUERY: string = `?saas=returned&service=TENDER&subscriptionId=${SUBSCRIPTION_ID}`;

function detail(serviceKey: string, name: string, status: string, overrides?: Partial<ISaasServiceStatus>): ISaasServiceStatus {
  return {
    serviceKey,
    name,
    description: null,
    status,
    source: status === 'Active' ? 'Marketplace' : null,
    validTo: null,
    termEnd: null,
    isFreeTrial: false,
    autoRenew: null,
    canPurchase: status === 'NotPurchased' || status === 'Inactive',
    marketplaceSubscriptionId: null,
    ...overrides
  };
}

function entitlements(isTenantAdmin: boolean, tenderStatus: string, overrides?: Partial<ISaasEntitlements>): ISaasEntitlements {
  return {
    tenantKnown: true,
    entraTenantId: 't',
    isTenantAdmin,
    isBillingEnforced: true,
    services: tenderStatus === 'Active' ? ['TENDER'] : [],
    features: tenderStatus === 'Active' ? ['TENDER.CORE'] : [],
    details: [
      detail('TENDER', 'Tender', tenderStatus),
      detail('MINUTES', 'Minutes of Meeting', 'Active', { isFreeTrial: true, source: 'Marketplace' })
    ],
    cacheSeconds: 300,
    evaluatedOn: '',
    ...overrides
  };
}

function subscription(marketplaceSubscriptionId: string, status: string, overrides?: Partial<ISaasSubscription>): ISaasSubscription {
  return {
    marketplaceSubscriptionId,
    subscriptionName: `Tender ${status}`,
    offerId: 'xs-tender',
    planId: 'monthly',
    planName: 'Monthly',
    quantity: null,
    status,
    services: ['TENDER'],
    termStart: null,
    termEnd: null,
    termUnit: null,
    autoRenew: true,
    isFreeTrial: false,
    ...overrides
  };
}

function activation(state: string, answered: ISaasEntitlements, overrides?: Partial<ISaasActivationResult>): ISaasActivationResult {
  return {
    state,
    subscription: subscription(SUBSCRIPTION_ID, state === 'Activated' || state === 'AlreadyActive' ? 'Subscribed' : state),
    services: ['TENDER'],
    entitlements: answered,
    message: `${state} message from the server.`,
    correlationId: 'corr-activate',
    ...overrides
  };
}

const CATALOG: ISaasCatalog = {
  services: [
    {
      serviceKey: 'MINUTES',
      name: 'Minutes of Meeting',
      description: 'Record meeting minutes.',
      displayOrder: 2,
      isAvailable: true,
      hasOffer: true,
      features: [{ featureKey: 'MINUTES.CORE', name: 'Minutes' }],
      plans: []
    },
    {
      serviceKey: 'TENDER',
      name: 'Tender',
      description: 'Run tenders.',
      displayOrder: 1,
      isAvailable: true,
      hasOffer: true,
      features: [{ featureKey: 'TENDER.CORE', name: 'Tender' }],
      plans: [{ offerId: 'xs-tender', planId: 'monthly', name: 'Monthly', pricingModel: 'Flat rate', billingTerm: 'Monthly' }]
    }
  ]
};

type Answer<T> = T | Error | Promise<T>;

interface IStubOptions {
  /** Entitlement answers, queued: each call takes the next one and the last repeats. */
  entitlements: Array<ISaasEntitlements | Error>;
  intent?: ISaasPurchaseIntent | Error;
  /** Activation answers, queued the same way; a promise holds the call until it settles. */
  activations?: Array<Answer<ISaasActivationResult>>;
  subscriptions?: ISaasSubscription[];
}

interface IStub {
  service: ISaasService;
  entitlementCalls: IGetEntitlementsOptions[];
  intents: Array<{ serviceKey: string; returnUrl: string }>;
  activated: string[];
  subscriptionCalls: { count: number };
}

function next<T>(queue: T[]): T {
  return queue.length > 1 ? (queue.shift() as T) : queue[0];
}

function stubService(options: IStubOptions): IStub {
  const entitlementCalls: IGetEntitlementsOptions[] = [];
  const intents: Array<{ serviceKey: string; returnUrl: string }> = [];
  const activated: string[] = [];
  const subscriptionCalls = { count: 0 };
  const intent: ISaasPurchaseIntent | Error = options.intent || { intentId: 'i-1', marketplaceOfferUrl: OFFER_URL, expiresOn: '' };
  const activations: Array<Answer<ISaasActivationResult>> = options.activations || [new Error('not expected')];

  return {
    entitlementCalls,
    intents,
    activated,
    subscriptionCalls,
    service: {
      getEntitlements: async (callOptions?: IGetEntitlementsOptions): Promise<ISaasEntitlements> => {
        entitlementCalls.push(callOptions || {});
        const answer: ISaasEntitlements | Error = next(options.entitlements);
        if (answer instanceof Error) {
          throw answer;
        }
        return answer;
      },
      getCatalog: async () => CATALOG,
      getSubscriptions: async (): Promise<ISaasSubscription[]> => {
        subscriptionCalls.count++;
        return options.subscriptions || [];
      },
      createPurchaseIntent: async (serviceKey: string, returnUrl: string): Promise<ISaasPurchaseIntent> => {
        intents.push({ serviceKey, returnUrl });
        if (intent instanceof Error) {
          throw intent;
        }
        return intent;
      },
      activateSubscription: async (marketplaceSubscriptionId: string): Promise<ISaasActivationResult> => {
        activated.push(marketplaceSubscriptionId);
        const answer: Answer<ISaasActivationResult> = next(activations);
        if (answer instanceof Error) {
          throw answer;
        }
        return answer;
      },
      invalidate: (): void => undefined
    }
  };
}

/** A promise settled from outside, to hold a call while the page is inspected. */
function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let settle: (value: T) => void = (): void => undefined;
  const promise: Promise<T> = new Promise<T>((resolve) => {
    settle = resolve;
  });
  return { promise, resolve: settle };
}

/** Lets the stubbed promises and the state updates they trigger settle. */
async function flush(): Promise<void> {
  for (let i: number = 0; i < 10; i++) {
    await Promise.resolve();
  }
}

describe('Subscription', () => {
  let host: HTMLDivElement;
  let navigated: string[];
  let cleared: number;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    navigated = [];
    cleared = 0;
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  async function render(service: ISaasService, query: string = ''): Promise<void> {
    await act(async () => {
      ReactDOM.render(
        <Subscription
          saasService={service}
          query={query}
          onClearQuery={(): void => {
            cleared++;
          }}
          getCurrentUrl={(): string => PAGE_URL}
          navigateToUrl={(url: string): void => {
            navigated.push(url);
          }}
        />,
        host
      );
      await flush();
    });
  }

  function card(serviceKey: string): HTMLElement {
    return host.querySelector(`li[data-service-key="${serviceKey}"]`) as HTMLElement;
  }

  function subscriptionCard(id: string): HTMLElement {
    return host.querySelector(`li[data-subscription-id="${id}"]`) as HTMLElement;
  }

  function banner(kind: string): HTMLElement | null {
    return host.querySelector(`[data-banner="${kind}"]`);
  }

  function button(text: string, within: ParentNode = host): HTMLButtonElement {
    return Array.prototype.slice
      .call(within.querySelectorAll('button'))
      .filter((candidate: HTMLButtonElement) => (candidate.textContent || '').trim() === text)[0];
  }

  async function click(target: HTMLButtonElement): Promise<void> {
    expect(target).toBeDefined();

    await act(async () => {
      target.click();
      await flush();
    });
  }

  describe('the service list', () => {
    it('lists the catalog services in order, each with its status', async () => {
      const stub = stubService({ entitlements: [entitlements(true, 'NotPurchased')] });
      await render(stub.service);

      const keys: string[] = Array.prototype.slice
        .call(host.querySelectorAll('li[data-service-key]'))
        .map((item: HTMLElement) => item.getAttribute('data-service-key') || '');
      expect(keys).toEqual(['TENDER', 'MINUTES']);

      expect(card('TENDER').textContent).toContain('Not enabled');
      expect(card('TENDER').textContent).toContain('Run tenders.');
      expect(card('TENDER').textContent).toContain('Monthly (Monthly, Flat rate)');
      expect(card('MINUTES').textContent).toContain('Active');
      expect(card('MINUTES').textContent).toContain('Free trial');
      expect(card('MINUTES').textContent).toContain('Microsoft Marketplace');
      // The page always asks for the current status, and nothing was activated.
      expect(stub.entitlementCalls).toEqual([{ refresh: true }]);
      expect(stub.activated).toEqual([]);
      expect(cleared).toBe(0);
    });

    it('gives an admin Enable buttons and a Manage billing link that opens in a new tab', async () => {
      await render(stubService({ entitlements: [entitlements(true, 'NotPurchased')] }).service);

      expect(card('TENDER').querySelector('button')?.textContent).toBe('Enable Tender');
      // Already active: nothing to buy.
      expect(card('MINUTES').querySelector('button')).toBeNull();

      const link: HTMLAnchorElement = host.querySelector(`a[href="${SAAS_MANAGE_BILLING_URL}"]`) as HTMLAnchorElement;
      expect(link.textContent).toBe('Manage billing');
      expect(link.target).toBe('_blank');
      expect(link.rel).toContain('noopener');
    });

    it('is read-only for a non-admin, and says who can buy', async () => {
      await render(stubService({ entitlements: [entitlements(false, 'NotPurchased')] }).service);

      expect(host.querySelectorAll('li[data-service-key] button')).toHaveLength(0);
      expect(host.querySelector(`a[href="${SAAS_MANAGE_BILLING_URL}"]`)).toBeNull();
      expect(host.textContent).toContain('Only a Global Administrator or SharePoint Administrator can enable services');
    });

    it('explains suspended, cancelled, unavailable and expiring services', async () => {
      const soon: string = new Date(Date.now() + 5 * 24 * 60 * 60 * 1000).toISOString();
      await render(
        stubService({
          entitlements: [
            entitlements(true, 'Suspended', {
              details: [
                detail('TENDER', 'Tender', 'Suspended'),
                detail('MINUTES', 'Minutes of Meeting', 'Active', { validTo: soon }),
                detail('KYP', 'Know Your Property', 'Inactive'),
                detail('ED_CONTROL', 'ED Control', 'Unavailable')
              ]
            })
          ]
        }).service
      );

      expect(card('TENDER').textContent).toContain('Microsoft 365 admin center');
      expect(card('MINUTES').textContent).toContain('Minutes of Meeting ends on');
      expect(card('KYP').textContent).toContain('cancelled or has ended');
      expect(card('ED_CONTROL').textContent).toContain('currently unavailable');
    });

    it('says why the page could not load, with the reference, and retries', async () => {
      const stub = stubService({
        entitlements: [new ApiError('Unauthorized', { status: 401, correlationId: 'corr-401' }), entitlements(false, 'NotPurchased')]
      });
      await render(stub.service);

      expect(host.textContent).toContain('Your subscriptions could not be loaded.');
      expect(host.textContent).toContain('session has expired');
      expect(host.textContent).toContain('Reference: corr-401');

      await click(button('Try again'));

      expect(card('TENDER')).not.toBeNull();
    });

    it('says so when the API cannot be reached', async () => {
      await render(stubService({ entitlements: [new ApiError('Unable to reach', { status: 0 })] }).service);

      expect(host.textContent).toContain('could not be reached');
    });

    it('notes a tenant whose services are included rather than billed', async () => {
      await render(stubService({ entitlements: [entitlements(false, 'Active', { isBillingEnforced: false })] }).service);

      expect(host.textContent).toContain('nothing needs to be bought');
    });
  });

  describe('enabling a service', () => {
    it('starts a purchase with this page as the return URL, then leaves for the offer in the same tab', async () => {
      const stub = stubService({ entitlements: [entitlements(true, 'NotPurchased')] });
      await render(stub.service);

      await click(button('Enable Tender'));

      expect(stub.intents).toEqual([
        {
          serviceKey: 'TENDER',
          returnUrl: 'https://contoso.sharepoint.com/sites/xs/SitePages/App.aspx?mode=1#/subscription?saas=returned&service=TENDER'
        }
      ]);
      expect(navigated).toEqual([OFFER_URL]);
    });

    it('reports a refused purchase with friendly text and its reference, and stays on the page', async () => {
      const refused: ApiError = new ApiError('raw', { status: 400, code: 'SAAS_RETURN_URL_INVALID', correlationId: 'corr-7' });
      await render(stubService({ entitlements: [entitlements(true, 'NotPurchased')], intent: refused }).service);

      await click(button('Enable Tender'));

      expect(navigated).toEqual([]);
      expect(host.textContent).toContain('not registered for your organisation');
      expect(host.textContent).toContain('Reference: corr-7');
      expect((card('TENDER').querySelector('button') as HTMLButtonElement).disabled).toBe(false);
    });

    it('refuses to follow an offer URL that is not https', async () => {
      const stub = stubService({
        entitlements: [entitlements(true, 'NotPurchased')],
        intent: { intentId: 'i', marketplaceOfferUrl: 'http://appsource.microsoft.com/offer', expiresOn: '' }
      });
      await render(stub.service);

      await click(button('Enable Tender'));

      expect(navigated).toEqual([]);
      expect(host.textContent).toContain('Marketplace address for this service is not valid');
    });
  });

  describe('the return from the Marketplace', () => {
    it('activates the subscription it names, shows the services from the answer and clears the query', async () => {
      const stub = stubService({
        entitlements: [entitlements(true, 'NotPurchased')],
        activations: [activation('Activated', entitlements(true, 'Active'))],
        subscriptions: [subscription(SUBSCRIPTION_ID, 'Subscribed')]
      });
      await render(stub.service, RETURN_QUERY);

      expect(stub.activated).toEqual([SUBSCRIPTION_ID]);
      expect(banner('activated')?.textContent).toContain('Tender is active.');
      // The rows come from the activation's answer, not from another entitlements read.
      expect(stub.entitlementCalls).toEqual([]);
      expect(card('TENDER').textContent).toContain('Active');
      expect(card('TENDER').className).toContain('highlighted');
      expect(subscriptionCard(SUBSCRIPTION_ID).textContent).toContain('Active');
      expect(cleared).toBe(1);
    });

    it('says the subscription is activating while the call runs', async () => {
      const held = deferred<ISaasActivationResult>();
      const stub = stubService({ entitlements: [entitlements(true, 'NotPurchased')], activations: [held.promise] });
      await render(stub.service, RETURN_QUERY);

      expect(banner('activating')?.textContent).toContain('Activating your subscription');
      expect(host.textContent).not.toContain('Loading your subscriptions');
      // The query is cleared at once, so a reload meanwhile does not activate again.
      expect(cleared).toBe(1);

      await act(async () => {
        held.resolve(activation('AlreadyActive', entitlements(true, 'Active')));
        await flush();
      });

      expect(banner('activating')).toBeNull();
      expect(banner('activated')?.textContent).toContain('Tender is active.');
    });

    it.each(['Suspended', 'Unsubscribed'])('warns with the server message when Microsoft reports %s', async (state: string) => {
      const stub = stubService({
        entitlements: [entitlements(true, 'NotPurchased')],
        activations: [activation(state, entitlements(true, 'NotPurchased'))]
      });
      await render(stub.service, RETURN_QUERY);

      const warning: HTMLElement | null = banner('notActivated');
      expect(warning?.textContent).toContain('The subscription was not activated.');
      expect(warning?.textContent).toContain(`${state} message from the server.`);
      expect(warning?.textContent).toContain('Reference: corr-activate');
      expect(banner('activated')).toBeNull();
      expect(card('TENDER').textContent).toContain('Not enabled');
    });

    it('reports a failed activation with its reference, loads the page, and repeats the activation on Try again', async () => {
      const stub = stubService({
        entitlements: [entitlements(true, 'NotPurchased')],
        activations: [
          new ApiError('Marketplace refused.', { status: 502, code: 'SAAS_ACTIVATION_FAILED', correlationId: 'corr-502' }),
          activation('Activated', entitlements(true, 'Active'))
        ]
      });
      await render(stub.service, RETURN_QUERY);

      const failure: HTMLElement | null = banner('failed');
      expect(failure?.textContent).toContain('Your subscription could not be activated.');
      expect(failure?.textContent).toContain('Reference: corr-502');
      expect(stub.entitlementCalls).toEqual([{ refresh: true }]);
      expect(card('TENDER').textContent).toContain('Not enabled');

      await click(button('Try again', failure as HTMLElement));

      expect(stub.activated).toEqual([SUBSCRIPTION_ID, SUBSCRIPTION_ID]);
      expect(banner('failed')).toBeNull();
      expect(banner('activated')?.textContent).toContain('Tender is active.');
      expect(card('TENDER').textContent).toContain('Active');
    });

    it('does not offer Try again for a refusal, and tells a non-admin whom to ask', async () => {
      const stub = stubService({
        entitlements: [entitlements(false, 'NotPurchased')],
        activations: [new ApiError('Admins only.', { status: 403, code: 'SAAS_NOT_TENANT_ADMIN', correlationId: 'corr-403' })]
      });
      await render(stub.service, RETURN_QUERY);

      const failure: HTMLElement = banner('failed') as HTMLElement;
      expect(failure.textContent).toContain('ask one of them to open this page');
      expect(failure.textContent).toContain('Reference: corr-403');
      expect(failure.querySelector('button')).toBeNull();
    });

    it('activates nothing when the buyer came back without a purchase', async () => {
      const stub = stubService({ entitlements: [entitlements(true, 'NotPurchased')] });
      await render(stub.service, '?saas=returned&service=TENDER');

      expect(stub.activated).toEqual([]);
      expect(stub.entitlementCalls).toEqual([{ refresh: true }]);
      expect(banner('returned')?.textContent).toContain('nothing was activated');
      expect(cleared).toBe(1);
    });

    it('ignores a subscriptionId that is not a GUID', async () => {
      const stub = stubService({ entitlements: [entitlements(true, 'NotPurchased')] });
      await render(stub.service, '?saas=returned&service=TENDER&subscriptionId=not-a-guid');

      expect(stub.activated).toEqual([]);
      expect(banner('returned')).not.toBeNull();
    });
  });

  describe('Marketplace subscriptions', () => {
    const SUBSCRIPTIONS: ISaasSubscription[] = [
      subscription(PENDING_ID, 'PendingFulfillmentStart'),
      subscription(SUBSCRIPTION_ID, 'Subscribed', { termEnd: '2026-12-05T00:00:00Z' })
    ];

    it('offers an admin Activate for a subscription awaiting activation, and activates it', async () => {
      const stub = stubService({
        entitlements: [entitlements(true, 'NotPurchased')],
        activations: [activation('Activated', entitlements(true, 'Active'))],
        subscriptions: SUBSCRIPTIONS
      });
      await render(stub.service);

      expect(subscriptionCard(PENDING_ID).textContent).toContain('Awaiting activation');
      expect(subscriptionCard(SUBSCRIPTION_ID).querySelector('button')).toBeNull();
      expect(stub.subscriptionCalls.count).toBe(1);

      await click(button('Activate', subscriptionCard(PENDING_ID)));

      expect(stub.activated).toEqual([PENDING_ID]);
      expect(banner('activated')?.textContent).toContain('Tender is active.');
      expect(card('TENDER').textContent).toContain('Active');
      // The rows come from the answer; only the subscriptions are read again.
      expect(stub.entitlementCalls).toHaveLength(1);
      expect(stub.subscriptionCalls.count).toBe(2);
    });

    it('does not offer Activate to a non-admin, and says who can', async () => {
      await render(stubService({ entitlements: [entitlements(false, 'NotPurchased')], subscriptions: SUBSCRIPTIONS }).service);

      expect(host.querySelectorAll('li[data-subscription-id] button')).toHaveLength(0);
      expect(subscriptionCard(PENDING_ID).textContent).toContain('Bought but not activated yet');
    });

    it('disables every action while an activation runs', async () => {
      const held = deferred<ISaasActivationResult>();
      const stub = stubService({
        entitlements: [entitlements(true, 'NotPurchased')],
        activations: [held.promise],
        subscriptions: SUBSCRIPTIONS
      });
      await render(stub.service);

      await click(button('Activate', subscriptionCard(PENDING_ID)));

      expect(button('Activating…').disabled).toBe(true);
      expect(button('Refresh').disabled).toBe(true);
      expect(button('Enable Tender').disabled).toBe(true);

      await act(async () => {
        held.resolve(activation('Activated', entitlements(true, 'Active')));
        await flush();
      });

      expect(button('Refresh').disabled).toBe(false);
    });

    it('shows nothing of the section while there are no subscriptions', async () => {
      await render(stubService({ entitlements: [entitlements(true, 'NotPurchased')] }).service);

      expect(host.textContent).not.toContain('Marketplace subscriptions');
    });
  });
});

describe('SubscriptionRoute', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  function LocationProbe(): React.ReactElement {
    const location = useLocation();
    return <output id="location">{`${location.pathname}${location.search}`}</output>;
  }

  it('hands the query to the page, which activates once, and clears it from the URL', async () => {
    const stub = stubService({
      entitlements: [entitlements(true, 'NotPurchased')],
      activations: [activation('Activated', entitlements(true, 'Active'))]
    });

    await act(async () => {
      ReactDOM.render(
        <MemoryRouter initialEntries={[`/subscription${RETURN_QUERY}`]}>
          <Routes>
            <Route path="/subscription" element={<SubscriptionRoute saasService={stub.service} />} />
          </Routes>
          <LocationProbe />
        </MemoryRouter>,
        host
      );
      await flush();
    });

    expect(stub.activated).toEqual([SUBSCRIPTION_ID]);
    expect(host.querySelector('[data-banner="activated"]')).not.toBeNull();
    expect((host.querySelector('#location') as HTMLElement).textContent).toBe('/subscription');
  });
});
