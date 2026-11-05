import { ApiError } from './ApiError';
import {
  DEFAULT_ENTITLEMENT_CACHE_SECONDS,
  ISaasCatalog,
  ISaasEntitlements,
  ISaasServiceStatus,
  buildPurchaseReturnUrl,
  canActivateSubscription,
  findServiceStatus,
  getSaasErrorInfo,
  getServiceEndDate,
  getServiceStatusLabel,
  getServiceStatusMessage,
  getServiceStatusTone,
  getSourceLabel,
  getSubscriptionStatusLabel,
  getSubscriptionStatusTone,
  isActivationSuccessful,
  isExpiringSoon,
  isFeatureEnabled,
  isRetryableSaasError,
  isSaasReturnFragment,
  isSafeOfferUrl,
  mergeServiceRows,
  parseSaasReturnQuery,
  toSaasActivationResult,
  toSaasCatalog,
  toSaasEntitlements,
  toSaasSubscriptions
} from './Saas';

const NOW: Date = new Date('2026-11-04T12:00:00Z');

const SUBSCRIPTION_ID: string = '0f8fad5b-d9cb-469f-a165-70867728950e';

function daysFromNow(days: number): string {
  return new Date(NOW.getTime() + days * 24 * 60 * 60 * 1000).toISOString();
}

function status(overrides: Partial<ISaasServiceStatus>): ISaasServiceStatus {
  return {
    serviceKey: 'TENDER',
    name: 'Tender',
    description: null,
    status: 'Active',
    source: 'Marketplace',
    validTo: null,
    termEnd: null,
    isFreeTrial: false,
    autoRenew: null,
    canPurchase: false,
    marketplaceSubscriptionId: null,
    ...overrides
  };
}

function entitlements(overrides?: Partial<ISaasEntitlements>): ISaasEntitlements {
  return {
    tenantKnown: true,
    entraTenantId: 't',
    isTenantAdmin: false,
    isBillingEnforced: true,
    services: ['TENDER'],
    features: ['TENDER.CORE'],
    details: [status({})],
    cacheSeconds: 60,
    evaluatedOn: '2026-11-04T12:00:00Z',
    ...overrides
  };
}

describe('toSaasEntitlements', () => {
  it('keeps a complete response as it is', () => {
    const wire = entitlements();
    expect(toSaasEntitlements(JSON.parse(JSON.stringify(wire)))).toEqual(wire);
  });

  it('defaults to the closed answer when fields are missing', () => {
    const result = toSaasEntitlements({ details: [{ serviceKey: 'KYP' }] }) as ISaasEntitlements;

    expect(result.isTenantAdmin).toBe(false);
    expect(result.features).toEqual([]);
    expect(result.cacheSeconds).toBe(DEFAULT_ENTITLEMENT_CACHE_SECONDS);
    expect(result.details[0].status).toBe('NotPurchased');
    expect(result.details[0].name).toBe('KYP');
    expect(result.details[0].canPurchase).toBe(false);
  });

  it('ignores a non-positive cacheSeconds and nulls in the lists', () => {
    const result = toSaasEntitlements({ cacheSeconds: 0, features: ['A', null, 3] }) as ISaasEntitlements;

    expect(result.cacheSeconds).toBe(DEFAULT_ENTITLEMENT_CACHE_SECONDS);
    expect(result.features).toEqual(['A']);
  });

  it('rejects a body that is not an object', () => {
    expect(toSaasEntitlements('<html>')).toBeUndefined();
    expect(toSaasEntitlements(null)).toBeUndefined();
    expect(toSaasEntitlements([])).toBeUndefined();
  });
});

describe('toSaasCatalog', () => {
  it('defaults missing lists and drops services without a key', () => {
    const catalog = toSaasCatalog({
      services: [{ serviceKey: 'TENDER', name: 'Tender', plans: [{ planId: 'monthly' }] }, { name: 'nameless' }]
    }) as ISaasCatalog;

    expect(catalog.services).toHaveLength(1);
    expect(catalog.services[0].features).toEqual([]);
    expect(catalog.services[0].plans[0].name).toBe('monthly');
  });

  it('rejects a body that is not an object', () => {
    expect(toSaasCatalog(undefined)).toBeUndefined();
  });
});

describe('isFeatureEnabled', () => {
  it('is true only for a feature the server listed', () => {
    expect(isFeatureEnabled(entitlements(), 'TENDER.CORE')).toBe(true);
    expect(isFeatureEnabled(entitlements(), 'MINUTES.CORE')).toBe(false);
  });

  it('compares keys case-insensitively', () => {
    expect(isFeatureEnabled(entitlements(), 'tender.core')).toBe(true);
  });

  it('fails closed without entitlements or a key', () => {
    expect(isFeatureEnabled(undefined, 'TENDER.CORE')).toBe(false);
    expect(isFeatureEnabled(entitlements(), '')).toBe(false);
  });

  it('does not infer a feature from an active service', () => {
    // The feature list is the server's evaluation; an Active detail alone grants nothing.
    expect(isFeatureEnabled(entitlements({ features: [] }), 'TENDER.CORE')).toBe(false);
  });
});

describe('findServiceStatus', () => {
  it('finds a service by key, case-insensitively', () => {
    expect(findServiceStatus(entitlements(), 'tender')?.name).toBe('Tender');
    expect(findServiceStatus(entitlements(), 'KYP')).toBeUndefined();
    expect(findServiceStatus(undefined, 'TENDER')).toBeUndefined();
  });
});

describe('status presentation', () => {
  it.each([
    ['Active', 'Active', 'success'],
    ['Pending', 'Activation pending', 'warning'],
    ['Suspended', 'Suspended', 'error'],
    ['Inactive', 'Inactive', 'neutral'],
    ['NotPurchased', 'Not enabled', 'neutral'],
    ['Unavailable', 'Unavailable', 'neutral']
  ])('labels %s as "%s" in the %s tone', (value: string, label: string, tone: string) => {
    expect(getServiceStatusLabel(value)).toBe(label);
    expect(getServiceStatusTone(value)).toBe(tone);
  });

  it('shows an unknown status as the server wrote it', () => {
    expect(getServiceStatusLabel('Frozen')).toBe('Frozen');
    expect(getServiceStatusTone('Frozen')).toBe('neutral');
  });

  it('describes each source', () => {
    expect(getSourceLabel('Marketplace')).toBe('Microsoft Marketplace');
    expect(getSourceLabel('Manual')).toBe('Granted by XSProject');
    expect(getSourceLabel('Included')).toContain('Included');
    expect(getSourceLabel(null)).toBeUndefined();
  });
});

describe('expiry', () => {
  it('flags an active service whose validTo is within 14 days', () => {
    expect(isExpiringSoon(status({ validTo: daysFromNow(10) }), NOW)).toBe(true);
    expect(isExpiringSoon(status({ validTo: daysFromNow(20) }), NOW)).toBe(false);
    expect(isExpiringSoon(status({ validTo: daysFromNow(-1) }), NOW)).toBe(false);
  });

  it('flags a term end only when the subscription will not renew', () => {
    expect(isExpiringSoon(status({ termEnd: daysFromNow(5), autoRenew: false }), NOW)).toBe(true);
    expect(isExpiringSoon(status({ termEnd: daysFromNow(5), autoRenew: null }), NOW)).toBe(true);
    expect(isExpiringSoon(status({ termEnd: daysFromNow(5), autoRenew: true }), NOW)).toBe(false);
  });

  it('prefers validTo over termEnd', () => {
    const end = getServiceEndDate(status({ validTo: daysFromNow(3), termEnd: daysFromNow(9) }));
    expect(end?.toISOString()).toBe(daysFromNow(3));
  });

  it('does not flag a service that is not active', () => {
    expect(isExpiringSoon(status({ status: 'Suspended', validTo: daysFromNow(3) }), NOW)).toBe(false);
  });
});

describe('getServiceStatusMessage', () => {
  it('sends a suspended service to the Microsoft 365 admin center', () => {
    expect(getServiceStatusMessage(status({ status: 'Suspended' }), NOW)).toContain('payment problem');
    expect(getServiceStatusMessage(status({ status: 'Suspended' }), NOW)).toContain('Microsoft 365 admin center');
  });

  it('explains inactive, unavailable and pending services', () => {
    expect(getServiceStatusMessage(status({ status: 'Inactive' }), NOW)).toContain('cancelled or has ended');
    expect(getServiceStatusMessage(status({ status: 'Unavailable' }), NOW)).toContain('currently unavailable');
    expect(getServiceStatusMessage(status({ status: 'Pending' }), NOW)).toContain('not activated yet');
  });

  it('warns about an expiring service and is silent about a healthy one', () => {
    expect(getServiceStatusMessage(status({ validTo: daysFromNow(7) }), NOW)).toContain('Tender ends on');
    expect(getServiceStatusMessage(status({}), NOW)).toBeUndefined();
    expect(getServiceStatusMessage(status({ status: 'NotPurchased' }), NOW)).toBeUndefined();
  });
});

describe('mergeServiceRows', () => {
  const catalog: ISaasCatalog = {
    services: [
      {
        serviceKey: 'TENDER',
        name: 'Tender',
        description: 'Tenders',
        displayOrder: 2,
        isAvailable: true,
        hasOffer: true,
        features: [],
        plans: [{ offerId: 'o', planId: 'p', name: 'Monthly', pricingModel: 'Flat', billingTerm: 'P1M' }]
      },
      {
        serviceKey: 'KYP',
        name: 'Know Your Property',
        description: null,
        displayOrder: 1,
        isAvailable: false,
        hasOffer: false,
        features: [],
        plans: []
      },
      {
        serviceKey: 'MINUTES',
        name: 'Minutes of Meeting',
        description: null,
        displayOrder: 3,
        isAvailable: true,
        hasOffer: true,
        features: [],
        plans: []
      }
    ]
  };

  it('orders catalog services and merges their status', () => {
    const rows = mergeServiceRows(
      catalog,
      entitlements({ details: [status({ canPurchase: true, termEnd: '2026-12-01T00:00:00Z' })] })
    );

    expect(rows.map((row) => row.serviceKey)).toEqual(['KYP', 'TENDER', 'MINUTES']);
    expect(rows[1].status).toBe('Active');
    expect(rows[1].canPurchase).toBe(true);
    expect(rows[1].termEnd).toBe('2026-12-01T00:00:00Z');
    expect(rows[1].plans).toHaveLength(1);
  });

  it('treats a catalog service without a status as not purchasable', () => {
    const rows = mergeServiceRows(catalog, entitlements({ details: [] }));

    expect(rows[0].status).toBe('Unavailable');
    expect(rows[2].status).toBe('NotPurchased');
    expect(rows.every((row) => !row.canPurchase)).toBe(true);
  });

  it('still lists a service the catalog does not know', () => {
    const rows = mergeServiceRows(undefined, entitlements({ details: [status({ serviceKey: 'ED_CONTROL', name: 'ED Control' })] }));

    expect(rows).toHaveLength(1);
    expect(rows[0].name).toBe('ED Control');
  });
});

describe('buildPurchaseReturnUrl', () => {
  it('replaces the fragment with the Subscription return route', () => {
    expect(
      buildPurchaseReturnUrl('https://contoso.sharepoint.com/sites/x/SitePages/App.aspx?env=1#/projects', 'TENDER')
    ).toBe('https://contoso.sharepoint.com/sites/x/SitePages/App.aspx?env=1#/subscription?saas=returned&service=TENDER');
  });

  it('works without a fragment and encodes the key', () => {
    expect(buildPurchaseReturnUrl('https://contoso.sharepoint.com/', 'ED CONTROL')).toBe(
      'https://contoso.sharepoint.com/#/subscription?saas=returned&service=ED%20CONTROL'
    );
  });
});

describe('parseSaasReturnQuery', () => {
  it('reads the return the landing page redirects to, with the subscription to activate', () => {
    expect(parseSaasReturnQuery(`?saas=returned&service=TENDER&subscriptionId=${SUBSCRIPTION_ID}`)).toEqual({
      isPurchaseReturn: true,
      serviceKey: 'TENDER',
      subscriptionId: SUBSCRIPTION_ID
    });
  });

  it('reads a return without a purchase, and one without a service', () => {
    expect(parseSaasReturnQuery('?saas=returned&service=TENDER')).toEqual({
      isPurchaseReturn: true,
      serviceKey: 'TENDER',
      subscriptionId: undefined
    });
    // A purchase not started from SPFx names no service.
    expect(parseSaasReturnQuery(`?saas=returned&subscriptionId=${SUBSCRIPTION_ID}`)).toEqual({
      isPurchaseReturn: true,
      serviceKey: undefined,
      subscriptionId: SUBSCRIPTION_ID
    });
  });

  it('ignores a subscriptionId that is not a GUID, or that is not on a purchase return', () => {
    expect(parseSaasReturnQuery('?saas=returned&subscriptionId=not-a-guid').subscriptionId).toBeUndefined();
    expect(parseSaasReturnQuery(`?saas=returned&subscriptionId=${SUBSCRIPTION_ID}x`).subscriptionId).toBeUndefined();
    expect(parseSaasReturnQuery(`?subscriptionId=${SUBSCRIPTION_ID}`)).toEqual({
      isPurchaseReturn: false,
      serviceKey: undefined,
      subscriptionId: undefined
    });
  });

  it('keeps the first value of each parameter', () => {
    const other: string = '11111111-2222-3333-4444-555555555555';
    expect(
      parseSaasReturnQuery(`?saas=returned&service=TENDER&subscriptionId=${SUBSCRIPTION_ID}&service=KYP&subscriptionId=${other}`)
    ).toEqual({ isPurchaseReturn: true, serviceKey: 'TENDER', subscriptionId: SUBSCRIPTION_ID });
  });

  it('reads a service on its own, and ignores anything else', () => {
    const nothing = { isPurchaseReturn: false, serviceKey: undefined, subscriptionId: undefined };

    expect(parseSaasReturnQuery('?service=KYP')).toEqual({ ...nothing, serviceKey: 'KYP' });
    expect(parseSaasReturnQuery('?saas=activated&x=%E0%A4%A')).toEqual(nothing);
    expect(parseSaasReturnQuery('')).toEqual(nothing);
    expect(parseSaasReturnQuery(undefined)).toEqual(nothing);
  });
});

describe('isSaasReturnFragment', () => {
  it('recognises the return the landing page redirects to', () => {
    expect(isSaasReturnFragment(`#/subscription?saas=returned&service=TENDER&subscriptionId=${SUBSCRIPTION_ID}`)).toBe(true);
    expect(isSaasReturnFragment('#/subscription?saas=returned&service=TENDER')).toBe(true);
  });

  it('ignores every other fragment', () => {
    expect(isSaasReturnFragment('#/subscription?service=TENDER')).toBe(false);
    expect(isSaasReturnFragment('#/subscription')).toBe(false);
    expect(isSaasReturnFragment('#/projects?saas=returned')).toBe(false);
    expect(isSaasReturnFragment('')).toBe(false);
    expect(isSaasReturnFragment(undefined)).toBe(false);
  });
});

describe('toSaasSubscriptions', () => {
  it('defaults missing fields and drops rows without a subscription id', () => {
    const subscriptions = toSaasSubscriptions([
      { marketplaceSubscriptionId: SUBSCRIPTION_ID, status: 'PendingFulfillmentStart', services: ['TENDER', null] },
      { subscriptionName: 'nameless' },
      'junk'
    ]);

    expect(subscriptions).toEqual([
      {
        marketplaceSubscriptionId: SUBSCRIPTION_ID,
        subscriptionName: null,
        offerId: '',
        planId: '',
        planName: null,
        quantity: null,
        status: 'PendingFulfillmentStart',
        services: ['TENDER'],
        termStart: null,
        termEnd: null,
        termUnit: null,
        autoRenew: null,
        isFreeTrial: false
      }
    ]);
  });

  it('rejects a body that is not an array', () => {
    expect(toSaasSubscriptions({ subscriptions: [] })).toBeUndefined();
    expect(toSaasSubscriptions('<html>')).toBeUndefined();
  });
});

describe('toSaasActivationResult', () => {
  const wire = {
    state: 'Activated',
    subscription: { marketplaceSubscriptionId: SUBSCRIPTION_ID, planId: 'monthly', status: 'Subscribed', services: ['TENDER'] },
    services: ['TENDER'],
    entitlements: entitlements(),
    message: 'Tender is active.',
    correlationId: 'corr-1'
  };

  it('reads a complete answer', () => {
    const result = toSaasActivationResult(JSON.parse(JSON.stringify(wire)));

    expect(result?.state).toBe('Activated');
    expect(result?.subscription.marketplaceSubscriptionId).toBe(SUBSCRIPTION_ID);
    expect(result?.subscription.status).toBe('Subscribed');
    expect(result?.services).toEqual(['TENDER']);
    expect(result?.entitlements).toEqual(entitlements());
    expect(result?.message).toBe('Tender is active.');
    expect(result?.correlationId).toBe('corr-1');
  });

  it('defaults the optional fields', () => {
    const result = toSaasActivationResult({ ...wire, services: undefined, message: null, correlationId: undefined });

    expect(result?.services).toEqual([]);
    expect(result?.message).toBe('');
    expect(result?.correlationId).toBeNull();
  });

  it('rejects an answer without a state, a subscription id or entitlements', () => {
    expect(toSaasActivationResult({ ...wire, state: undefined })).toBeUndefined();
    expect(toSaasActivationResult({ ...wire, subscription: { status: 'Subscribed' } })).toBeUndefined();
    expect(toSaasActivationResult({ ...wire, subscription: null })).toBeUndefined();
    expect(toSaasActivationResult({ ...wire, entitlements: 'none' })).toBeUndefined();
    expect(toSaasActivationResult('<html>')).toBeUndefined();
  });
});

describe('isActivationSuccessful', () => {
  it.each([
    ['Activated', true],
    ['AlreadyActive', true],
    ['Suspended', false],
    ['Unsubscribed', false],
    ['Frozen', false]
  ])('treats %s as %s', (state: string, expected: boolean) => {
    expect(isActivationSuccessful({ state })).toBe(expected);
  });
});

describe('subscription presentation', () => {
  it.each([
    ['PendingFulfillmentStart', 'Awaiting activation', 'warning'],
    ['Subscribed', 'Active', 'success'],
    ['Suspended', 'Suspended', 'error'],
    ['Unsubscribed', 'Cancelled', 'neutral'],
    ['NotStarted', 'Not started', 'neutral']
  ])('labels %s as "%s" in the %s tone', (value: string, label: string, tone: string) => {
    expect(getSubscriptionStatusLabel(value)).toBe(label);
    expect(getSubscriptionStatusTone(value)).toBe(tone);
  });

  it('shows an unknown status as the server wrote it', () => {
    expect(getSubscriptionStatusLabel('Frozen')).toBe('Frozen');
    expect(getSubscriptionStatusTone('Frozen')).toBe('neutral');
  });
});

describe('canActivateSubscription', () => {
  it('offers Activate only to a tenant admin, only for a subscription awaiting activation', () => {
    const admin = { isTenantAdmin: true };

    expect(canActivateSubscription({ status: 'PendingFulfillmentStart' }, admin)).toBe(true);
    expect(canActivateSubscription({ status: 'PendingFulfillmentStart' }, { isTenantAdmin: false })).toBe(false);
    expect(canActivateSubscription({ status: 'PendingFulfillmentStart' }, undefined)).toBe(false);
    expect(canActivateSubscription({ status: 'Subscribed' }, admin)).toBe(false);
    expect(canActivateSubscription({ status: 'Suspended' }, admin)).toBe(false);
  });
});

describe('isSafeOfferUrl', () => {
  it('accepts https only', () => {
    expect(isSafeOfferUrl('https://appsource.microsoft.com/product/web-apps/x')).toBe(true);
    expect(isSafeOfferUrl('http://appsource.microsoft.com/')).toBe(false);
    expect(isSafeOfferUrl('data:text/html,offer')).toBe(false);
    expect(isSafeOfferUrl(undefined)).toBe(false);
  });
});

describe('getSaasErrorInfo', () => {
  it.each([
    ['SAAS_NOT_TENANT_ADMIN', 'Global Administrator'],
    ['SAAS_SERVICE_NOT_PURCHASABLE', 'cannot be bought'],
    ['SAAS_RETURN_URL_INVALID', 'not registered for your organisation'],
    ['SAAS_TENANT_UNKNOWN', 'not registered with XSProject'],
    ['SAAS_UNAVAILABLE', 'temporarily unavailable'],
    ['SAAS_NOT_CONFIGURED', 'not set up'],
    ['SAAS_VALIDATION', 'could not accept this request'],
    ['SAAS_WRONG_TENANT', 'belongs to another organisation'],
    ['SAAS_NOT_FOUND', 'Microsoft does not know this subscription'],
    ['SAAS_UNKNOWN_PLAN', 'plan bought for this subscription is not known'],
    ['SAAS_ACTIVATION_FAILED', 'could not be activated'],
    ['SAAS_MARKETPLACE_ERROR', 'Microsoft Marketplace did not answer']
  ])('maps %s to friendly text', (code: string, text: string) => {
    const info = getSaasErrorInfo(new ApiError('raw', { status: 400, code, correlationId: 'corr-1' }));

    expect(info.message).toContain(text);
    expect(info.code).toBe(code);
    expect(info.correlationId).toBe('corr-1');
  });

  it('explains an unreachable API, 401 and 403 without a code', () => {
    expect(getSaasErrorInfo(new ApiError('x', { status: 0 })).message).toContain('could not be reached');
    expect(getSaasErrorInfo(new ApiError('x', { status: 401 })).message).toContain('session has expired');
    expect(getSaasErrorInfo(new ApiError('x', { status: 403 })).message).toContain('not allowed');
  });

  it("falls back to the server's message, then to a generic one", () => {
    expect(getSaasErrorInfo(new ApiError('Bad thing', { status: 400 })).message).toBe('Bad thing');
    expect(getSaasErrorInfo('nope').message).toContain('Something went wrong');
  });

  it('tells a non-admin to ask an administrator to open the page', () => {
    const info = getSaasErrorInfo(new ApiError('raw', { status: 403, code: 'SAAS_NOT_TENANT_ADMIN' }));

    expect(info.message).toContain('activate');
    expect(info.message).toContain('ask one of them to open this page');
  });
});

describe('isRetryableSaasError', () => {
  it.each([
    ['SAAS_ACTIVATION_FAILED', 502],
    ['SAAS_MARKETPLACE_ERROR', 502],
    ['SAAS_UNAVAILABLE', 503]
  ])('offers a retry for %s', (code: string, status: number) => {
    expect(isRetryableSaasError(new ApiError('x', { status, code }))).toBe(true);
  });

  it.each([
    ['SAAS_VALIDATION', 400],
    ['SAAS_NOT_AUTHENTICATED', 401],
    ['SAAS_NOT_TENANT_ADMIN', 403],
    ['SAAS_WRONG_TENANT', 403],
    ['SAAS_TENANT_UNKNOWN', 404],
    ['SAAS_NOT_FOUND', 404],
    ['SAAS_UNKNOWN_PLAN', 409],
    ['SAAS_NOT_CONFIGURED', 503]
  ])('does not offer one for %s, which would fail the same way again', (code: string, status: number) => {
    expect(isRetryableSaasError(new ApiError('x', { status, code }))).toBe(false);
  });

  it('decides by status without a code, and retries a failure that is not an ApiError', () => {
    expect(isRetryableSaasError(new ApiError('x', { status: 0 }))).toBe(true);
    expect(isRetryableSaasError(new ApiError('x', { status: 500 }))).toBe(true);
    expect(isRetryableSaasError(new ApiError('x', { status: 400 }))).toBe(false);
    expect(isRetryableSaasError(new Error('timeout'))).toBe(true);
  });
});
