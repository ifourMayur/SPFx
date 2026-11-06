/**
 * Render tests for `FeatureGate`: each of its states, against a stubbed `ISaasService`.
 *
 * Every `ReactDOM.render` here is unmounted in `afterEach`.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import FeatureGate from './FeatureGate';
import { ApiError } from '../../../../models/ApiError';
import { ISaasEntitlements, ISaasServiceStatus } from '../../../../models/Saas';
import { IGetEntitlementsOptions, ISaasService } from '../../../../services/SaasService';

function detail(status: string): ISaasServiceStatus {
  return {
    serviceKey: 'TENDER',
    name: 'Tender Management',
    description: null,
    status,
    source: null,
    validTo: null,
    termEnd: null,
    isFreeTrial: false,
    autoRenew: null,
    canPurchase: status === 'NotPurchased',
    marketplaceSubscriptionId: null
  };
}

function entitlements(features: string[], isTenantAdmin: boolean, status: string = 'NotPurchased'): ISaasEntitlements {
  return {
    tenantKnown: true,
    entraTenantId: 't',
    isTenantAdmin,
    isBillingEnforced: true,
    services: [],
    features,
    details: [detail(status)],
    cacheSeconds: 300,
    evaluatedOn: ''
  };
}

/** A service whose answers are queued: each call takes the next one (the last repeats). */
function stubService(answers: Array<ISaasEntitlements | Error>, calls?: IGetEntitlementsOptions[]): ISaasService {
  return {
    getEntitlements: async (options?: IGetEntitlementsOptions): Promise<ISaasEntitlements> => {
      if (calls) {
        calls.push(options || {});
      }
      const answer: ISaasEntitlements | Error = answers.length > 1 ? (answers.shift() as ISaasEntitlements | Error) : answers[0];
      if (answer instanceof Error) {
        throw answer;
      }
      return answer;
    },
    getCatalog: async () => ({ services: [] }),
    getSubscriptions: async () => [],
    createPurchaseIntent: async () => {
      throw new Error('not expected');
    },
    activateSubscription: async () => {
      throw new Error('not expected');
    },
    invalidate: (): void => undefined
  };
}

describe('FeatureGate', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  /** The one render call, so it pairs with the unmount in `afterEach`. */
  function mount(service: ISaasService, onEnableService?: (serviceKey: string) => void, hideWhenLocked?: boolean): void {
    ReactDOM.render(
      <FeatureGate
        saasService={service}
        featureKey="TENDER.CORE"
        serviceKey="TENDER"
        serviceName="Tender"
        onEnableService={onEnableService}
        hideWhenLocked={hideWhenLocked}
      >
        <div id="paid">Paid feature</div>
      </FeatureGate>,
      host
    );
  }

  async function render(
    service: ISaasService,
    onEnableService?: (serviceKey: string) => void,
    hideWhenLocked?: boolean
  ): Promise<void> {
    await act(async () => {
      mount(service, onEnableService, hideWhenLocked);
    });
  }

  async function click(text: string): Promise<void> {
    const button: HTMLButtonElement = Array.prototype.slice
      .call(host.querySelectorAll('button'))
      .filter((candidate: HTMLButtonElement) => (candidate.textContent || '').trim() === text)[0];

    await act(async () => {
      button.click();
    });
  }

  it('shows a checking state, and not the feature, while loading', () => {
    const never: ISaasService = {
      ...stubService([entitlements([], false)]),
      getEntitlements: () => new Promise<ISaasEntitlements>(() => undefined)
    };

    act(() => {
      mount(never);
    });

    expect(host.textContent).toContain('Checking your subscription');
    expect(host.querySelector('#paid')).toBeNull();
  });

  it('renders the children when the feature is enabled', async () => {
    await render(stubService([entitlements(['TENDER.CORE'], false, 'Active')]));

    expect(host.querySelector('#paid')).not.toBeNull();
    expect(host.querySelector('[data-feature-gate]')).toBeNull();
  });

  it('tells a non-admin to contact their administrator, using the server name for the service', async () => {
    const opened: string[] = [];
    await render(stubService([entitlements([], false)]), (key: string) => opened.push(key));

    expect(host.querySelector('#paid')).toBeNull();
    expect(host.textContent).toContain('Tender Management is not enabled for your organisation.');
    expect(host.textContent).toContain('Contact your administrator to enable Tender Management.');
    expect(host.querySelector('button')).toBeNull();
  });

  it('offers an admin "Enable" that opens the subscription page for the service', async () => {
    const opened: string[] = [];
    await render(stubService([entitlements([], true)]), (key: string) => opened.push(key));

    expect(host.textContent).not.toContain('Contact your administrator');
    await click('Enable Tender Management');

    expect(opened).toEqual(['TENDER']);
  });

  it('explains a suspended service', async () => {
    await render(stubService([entitlements([], true, 'Suspended')]), () => undefined);

    expect(host.textContent).toContain('payment problem');
    expect(host.querySelector('#paid')).toBeNull();
  });

  it('offers nothing to buy for a service that is unavailable to everyone', async () => {
    await render(stubService([entitlements([], true, 'Unavailable')]), () => undefined);

    expect(host.textContent).toContain('currently unavailable');
    expect(host.querySelector('button')).toBeNull();
  });

  describe('hideWhenLocked', () => {
    it('renders nothing while loading', () => {
      const never: ISaasService = {
        ...stubService([entitlements([], true)]),
        getEntitlements: () => new Promise<ISaasEntitlements>(() => undefined)
      };

      act(() => {
        mount(never, undefined, true);
      });

      expect(host.innerHTML).toBe('');
    });

    it('renders nothing when locked, not even "Enable" for an admin', async () => {
      await render(stubService([entitlements([], true)]), () => undefined, true);

      expect(host.innerHTML).toBe('');
    });

    it('still renders the children when the feature is enabled', async () => {
      await render(stubService([entitlements(['TENDER.CORE'], false, 'Active')]), undefined, true);

      expect(host.querySelector('#paid')).not.toBeNull();
    });

    it('still says so, with a Retry, when the check fails', async () => {
      await render(stubService([new ApiError('Service Unavailable', { status: 503 })]), undefined, true);

      expect(host.textContent).toContain("We couldn't check your subscription, so Tender is hidden for now.");
      expect(host.textContent).toContain('Retry');
    });
  });

  it('fails closed with the reason and reference when the check fails, and recovers on Retry', async () => {
    const calls: IGetEntitlementsOptions[] = [];
    const failure: ApiError = new ApiError('Service Unavailable', {
      status: 503,
      code: 'SAAS_UNAVAILABLE',
      correlationId: 'corr-42'
    });
    await render(stubService([failure, entitlements(['TENDER.CORE'], false, 'Active')], calls));

    expect(host.querySelector('#paid')).toBeNull();
    expect(host.textContent).toContain("We couldn't check your subscription, so Tender is hidden for now.");
    expect(host.textContent).toContain('temporarily unavailable');
    expect(host.textContent).toContain('Reference: corr-42');

    await click('Retry');

    expect(calls[1]).toEqual({ refresh: true });
    expect(host.querySelector('#paid')).not.toBeNull();
  });
});
