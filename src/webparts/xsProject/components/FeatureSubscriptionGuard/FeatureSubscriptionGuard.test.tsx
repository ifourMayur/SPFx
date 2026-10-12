/**
 * Behaviour tests for `FeatureSubscriptionGuard`: the real component tree - the guard,
 * `SubscriptionPrompt` and `SubscriptionStatus` - over a stubbed `ISubscriptionService`.
 *
 * What they pin down is that only the service's answer opens the gate: not a click on
 * Subscribe, not a return from Marketplace, and not another feature's subscription.
 *
 * Every `ReactDOM.render` here is unmounted in `afterEach`.
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import FeatureSubscriptionGuard from './FeatureSubscriptionGuard';
import type { FeatureCode, IFeatureSubscription, ISubscriptionLinks } from '../../../../models/FeatureSubscription';
import type {
  IFeatureSubscriptionRequestOptions,
  ISubscriptionService
} from '../../../../services/SubscriptionService';

const MARKETPLACE_URL: string = 'https://marketplace.microsoft.com/product/saas/contoso.xsproject';
const MANAGE_URL: string = 'https://admin.microsoft.com/#/subscriptions';

const LINKS: ISubscriptionLinks = { marketplaceOfferUrl: MARKETPLACE_URL, manageSubscriptionUrl: MANAGE_URL };

function subscription(
  feature: FeatureCode,
  state: IFeatureSubscription['state'],
  extra?: Partial<IFeatureSubscription>
): IFeatureSubscription {
  const status: { [key: string]: string } = {
    active: 'Active',
    pending: 'Pending',
    suspended: 'Suspended',
    inactive: 'Inactive'
  };

  return {
    feature,
    state,
    isActive: state === 'active',
    isEnabled: state === 'active',
    subscriptionStatus: status[state],
    ...extra
  };
}

interface ICall {
  feature: FeatureCode;
  refresh: boolean;
}

/**
 * A service answering from `answer`, recording every call. `answer` may throw or return a
 * promise that never settles, which is how the loading and failure states are reached.
 */
function stubService(
  answer: (feature: FeatureCode, call: number) => IFeatureSubscription | Promise<IFeatureSubscription>,
  calls: ICall[],
  links: ISubscriptionLinks = LINKS
): ISubscriptionService {
  return {
    links,
    getSubscriptionStatus: async (): Promise<IFeatureSubscription[]> => [],
    getFeatureSubscription: async (
      feature: FeatureCode,
      options?: IFeatureSubscriptionRequestOptions
    ): Promise<IFeatureSubscription> => {
      calls.push({ feature, refresh: !!options?.refresh });
      return answer(feature, calls.length);
    },
    setFeatureEnabled: async (): Promise<IFeatureSubscription> => {
      throw new Error('The guard never switches a feature.');
    }
  };
}

describe('FeatureSubscriptionGuard', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    // jsdom does not navigate, and says so on the console; a link's own default action is
    // the browser's business, not what these tests are about.
    host.addEventListener('click', (event: Event) => event.preventDefault());
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  /** The one `ReactDOM.render` of this file, paired with the unmount in `afterEach`. */
  async function mount(element: React.ReactElement): Promise<void> {
    await act(async () => {
      ReactDOM.render(element, host);
    });
  }

  async function render(
    service: ISubscriptionService,
    feature: FeatureCode = 'ED_CONTROL',
    isReturningFromMarketplace: boolean = false
  ): Promise<void> {
    await mount(
      <FeatureSubscriptionGuard
        feature={feature}
        subscriptionService={service}
        isReturningFromMarketplace={isReturningFromMarketplace}
      >
        <div id="feature-content">ED Control workspace</div>
      </FeatureSubscriptionGuard>
    );
  }

  function subscribeLink(): HTMLAnchorElement | null {
    return Array.prototype.slice
      .call(host.querySelectorAll('a'))
      .filter((link: HTMLAnchorElement) => link.textContent === 'Subscribe')[0] || null;
  }

  function button(text: string): HTMLButtonElement | undefined {
    return Array.prototype.slice
      .call(host.querySelectorAll('button'))
      .filter((item: HTMLButtonElement) => item.textContent === text)[0];
  }

  async function click(element: HTMLElement): Promise<void> {
    await act(async () => {
      element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    });
  }

  it('says it is checking, and offers nothing, until the API answers', async () => {
    await render(stubService(() => new Promise<IFeatureSubscription>(() => undefined), []));

    expect(host.textContent).toContain('Checking ED Control subscription');
    expect(subscribeLink()).toBeNull();
    expect(host.querySelector('#feature-content')).toBeNull();
  });

  it('opens the feature when the API reports it active', async () => {
    await render(
      stubService(() => subscription('ED_CONTROL', 'active', { planId: 'EDCONTROL_PRO', planName: 'Professional' }), [])
    );

    expect(host.querySelector('#feature-content')).not.toBeNull();
    expect(subscribeLink()).toBeNull();
    expect(host.textContent).toContain('Plan: Professional');
    expect(host.textContent).toContain('Status: Active');
    expect((host.querySelector('a[href="' + MANAGE_URL + '"]') as HTMLAnchorElement).textContent).toBe(
      'Manage subscription'
    );
  });

  it('locks the feature and offers Subscribe when the API reports it inactive', async () => {
    await render(stubService(() => subscription('ED_CONTROL', 'inactive'), []));

    expect(host.querySelector('#feature-content')).toBeNull();
    expect(host.textContent).toContain('This feature requires an active subscription.');

    const link: HTMLAnchorElement | null = subscribeLink();
    expect(link?.href).toBe(MARKETPLACE_URL);
    expect(link?.target).toBe('_blank');
    expect(link?.rel).toBe('noopener noreferrer');
  });

  it('treats Subscribe as a trip to Marketplace, not as a purchase', async () => {
    const calls: ICall[] = [];
    await render(stubService(() => subscription('ED_CONTROL', 'inactive'), calls));

    await click(subscribeLink() as HTMLAnchorElement);

    // Still locked: clicking is not paying.
    expect(host.querySelector('#feature-content')).toBeNull();
    expect(host.textContent).toContain('your subscription is being processed');
    expect(button('Check again')).toBeDefined();
    expect(calls).toEqual([{ feature: 'ED_CONTROL', refresh: false }]);
  });

  it('re-checks, past the cache, when the user switches back from Marketplace', async () => {
    const calls: ICall[] = [];
    await render(
      stubService(
        (feature: FeatureCode, call: number) => subscription(feature, call === 1 ? 'inactive' : 'active'),
        calls
      )
    );

    await click(subscribeLink() as HTMLAnchorElement);
    await act(async () => {
      document.dispatchEvent(new Event('visibilitychange'));
    });

    expect(calls[1]).toEqual({ feature: 'ED_CONTROL', refresh: true });
    expect(host.querySelector('#feature-content')).not.toBeNull();
  });

  it('does not re-check on a tab switch the user did not make for Marketplace', async () => {
    const calls: ICall[] = [];
    await render(stubService(() => subscription('ED_CONTROL', 'inactive'), calls));

    await act(async () => {
      document.dispatchEvent(new Event('visibilitychange'));
    });

    expect(calls.length).toBe(1);
  });

  it('says a pending subscription is being processed, without offering a second purchase', async () => {
    await render(stubService(() => subscription('ED_CONTROL', 'pending'), []));

    expect(host.textContent).toContain('Your subscription is being processed.');
    expect(host.textContent).toContain('Please try again in a few moments.');
    expect(subscribeLink()).toBeNull();
    expect(host.querySelector('#feature-content')).toBeNull();
  });

  it('opens the feature once a pending subscription turns active on Check again', async () => {
    const calls: ICall[] = [];
    await render(
      stubService(
        (feature: FeatureCode, call: number) => subscription(feature, call === 1 ? 'pending' : 'active'),
        calls
      )
    );

    await click(button('Check again') as HTMLButtonElement);

    expect(calls[1].refresh).toBe(true);
    expect(host.querySelector('#feature-content')).not.toBeNull();
  });

  it('sends a suspended subscription to Manage subscription, not to Subscribe', async () => {
    await render(stubService(() => subscription('ED_CONTROL', 'suspended'), []));

    expect(host.textContent).toContain('subscription is suspended');
    expect(subscribeLink()).toBeNull();
    expect((host.querySelector('a[href="' + MANAGE_URL + '"]') as HTMLAnchorElement).target).toBe('_blank');
  });

  it('reports a failed check as unverified - not as inactive - and retries', async () => {
    const calls: ICall[] = [];
    await render(
      stubService((feature: FeatureCode, call: number) => {
        if (call === 1) {
          throw new Error('Network down');
        }

        return subscription(feature, 'active');
      }, calls)
    );

    expect(host.querySelector('[role="alert"]')?.textContent).toContain('Unable to verify your subscription.');
    expect(subscribeLink()).toBeNull();
    expect(host.querySelector('#feature-content')).toBeNull();

    await click(button('Try again') as HTMLButtonElement);

    expect(calls[1]).toEqual({ feature: 'ED_CONTROL', refresh: true });
    expect(host.querySelector('#feature-content')).not.toBeNull();
  });

  it('reads past the cache when the user arrives back from Marketplace', async () => {
    const calls: ICall[] = [];
    await render(stubService(() => subscription('ED_CONTROL', 'inactive'), calls), 'ED_CONTROL', true);

    expect(calls).toEqual([{ feature: 'ED_CONTROL', refresh: true }]);
    // The return marker changes the wording, never the access.
    expect(host.querySelector('#feature-content')).toBeNull();
    expect(host.textContent).toContain('your subscription is being processed');
  });

  it('keeps a paid-for feature locked while it is switched off, without offering to buy it again', async () => {
    await render(
      stubService(() => subscription('ED_CONTROL', 'active', { isActive: false, isEnabled: false }), [])
    );

    expect(host.querySelector('#feature-content')).toBeNull();
    expect(host.textContent).toContain('ED Control is turned off.');
    expect(subscribeLink()).toBeNull();
    expect((host.querySelector('a[href="' + MANAGE_URL + '"]') as HTMLAnchorElement).textContent).toBe(
      'Manage subscription'
    );
  });

  it('offers no Subscribe link where Marketplace is not configured', async () => {
    await render(
      stubService(() => subscription('ED_CONTROL', 'inactive'), [], { marketplaceOfferUrl: '', manageSubscriptionUrl: '' })
    );

    expect(subscribeLink()).toBeNull();
    expect(host.textContent).toContain('Subscribing is not available here yet.');
  });

  it('keeps an active ED Control from unlocking Minutes', async () => {
    const calls: ICall[] = [];
    const service: ISubscriptionService = stubService(
      (feature: FeatureCode) => subscription(feature, feature === 'ED_CONTROL' ? 'active' : 'inactive'),
      calls
    );

    await mount(
      <>
        <FeatureSubscriptionGuard feature="ED_CONTROL" subscriptionService={service}>
          <div id="ed-control">ED Control</div>
        </FeatureSubscriptionGuard>
        <FeatureSubscriptionGuard feature="MINUTES" subscriptionService={service}>
          <div id="minutes">Minutes</div>
        </FeatureSubscriptionGuard>
      </>
    );

    expect(host.querySelector('#ed-control')).not.toBeNull();
    expect(host.querySelector('#minutes')).toBeNull();
    expect(subscribeLink()).not.toBeNull();
    expect(calls.map((call: ICall) => call.feature).sort()).toEqual(['ED_CONTROL', 'MINUTES']);
  });

  it('asks again when the guarded feature changes', async () => {
    const calls: ICall[] = [];
    const service: ISubscriptionService = stubService(
      (feature: FeatureCode) => subscription(feature, feature === 'ED_CONTROL' ? 'active' : 'inactive'),
      calls
    );

    await render(service, 'ED_CONTROL');
    expect(host.querySelector('#feature-content')).not.toBeNull();

    await render(service, 'MINUTES');
    expect(host.querySelector('#feature-content')).toBeNull();
    expect(calls.map((call: ICall) => call.feature)).toEqual(['ED_CONTROL', 'MINUTES']);
  });
});
