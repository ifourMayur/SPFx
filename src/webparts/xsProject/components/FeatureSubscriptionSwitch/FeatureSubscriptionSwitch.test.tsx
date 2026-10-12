/**
 * Behaviour tests for `FeatureSubscriptionSwitch`: the real component tree over a stubbed
 * `ISubscriptionService`, with `window.open` spied on.
 *
 * What they pin down: switching on records the choice and opens Marketplace only when the
 * feature has not been bought; switching on never makes the feature usable by itself; and
 * a failed save puts the switch back.
 *
 * Every `ReactDOM.render` here is unmounted in `afterEach`.
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import FeatureSubscriptionSwitch from './FeatureSubscriptionSwitch';
import type { FeatureCode, IFeatureSubscription, ISubscriptionLinks } from '../../../../models/FeatureSubscription';
import type {
  IFeatureSubscriptionRequestOptions,
  ISubscriptionService
} from '../../../../services/SubscriptionService';

const MARKETPLACE_URL: string = 'https://marketplace.microsoft.com/';
const MANAGE_URL: string = 'https://portal.azure.com/';

const LINKS: ISubscriptionLinks = { marketplaceOfferUrl: MARKETPLACE_URL, manageSubscriptionUrl: MANAGE_URL };

/** Never bought, switched off - where every tenant starts. */
const NOT_BOUGHT: IFeatureSubscription = {
  feature: 'ED_CONTROL',
  state: 'inactive',
  isActive: false,
  isEnabled: false,
  subscriptionStatus: 'Inactive'
};

/** Bought and switched on. */
const ACTIVE: IFeatureSubscription = {
  feature: 'ED_CONTROL',
  state: 'active',
  isActive: true,
  isEnabled: true,
  subscriptionStatus: 'Active',
  planName: 'Professional'
};

interface ICalls {
  reads: { feature: FeatureCode; refresh: boolean }[];
  writes: { feature: FeatureCode; isEnabled: boolean }[];
}

/**
 * A service holding one feature's state. Switching it records the switch and nothing more,
 * as the real API does; `paid` stands in for Marketplace reporting the purchase.
 */
function stubService(
  initial: IFeatureSubscription,
  calls: ICalls,
  options: { failWrites?: boolean; failReads?: boolean; links?: ISubscriptionLinks } = {}
): ISubscriptionService & { pay: () => void } {
  let current: IFeatureSubscription = initial;

  const settle = (isEnabled: boolean): IFeatureSubscription => {
    const isPaid: boolean = current.state === 'active';
    current = { ...current, isEnabled, isActive: isPaid && isEnabled };
    return current;
  };

  return {
    links: options.links || LINKS,
    getSubscriptionStatus: async (): Promise<IFeatureSubscription[]> => [current],
    getFeatureSubscription: async (
      feature: FeatureCode,
      requestOptions?: IFeatureSubscriptionRequestOptions
    ): Promise<IFeatureSubscription> => {
      calls.reads.push({ feature, refresh: !!requestOptions?.refresh });

      if (options.failReads) {
        throw new Error('Network down');
      }

      return current;
    },
    setFeatureEnabled: async (feature: FeatureCode, isEnabled: boolean): Promise<IFeatureSubscription> => {
      calls.writes.push({ feature, isEnabled });

      if (options.failWrites) {
        throw new Error('You do not have permission to change this feature.');
      }

      return settle(isEnabled);
    },
    pay: (): void => {
      current = { ...current, state: 'active', subscriptionStatus: 'Active', isActive: current.isEnabled };
    }
  };
}

function newCalls(): ICalls {
  return { reads: [], writes: [] };
}

describe('FeatureSubscriptionSwitch', () => {
  let host: HTMLDivElement;
  let openSpy: jest.SpyInstance;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    // jsdom does not navigate; a link's own default action is not what these tests are about.
    host.addEventListener('click', (event: Event) => event.preventDefault());
    openSpy = jest.spyOn(window, 'open').mockImplementation(() => null);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
    openSpy.mockRestore();
  });

  /** The one `ReactDOM.render` of this file, paired with the unmount in `afterEach`. */
  async function render(
    service: ISubscriptionService,
    canManage: boolean = true,
    isReturningFromMarketplace: boolean = false
  ): Promise<void> {
    await act(async () => {
      ReactDOM.render(
        <FeatureSubscriptionSwitch
          feature="ED_CONTROL"
          subscriptionService={service}
          canManage={canManage}
          isReturningFromMarketplace={isReturningFromMarketplace}
        />,
        host
      );
    });
  }

  function toggle(): HTMLInputElement | null {
    return host.querySelector('#feature-switch-ED_CONTROL');
  }

  async function flip(): Promise<void> {
    await act(async () => {
      (toggle() as HTMLInputElement).click();
    });
  }

  it('offers no switch until the API answers', async () => {
    await render({
      ...stubService(NOT_BOUGHT, newCalls()),
      getFeatureSubscription: () => new Promise<IFeatureSubscription>(() => undefined)
    });

    expect(host.textContent).toContain('Checking ED Control subscription');
    expect(toggle()).toBeNull();
  });

  it('shows a feature never bought as switched off, with how to subscribe', async () => {
    await render(stubService(NOT_BOUGHT, newCalls()));

    expect(toggle()?.checked).toBe(false);
    expect(toggle()?.getAttribute('role')).toBe('switch');
    expect(host.textContent).toContain('Turn on ED Control to subscribe in Microsoft Marketplace.');
  });

  it('switching on records the choice and opens the Marketplace payment screen', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(NOT_BOUGHT, calls));

    await flip();

    expect(calls.writes).toEqual([{ feature: 'ED_CONTROL', isEnabled: true }]);
    expect(openSpy).toHaveBeenCalledWith(MARKETPLACE_URL, '_blank', 'noopener,noreferrer');
    expect(toggle()?.checked).toBe(true);
  });

  // The whole point: the switch is a choice, payment is Marketplace's.
  it('keeps the feature locked after switching on, until the purchase is reported', async () => {
    const calls: ICalls = newCalls();
    const service = stubService(NOT_BOUGHT, calls);
    await render(service);

    await flip();

    expect(host.textContent).not.toContain('Plan: Professional');
    expect(host.textContent).toContain('This feature requires an active subscription.');
    expect(host.textContent).toContain('your subscription is being processed');

    // Marketplace reports the purchase while the user is on the other tab...
    service.pay();
    await act(async () => {
      document.dispatchEvent(new Event('visibilitychange'));
    });

    // ...and switching back re-reads past the cache and finds it active.
    expect(calls.reads[calls.reads.length - 1]).toEqual({ feature: 'ED_CONTROL', refresh: true });
    expect(host.textContent).toContain('Status: Active');
  });

  it('switches a paid-for feature back on without sending the user to pay again', async () => {
    const calls: ICalls = newCalls();
    await render(stubService({ ...ACTIVE, isActive: false, isEnabled: false }, calls));

    expect(host.textContent).toContain('ED Control is turned off.');

    await flip();

    expect(openSpy).not.toHaveBeenCalled();
    expect(calls.writes).toEqual([{ feature: 'ED_CONTROL', isEnabled: true }]);
    expect(host.textContent).toContain('Plan: Professional');
  });

  it('does not send a pending purchase to Marketplace a second time', async () => {
    await render(stubService({ ...NOT_BOUGHT, state: 'pending', subscriptionStatus: 'Pending' }, newCalls()));

    await flip();

    expect(openSpy).not.toHaveBeenCalled();
    expect(host.textContent).toContain('Your subscription is being processed.');
  });

  it('switching off records it and locks the feature at once', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(ACTIVE, calls));

    expect(toggle()?.checked).toBe(true);
    expect(host.textContent).toContain('Plan: Professional');

    await flip();

    expect(calls.writes).toEqual([{ feature: 'ED_CONTROL', isEnabled: false }]);
    expect(openSpy).not.toHaveBeenCalled();
    expect(toggle()?.checked).toBe(false);
    expect(host.textContent).toContain('ED Control is turned off.');
    expect((host.querySelector('a[href="' + MANAGE_URL + '"]') as HTMLAnchorElement).textContent).toBe(
      'Manage subscription'
    );
  });

  it('puts the switch back and says so when the save fails', async () => {
    await render(stubService(ACTIVE, newCalls(), { failWrites: true }));

    await flip();

    expect(toggle()?.checked).toBe(true);
    expect(host.querySelector('[role="alert"]')?.textContent).toContain('Unable to update ED Control.');
  });

  it('shows the switch to a user who may not change it, disabled', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(NOT_BOUGHT, calls), false);

    expect(toggle()?.disabled).toBe(true);
    expect(host.textContent).toContain('You do not have permission to change this setting.');
  });

  it('switches on without opening anything where Marketplace is not configured', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(NOT_BOUGHT, calls, { links: { marketplaceOfferUrl: '', manageSubscriptionUrl: '' } }));

    expect(host.textContent).toContain('Subscribing is not available here yet.');

    await flip();

    expect(openSpy).not.toHaveBeenCalled();
    expect(calls.writes.length).toBe(1);
  });

  it('reports a failed read as unverified, offers no switch, and retries', async () => {
    const calls: ICalls = newCalls();
    const service = stubService(NOT_BOUGHT, calls, { failReads: true });
    await render(service);

    expect(host.querySelector('[role="alert"]')?.textContent).toContain('Unable to verify your subscription.');
    expect(toggle()).toBeNull();

    await act(async () => {
      Array.prototype.slice.call(host.querySelectorAll('button'))
        .filter((button: HTMLButtonElement) => button.textContent === 'Try again')[0]
        .click();
    });

    expect(calls.reads.map((read) => read.refresh)).toEqual([false, true]);
  });

  it('reads past the cache when the user arrives back from Marketplace', async () => {
    const calls: ICalls = newCalls();
    await render(stubService({ ...NOT_BOUGHT, isEnabled: true }, calls), true, true);

    expect(calls.reads).toEqual([{ feature: 'ED_CONTROL', refresh: true }]);
    // The return marker changes the wording, never the access.
    expect(host.textContent).not.toContain('Status: Active');
    expect(host.textContent).toContain('your subscription is being processed');
  });
});
