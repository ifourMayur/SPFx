/**
 * Behaviour tests for `FeatureSubscriptionSwitch`: the real component tree over a stubbed
 * `ISubscriptionService`, with `window.open` answering a fake tab.
 *
 * What they pin down: switching on opens a blank tab inside the click, records the choice
 * with where Marketplace should send the buyer back to, and only then sends the tab to
 * Marketplace - or closes it, when the save was refused or nothing is left to buy. Switching
 * on never makes the feature usable by itself; a failed save puts the switch back; and
 * where buying is not possible (Teams) or not allowed, nothing is offered for sale.
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

/** The SharePoint page the tests run on: a workbench query string, and a route in the fragment. */
const PAGE_PATH: string = '/sites/projects/_layouts/15/workbench.aspx?debug=true&noredir=true';
const ROUTE_FRAGMENT: string = '#/subscription';

const TEAMS_MESSAGE: string = 'Open XSProject in SharePoint to subscribe';

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

/** Switched on but not bought: Marketplace has not (yet) reported a purchase. */
const ENABLED_NOT_BOUGHT: IFeatureSubscription = { ...NOT_BOUGHT, isEnabled: true };

interface ICalls {
  reads: { feature: FeatureCode; refresh: boolean }[];
  writes: { feature: FeatureCode; isEnabled: boolean; returnUrl: string | undefined }[];
}

/** What `window.open('', '_blank')` hands back: just the parts the switch touches. */
interface IFakeTab {
  opener: unknown;
  location: { href: string };
  close: jest.Mock;
}

function fakeTab(): IFakeTab {
  return { opener: window, location: { href: '' }, close: jest.fn() };
}

/**
 * A service holding one feature's state. Switching it records the switch and nothing more,
 * as the real API does; `paid` stands in for Marketplace reporting the purchase.
 */
function stubService(
  initial: IFeatureSubscription,
  calls: ICalls,
  options: {
    failWrites?: boolean;
    failReads?: boolean;
    links?: ISubscriptionLinks;
    /** What the API answers a write with, in place of the stub's own bookkeeping. */
    answerWrite?: (isEnabled: boolean) => IFeatureSubscription;
  } = {}
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
    setFeatureEnabled: async (
      feature: FeatureCode,
      isEnabled: boolean,
      returnUrl?: string
    ): Promise<IFeatureSubscription> => {
      calls.writes.push({ feature, isEnabled, returnUrl });

      if (options.failWrites) {
        throw new Error('You do not have permission to change this feature.');
      }

      if (options.answerWrite) {
        current = options.answerWrite(isEnabled);
        return current;
      }

      return settle(isEnabled);
    },
    activateMarketplacePurchase: async (): Promise<IFeatureSubscription> => {
      throw new Error('The switch never activates a purchase.');
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
  let tab: IFakeTab;
  let originalUrl: string;
  /** Where Marketplace is expected to send the buyer back to: this page, without its route. */
  let returnUrl: string;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    // jsdom does not navigate; a link's own default action is not what these tests are about.
    host.addEventListener('click', (event: Event) => event.preventDefault());
    tab = fakeTab();
    openSpy = jest.spyOn(window, 'open').mockImplementation(() => tab as unknown as Window);
    originalUrl = window.location.href;
    window.history.replaceState(null, '', PAGE_PATH + ROUTE_FRAGMENT);
    returnUrl = window.location.origin + PAGE_PATH;
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
    openSpy.mockRestore();
    window.history.replaceState(null, '', originalUrl);
  });

  /** The one `ReactDOM.render` of this file, paired with the unmount in `afterEach`. */
  async function render(
    service: ISubscriptionService,
    canManage: boolean = true,
    isReturningFromMarketplace: boolean = false,
    isPurchaseAvailable: boolean = true
  ): Promise<void> {
    await act(async () => {
      ReactDOM.render(
        <FeatureSubscriptionSwitch
          feature="ED_CONTROL"
          subscriptionService={service}
          canManage={canManage}
          isReturningFromMarketplace={isReturningFromMarketplace}
          isPurchaseAvailable={isPurchaseAvailable}
        />,
        host
      );
    });
  }

  /** Subscribe in any form - a button, or the plain link it used to be. */
  function subscribeAction(): HTMLElement | undefined {
    return Array.prototype.slice
      .call(host.querySelectorAll('a, button'))
      .filter((element: HTMLElement) => element.textContent === 'Subscribe')[0];
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

  it('switching on records the choice, with the way back, then sends the tab to Marketplace', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(NOT_BOUGHT, calls));

    await flip();

    // The return URL is this page as Marketplace can reach it: the query string kept, the
    // app's route dropped (the landing page appends its own).
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl }]);
    expect(openSpy).toHaveBeenCalledTimes(1);
    expect(openSpy).toHaveBeenCalledWith('', '_blank');
    expect(tab.opener).toBeNull();
    expect(tab.location.href).toBe(MARKETPLACE_URL);
    expect(tab.close).not.toHaveBeenCalled();
    expect(toggle()?.checked).toBe(true);
  });

  // Opened after the save had answered, the tab would no longer count as the user's own
  // action and be popup-blocked; sent to Marketplace before it, a refused save would leave
  // the buyer paying for a purchase the API has no checkout for.
  it('opens the tab inside the click, and sends it to Marketplace only once the save is recorded', async () => {
    const calls: ICalls = newCalls();
    let answer: (subscription: IFeatureSubscription) => void = () => undefined;
    await render({
      ...stubService(NOT_BOUGHT, calls),
      setFeatureEnabled: (feature: FeatureCode, isEnabled: boolean, url?: string): Promise<IFeatureSubscription> => {
        calls.writes.push({ feature, isEnabled, returnUrl: url });
        return new Promise<IFeatureSubscription>((resolve) => {
          answer = resolve;
        });
      }
    });

    await flip();

    expect(openSpy).toHaveBeenCalledWith('', '_blank');
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl }]);
    expect(tab.location.href).toBe('');
    expect(host.textContent).toContain('Saving ED Control');

    await act(async () => {
      answer(ENABLED_NOT_BOUGHT);
    });

    expect(tab.location.href).toBe(MARKETPLACE_URL);
    expect(tab.opener).toBeNull();
  });

  it('closes the tab and says why when the save is refused', async () => {
    await render(stubService(NOT_BOUGHT, newCalls(), { failWrites: true }));

    await flip();

    expect(openSpy).toHaveBeenCalledWith('', '_blank');
    expect(tab.close).toHaveBeenCalledTimes(1);
    expect(tab.location.href).toBe('');
    expect(toggle()?.checked).toBe(false);
    const alert: string | null | undefined = host.querySelector('[role="alert"]')?.textContent;
    expect(alert).toContain('Unable to update ED Control.');
    expect(alert).toContain('You do not have permission to change this feature.');
  });

  it('closes the tab when the API answers that nothing is left to buy', async () => {
    await render(
      stubService(NOT_BOUGHT, newCalls(), {
        // Bought meanwhile, elsewhere: the read this switch started from was out of date.
        answerWrite: () => ({ ...NOT_BOUGHT, state: 'pending', subscriptionStatus: 'Pending', isEnabled: true })
      })
    );

    await flip();

    expect(openSpy).toHaveBeenCalledWith('', '_blank');
    expect(tab.close).toHaveBeenCalledTimes(1);
    expect(tab.location.href).toBe('');
    expect(host.textContent).toContain('Your subscription is being processed.');
  });

  it('still records the choice when the browser blocks the tab, and offers Marketplace as a link', async () => {
    const calls: ICalls = newCalls();
    openSpy.mockImplementation(() => null);
    await render(stubService(NOT_BOUGHT, calls));

    await flip();

    // The checkout exists once the save is recorded, so following the link is safe.
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl }]);
    expect(host.textContent).toContain('Your browser blocked the Microsoft Marketplace tab.');
    const link: HTMLAnchorElement | null = host.querySelector('a[href="' + MARKETPLACE_URL + '"]');
    expect(link?.target).toBe('_blank');
    expect(link?.rel).toBe('noopener noreferrer');

    await act(async () => {
      (link as HTMLAnchorElement).click();
    });

    // Gone to Marketplace by the link: coming back re-checks, as it does from the tab.
    expect(host.textContent).toContain('your subscription is being processed');
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

    // No tab flashes open for a feature that is already paid for.
    expect(openSpy).not.toHaveBeenCalled();
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl }]);
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

    // Switching off starts no purchase, so it says nothing about where to come back to.
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: false, returnUrl: undefined }]);
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
    // No checkout can exist without an offer to buy, so no way back is sent either.
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl: undefined }]);
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

  it('in Teams, records the choice without a tab or a way back, and says where to subscribe', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(NOT_BOUGHT, calls), true, false, false);

    expect(host.textContent).toContain(TEAMS_MESSAGE);
    expect(host.textContent).not.toContain('Turn on ED Control to subscribe in Microsoft Marketplace.');

    await flip();

    expect(openSpy).not.toHaveBeenCalled();
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl: undefined }]);
    expect(toggle()?.checked).toBe(true);
    expect(host.textContent).toContain('This feature requires an active subscription.');
    expect(host.textContent).toContain(TEAMS_MESSAGE);
    expect(subscribeAction()).toBeUndefined();
  });

  it('offers Subscribe for a feature switched on but not bought, through the same save', async () => {
    const calls: ICalls = newCalls();
    await render(stubService(ENABLED_NOT_BOUGHT, calls));

    const subscribe: HTMLElement | undefined = subscribeAction();
    // Not a plain link to Marketplace: that would be a purchase with no checkout behind it.
    expect(subscribe?.tagName).toBe('BUTTON');

    await act(async () => {
      (subscribe as HTMLElement).click();
    });

    expect(openSpy).toHaveBeenCalledWith('', '_blank');
    expect(calls.writes).toStrictEqual([{ feature: 'ED_CONTROL', isEnabled: true, returnUrl }]);
    expect(tab.opener).toBeNull();
    expect(tab.location.href).toBe(MARKETPLACE_URL);
  });

  it('closes the tab Subscribe opened when the save is refused', async () => {
    await render(stubService(ENABLED_NOT_BOUGHT, newCalls(), { failWrites: true }));

    await act(async () => {
      (subscribeAction() as HTMLElement).click();
    });

    expect(tab.close).toHaveBeenCalledTimes(1);
    expect(tab.location.href).toBe('');
    expect(host.querySelector('[role="alert"]')?.textContent).toContain('Unable to update ED Control.');
  });

  it('offers no Subscribe to a user who may not change the setting', async () => {
    await render(stubService(ENABLED_NOT_BOUGHT, newCalls()), false);

    expect(host.textContent).toContain('This feature requires an active subscription.');
    expect(subscribeAction()).toBeUndefined();
  });

  it('offers no Subscribe in Teams', async () => {
    await render(stubService(ENABLED_NOT_BOUGHT, newCalls()), true, false, false);

    expect(subscribeAction()).toBeUndefined();
    expect(host.textContent).toContain(TEAMS_MESSAGE);
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
