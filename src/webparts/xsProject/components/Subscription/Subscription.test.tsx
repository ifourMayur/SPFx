/**
 * Behaviour tests for the Subscription page: the real component tree - the page and its
 * `FeatureSubscriptionSwitch`es - over a stubbed `ISubscriptionService`.
 *
 * What they pin down: one switch per paid feature; a Marketplace purchase token is handed
 * to the API once, however often the page mounts and however long the API takes; the
 * answer re-reads every switch and takes the token out of the URL, whether the API
 * activated anything or refused; and nothing is offered for sale to a user who may not buy,
 * or in a host (Teams) that Marketplace cannot send the buyer back to.
 *
 * The consumed-token record is module-level and outlives every mount, which is the point -
 * so each test uses a token of its own.
 *
 * Every `ReactDOM.render` here is unmounted in `afterEach`.
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import Subscription from './Subscription';
import {
  FEATURE_CODES,
  FeatureCode,
  IFeatureSubscription,
  ISubscriptionLinks
} from '../../../../models/FeatureSubscription';
import type {
  IFeatureSubscriptionRequestOptions,
  ISubscriptionService
} from '../../../../services/SubscriptionService';

const LINKS: ISubscriptionLinks = {
  marketplaceOfferUrl: 'https://marketplace.microsoft.com/',
  manageSubscriptionUrl: 'https://portal.azure.com/'
};

const TEAMS_MESSAGE: string = 'Open XSProject in SharePoint to subscribe';
const NON_MANAGER_MESSAGE: string = 'Ask an administrator to subscribe';

function notBought(feature: FeatureCode): IFeatureSubscription {
  return { feature, state: 'inactive', isActive: false, isEnabled: false, subscriptionStatus: 'Inactive' };
}

/** What the API answers once Marketplace has confirmed the purchase. */
function bought(feature: FeatureCode): IFeatureSubscription {
  return {
    feature,
    state: 'active',
    isActive: true,
    isEnabled: true,
    subscriptionStatus: 'Active',
    planName: 'Professional'
  };
}

interface IStub extends ISubscriptionService {
  /** Every token the page handed to the API. */
  activations: string[];
  /** Every read, by feature. */
  reads: FeatureCode[];
  /** Settles a held activation: with the feature it activated, or by refusing. */
  resolveActivation: (feature: FeatureCode) => void;
  rejectActivation: (error: Error) => void;
}

/**
 * A service holding every feature never bought. Activating records the token and, unless
 * held or refused, marks the purchase's feature bought - as the API does.
 */
function stubService(
  options: {
    /** Keeps the activation pending until the test settles it. */
    hold?: boolean;
    /** Refuses the activation with this error. */
    refuse?: Error;
    /** What the activation buys. */
    feature?: FeatureCode;
    /** What the API answers the activation with, in place of `bought(feature)`. */
    answer?: IFeatureSubscription;
  } = {}
): IStub {
  const state: { [feature: string]: IFeatureSubscription } = {};
  FEATURE_CODES.forEach((feature: FeatureCode) => {
    state[feature] = notBought(feature);
  });

  let settle: { resolve: (feature: FeatureCode) => void; reject: (error: Error) => void } | undefined;

  const stub: IStub = {
    links: LINKS,
    activations: [],
    reads: [],
    getSubscriptionStatus: async (): Promise<IFeatureSubscription[]> =>
      FEATURE_CODES.map((feature: FeatureCode) => state[feature]),
    getFeatureSubscription: async (
      feature: FeatureCode,
      _options?: IFeatureSubscriptionRequestOptions
    ): Promise<IFeatureSubscription> => {
      stub.reads.push(feature);
      return state[feature];
    },
    setFeatureEnabled: async (feature: FeatureCode, isEnabled: boolean): Promise<IFeatureSubscription> => {
      state[feature] = { ...state[feature], isEnabled };
      return state[feature];
    },
    activateMarketplacePurchase: (token: string): Promise<IFeatureSubscription> => {
      stub.activations.push(token);

      const activate = (feature: FeatureCode): IFeatureSubscription => {
        state[feature] = options.answer || bought(feature);
        return state[feature];
      };

      if (options.refuse) {
        return Promise.reject(options.refuse);
      }

      if (!options.hold) {
        return Promise.resolve(activate(options.feature || 'ED_CONTROL'));
      }

      return new Promise<IFeatureSubscription>((resolve, reject) => {
        settle = {
          resolve: (feature: FeatureCode): void => resolve(activate(feature)),
          reject
        };
      });
    },
    resolveActivation: (feature: FeatureCode): void => settle?.resolve(feature),
    rejectActivation: (error: Error): void => settle?.reject(error)
  };

  return stub;
}

describe('Subscription', () => {
  let host: HTMLDivElement;
  let openSpy: jest.SpyInstance;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    openSpy = jest.spyOn(window, 'open').mockImplementation(() => null);
  });

  afterEach(() => {
    unmountHost();
    host.remove();
    openSpy.mockRestore();
  });

  /** The one `ReactDOM.render` of this file, paired with the unmount in `afterEach`. */
  async function render(
    service: ISubscriptionService,
    props: {
      canManage?: boolean;
      isPurchaseAvailable?: boolean;
      marketplaceToken?: string;
      isReturningFromMarketplace?: boolean;
      onTokenConsumed?: () => void;
    } = {}
  ): Promise<void> {
    await act(async () => {
      ReactDOM.render(
        <Subscription
          subscriptionService={service}
          canManage={props.canManage === undefined ? true : props.canManage}
          isPurchaseAvailable={props.isPurchaseAvailable === undefined ? true : props.isPurchaseAvailable}
          marketplaceToken={props.marketplaceToken}
          isReturningFromMarketplace={!!props.isReturningFromMarketplace}
          onTokenConsumed={props.onTokenConsumed || ((): void => undefined)}
        />,
        host
      );
    });
  }

  /** The one unmount of this file: after each test, and where a test remounts the page. */
  function unmountHost(): void {
    ReactDOM.unmountComponentAtNode(host);
  }

  async function unmount(): Promise<void> {
    await act(async () => {
      unmountHost();
    });
  }

  /** Lets settled promises run their `then`s, and the renders they cause. */
  async function flush(): Promise<void> {
    await act(async () => {
      await Promise.resolve();
    });
  }

  function toggle(feature: FeatureCode): HTMLInputElement | null {
    return host.querySelector(`#feature-switch-${feature}`);
  }

  it('shows one switch for every paid feature', async () => {
    await render(stubService());

    expect(host.querySelector('h3')?.textContent).toBe('Subscription');
    FEATURE_CODES.forEach((feature: FeatureCode) => {
      expect(toggle(feature)).not.toBeNull();
    });
    expect(host.querySelectorAll('[role="switch"]').length).toBe(FEATURE_CODES.length);
  });

  it('activates nothing, and leaves the URL alone, without a token', async () => {
    const service: IStub = stubService();
    const onTokenConsumed: jest.Mock = jest.fn();

    await render(service, { onTokenConsumed });

    expect(service.activations).toStrictEqual([]);
    expect(onTokenConsumed).not.toHaveBeenCalled();
    expect(host.textContent).not.toContain('Activating your purchase');
  });

  it('says it is activating until the API answers, then that the feature is active', async () => {
    const service: IStub = stubService({ hold: true });
    const onTokenConsumed: jest.Mock = jest.fn();

    await render(service, { marketplaceToken: 'token-activating', isReturningFromMarketplace: true, onTokenConsumed });

    expect(service.activations).toStrictEqual(['token-activating']);
    expect(host.textContent).toContain('Activating your purchase…');
    // The token stays in the URL until the API has answered.
    expect(onTokenConsumed).not.toHaveBeenCalled();

    await act(async () => {
      service.resolveActivation('MINUTES');
    });

    expect(host.textContent).not.toContain('Activating your purchase');
    expect(host.textContent).toContain('Minutes of Meetings is now active for your organisation');
    expect(onTokenConsumed).toHaveBeenCalledTimes(1);
  });

  it('hands a token to the API once across a remount, however long the API takes', async () => {
    const service: IStub = stubService({ hold: true });
    const firstConsumed: jest.Mock = jest.fn();
    const secondConsumed: jest.Mock = jest.fn();

    await render(service, { marketplaceToken: 'token-remount', onTokenConsumed: firstConsumed });
    // Re-rendered with the same token while the API is still answering.
    await render(service, { marketplaceToken: 'token-remount', onTokenConsumed: firstConsumed });
    await unmount();
    await render(service, { marketplaceToken: 'token-remount', onTokenConsumed: secondConsumed });

    expect(service.activations).toStrictEqual(['token-remount']);
    // The remounted page does not leave a spent token in the URL.
    expect(secondConsumed).toHaveBeenCalledTimes(1);

    await act(async () => {
      service.resolveActivation('ED_CONTROL');
    });
    await unmount();
    await render(service, { marketplaceToken: 'token-remount' });

    expect(service.activations).toStrictEqual(['token-remount']);
    // The page that started it is gone, so it navigates nowhere.
    expect(firstConsumed).not.toHaveBeenCalled();
  });

  it('activates a second, different token', async () => {
    const service: IStub = stubService();

    await render(service, { marketplaceToken: 'token-first' });
    await render(service, { marketplaceToken: 'token-second' });
    await flush();

    expect(service.activations).toStrictEqual(['token-first', 'token-second']);
  });

  it('re-reads every switch once the purchase is activated, which then shows it active', async () => {
    const service: IStub = stubService({ hold: true });

    await render(service, { marketplaceToken: 'token-rekey', isReturningFromMarketplace: true });

    expect(service.reads.length).toBe(FEATURE_CODES.length);
    expect(toggle('ED_CONTROL')?.checked).toBe(false);

    await act(async () => {
      service.resolveActivation('ED_CONTROL');
    });
    await flush();

    // Every switch mounted afresh and asked again, rather than keeping the answer it read
    // before the purchase was activated.
    expect(service.reads.length).toBe(FEATURE_CODES.length * 2);
    expect(toggle('ED_CONTROL')?.checked).toBe(true);
    expect(host.textContent).toContain('Plan: Professional');
  });

  it('shows the API\'s refusal, activates nothing and still takes the token out of the URL', async () => {
    const service: IStub = stubService({ refuse: new Error('This purchase could not be verified with Microsoft.') });
    const onTokenConsumed: jest.Mock = jest.fn();

    await render(service, { marketplaceToken: 'token-refused', isReturningFromMarketplace: true, onTokenConsumed });
    await flush();

    const alert: Element | null = host.querySelector('[role="alert"]');

    expect(alert?.textContent).toContain('This purchase could not be verified with Microsoft.');
    expect(host.textContent).not.toContain('is now active');
    expect(onTokenConsumed).toHaveBeenCalledTimes(1);
    FEATURE_CODES.forEach((feature: FeatureCode) => {
      expect(toggle(feature)?.checked).toBe(false);
    });
    // Re-read after the refusal too, as after a success.
    expect(service.reads.length).toBe(FEATURE_CODES.length * 2);
  });

  it('does not call a feature active that the API did not report active', async () => {
    const service: IStub = stubService({ answer: { ...notBought('TENDER'), isEnabled: true, state: 'pending' } });

    await render(service, { marketplaceToken: 'token-pending' });
    await flush();

    expect(host.textContent).not.toContain('is now active');
    expect(host.textContent).toContain('Your Tender purchase has been received');
  });

  it('shows a non-administrator every switch disabled, and who to ask', async () => {
    await render(stubService(), { canManage: false });

    FEATURE_CODES.forEach((feature: FeatureCode) => {
      expect(toggle(feature)?.disabled).toBe(true);
    });
    expect(host.textContent).toContain(NON_MANAGER_MESSAGE);
    expect(host.textContent).not.toContain('Subscribe​');
  });

  it('does not tell an administrator to ask an administrator', async () => {
    await render(stubService());

    expect(toggle('ED_CONTROL')?.disabled).toBe(false);
    expect(host.textContent).not.toContain(NON_MANAGER_MESSAGE);
  });

  it('offers no purchase in Teams', async () => {
    await render(stubService(), { isPurchaseAvailable: false });

    expect(host.textContent).toContain(TEAMS_MESSAGE);
    expect(host.textContent).not.toContain('subscribe in Microsoft Marketplace');

    await act(async () => {
      (toggle('ED_CONTROL') as HTMLInputElement).click();
    });

    expect(openSpy).not.toHaveBeenCalled();
    const subscribe: HTMLElement[] = Array.prototype.slice
      .call(host.querySelectorAll('a, button'))
      .filter((element: HTMLElement) => element.textContent === 'Subscribe');
    expect(subscribe).toStrictEqual([]);
  });

  it('offers the purchase in SharePoint', async () => {
    await render(stubService());

    expect(host.textContent).toContain('Turn on ED Control to subscribe in Microsoft Marketplace.');
    expect(host.textContent).not.toContain(TEAMS_MESSAGE);
  });
});
