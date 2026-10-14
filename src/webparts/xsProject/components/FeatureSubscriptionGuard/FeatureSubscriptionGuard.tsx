import * as React from 'react';

import styles from './FeatureSubscriptionGuard.module.scss';
import type { IFeatureSubscriptionGuardProps } from './IFeatureSubscriptionGuardProps';
import type { IFeatureSubscriptionGuardState } from './IFeatureSubscriptionGuardState';
import SubscriptionPrompt from '../SubscriptionPrompt/SubscriptionPrompt';
import type { LockedSubscriptionState } from '../SubscriptionPrompt/ISubscriptionPromptProps';
import SubscriptionStatus from '../SubscriptionStatus/SubscriptionStatus';
import { FEATURE_NAMES, IFeatureSubscription } from '../../../../models/FeatureSubscription';

/**
 * Renders a paid feature only once the Web API reports that the tenant has bought it.
 *
 * ```tsx
 * <FeatureSubscriptionGuard feature="MINUTES" subscriptionService={...}>
 *   <MinutesContent />
 * </FeatureSubscriptionGuard>
 * ```
 *
 * One component for every paid feature, keyed on the feature code, so a new one needs no
 * subscription logic of its own. It shows, in turn: "Checking … subscription" until the
 * API answers (no action is offered before that), then the children under a plan/status
 * line when the answer is active, or `SubscriptionPrompt` when it is pending, suspended,
 * inactive or could not be read.
 *
 * **Only the API's answer opens the gate.** Following Manage subscription changes nothing
 * but when to look again; returning from Marketplace - by switching back to this tab, or
 * via the landing page's `?subscription=returned` redirect - triggers a fresh read rather
 * than being taken as proof of purchase. And this is still only a UI gate: the API
 * endpoints behind a paid feature must check the subscription themselves.
 *
 * **The guard sells nothing.** It offers no Subscribe: a purchase is only activated when
 * the Web API recorded a checkout before it, and that is `FeatureSubscriptionSwitch`'s
 * job, on the Subscription page (`#/subscription`), which also knows whether this user
 * and this host may buy.
 *
 * Reads go through `ISubscriptionService.getFeatureSubscription`, which caches per
 * feature for the session, so the router remounting this screen does not ask again; only
 * "Check again", a return from Marketplace and a changed `feature` bypass it.
 */
export default class FeatureSubscriptionGuard extends React.Component<
  IFeatureSubscriptionGuardProps,
  IFeatureSubscriptionGuardState
> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;
  /** Identifies the latest read, so an older one finishing late cannot overwrite it. */
  private _requestId: number = 0;

  public constructor(props: IFeatureSubscriptionGuardProps) {
    super(props);

    this.state = {
      isLoading: true,
      isUnverified: false,
      isAwaitingMarketplace: !!props.isReturningFromMarketplace
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    document.addEventListener('visibilitychange', this._onVisibilityChange);
    this._load(!!this.props.isReturningFromMarketplace);
  }

  public componentDidUpdate(previous: IFeatureSubscriptionGuardProps): void {
    if (previous.feature !== this.props.feature) {
      // A different feature: nothing about the previous answer applies to it.
      this.setState({ subscription: undefined, isAwaitingMarketplace: !!this.props.isReturningFromMarketplace });
      this._load(!!this.props.isReturningFromMarketplace);
    } else if (this.props.isReturningFromMarketplace && !previous.isReturningFromMarketplace) {
      this.setState({ isAwaitingMarketplace: true });
      this._load(true);
    }
  }

  public componentWillUnmount(): void {
    this._isActive = false;
    document.removeEventListener('visibilitychange', this._onVisibilityChange);
  }

  public render(): React.ReactElement<IFeatureSubscriptionGuardProps> {
    const { subscriptionService, feature, children } = this.props;
    const { isLoading, subscription, isUnverified, isAwaitingMarketplace } = this.state;
    const featureName: string = FEATURE_NAMES[feature];

    if (isLoading) {
      return (
        <div className={styles.status} role="status" aria-live="polite">
          Checking {featureName} subscription&hellip;
        </div>
      );
    }

    if (subscription && subscription.isActive && !isUnverified) {
      return (
        <>
          <SubscriptionStatus
            subscription={subscription}
            manageSubscriptionUrl={subscriptionService.links.manageSubscriptionUrl}
          />
          {children}
        </>
      );
    }

    return (
      <SubscriptionPrompt
        featureName={featureName}
        state={isUnverified || !subscription ? 'error' : FeatureSubscriptionGuard._toLockedState(subscription)}
        links={subscriptionService.links}
        isAwaitingMarketplace={isAwaitingMarketplace}
        onManage={this._onManage}
        onCheckAgain={this._onCheckAgain}
      />
    );
  }

  private _load(refresh: boolean): void {
    const requestId: number = ++this._requestId;
    const { feature } = this.props;

    this.setState({ isLoading: true, isUnverified: false });

    this.props.subscriptionService.getFeatureSubscription(feature, { refresh })
      .then((subscription: IFeatureSubscription) => {
        if (this._isCurrent(requestId)) {
          this.setState((state: IFeatureSubscriptionGuardState) => ({
            isLoading: false,
            subscription,
            // Nothing left to wait for once the API says it is active.
            isAwaitingMarketplace: subscription.isActive ? false : state.isAwaitingMarketplace
          }));
        }
      })
      .catch(() => {
        // The service has already logged the failure in full; what is shown is fixed text,
        // because no API message tells the user anything they can act on here.
        if (this._isCurrent(requestId)) {
          this.setState({ isLoading: false, subscription: undefined, isUnverified: true });
        }
      });
  }

  private _isCurrent(requestId: number): boolean {
    return this._isActive && requestId === this._requestId;
  }

  /** `subscription` is known not to be usable here; paid for means it is switched off. */
  private static _toLockedState(subscription: IFeatureSubscription): LockedSubscriptionState {
    return subscription.state === 'active' ? 'disabled' : subscription.state;
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` uses for its own click handler.

  /**
   * The user followed Manage subscription to Marketplace in a new tab.
   *
   * Deliberately changes nothing about the subscription itself: clicking is not paying.
   */
  private _onManage = (): void => {
    this.setState({ isAwaitingMarketplace: true });
  };

  private _onCheckAgain = (): void => {
    this._load(true);
  };

  /** Re-checks when the user switches back to this tab after going to Marketplace. */
  private _onVisibilityChange = (): void => {
    if (document.visibilityState === 'visible' && this.state.isAwaitingMarketplace && !this.state.isLoading) {
      this._load(true);
    }
  };
}
