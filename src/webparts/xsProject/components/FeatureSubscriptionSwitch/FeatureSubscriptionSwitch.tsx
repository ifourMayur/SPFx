import * as React from 'react';

import styles from './FeatureSubscriptionSwitch.module.scss';
import type { IFeatureSubscriptionSwitchProps } from './IFeatureSubscriptionSwitchProps';
import type { IFeatureSubscriptionSwitchState } from './IFeatureSubscriptionSwitchState';
import SubscriptionPrompt from '../SubscriptionPrompt/SubscriptionPrompt';
import type { LockedSubscriptionState } from '../SubscriptionPrompt/ISubscriptionPromptProps';
import SubscriptionStatus from '../SubscriptionStatus/SubscriptionStatus';
import {
  FEATURE_NAMES,
  IFeatureSubscription,
  isNavigableUrl,
  needsMarketplaceCheckout
} from '../../../../models/FeatureSubscription';

/**
 * An on/off switch for one paid feature, for the whole tenant.
 *
 * Switching on does two things: it records the choice through `POST …/enable`, and - when
 * the feature has not been bought yet - opens Microsoft Marketplace, where the plan is
 * chosen and paid for. **Switching on is not paying.** The API answers with the feature
 * switched on but still locked, and it only becomes usable once Marketplace has reported
 * the purchase and the API says `isActive`. Until then the switch shows why, under it:
 * "complete your subscription", "being processed", "suspended".
 *
 * Switching off records that too (`POST …/disable`) and takes effect at once; billing is
 * Marketplace's, so a paid subscription keeps running there until it is cancelled there.
 *
 * Marketplace opens in a **new tab**, from inside the click: a tab opened after the save
 * had answered would no longer count as the user's own action, and the browser would block
 * it. A new tab also keeps the sign-in and the chosen site, and works inside a Teams tab,
 * which Marketplace refuses to be framed in. Switching back to this tab re-checks.
 *
 * No setting is offered until the API has answered, and a user whose role may not change
 * it (`canManage`) sees the switch disabled - the API refuses them regardless.
 */
export default class FeatureSubscriptionSwitch extends React.Component<
  IFeatureSubscriptionSwitchProps,
  IFeatureSubscriptionSwitchState
> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;
  /** Identifies the latest read or write, so an older one finishing late cannot win. */
  private _requestId: number = 0;

  public constructor(props: IFeatureSubscriptionSwitchProps) {
    super(props);

    this.state = {
      isLoading: true,
      isSaving: false,
      isUnverified: false,
      isSaveFailed: false,
      isAwaitingMarketplace: !!props.isReturningFromMarketplace
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    document.addEventListener('visibilitychange', this._onVisibilityChange);
    this._load(!!this.props.isReturningFromMarketplace);
  }

  public componentDidUpdate(previous: IFeatureSubscriptionSwitchProps): void {
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

  public render(): React.ReactElement<IFeatureSubscriptionSwitchProps> {
    const { feature, canManage } = this.props;
    const { isLoading, isSaving, pendingEnabled, subscription, isSaveFailed } = this.state;
    const featureName: string = FEATURE_NAMES[feature];

    if (!subscription) {
      return isLoading
        ? (
          <div className={styles.status} role="status" aria-live="polite">
            Checking {featureName} subscription&hellip;
          </div>
        )
        : this._renderPrompt('error');
    }

    const switchId: string = `feature-switch-${feature}`;

    return (
      <div className={styles.featureSwitch}>
        <label className={styles.toggleField} htmlFor={switchId}>
          <span className={styles.toggle}>
            <input
              id={switchId}
              type="checkbox"
              role="switch"
              checked={isSaving && pendingEnabled !== undefined ? pendingEnabled : subscription.isEnabled}
              disabled={!canManage || isSaving || isLoading}
              onChange={this._onSwitchChange}
            />
            <span className={styles.toggleTrack} aria-hidden="true" />
          </span>
          <span className={styles.toggleLabel}>Enable {featureName}</span>
        </label>

        {!canManage && <p className={styles.note}>You do not have permission to change this setting.</p>}

        {isSaveFailed && (
          <p className={styles.failure} role="alert">
            Unable to update {featureName}. Please try again.
          </p>
        )}

        {this._renderDetail(subscription, featureName)}
      </div>
    );
  }

  /** What is said under the switch: the plan when usable, otherwise why it is not. */
  private _renderDetail(subscription: IFeatureSubscription, featureName: string): React.ReactNode {
    const { isLoading, isSaving, isUnverified } = this.state;
    const { links } = this.props.subscriptionService;

    if (isLoading || isSaving) {
      return (
        <p className={styles.note} role="status" aria-live="polite">
          {isSaving ? `Saving ${featureName}` : `Checking ${featureName} subscription`}&hellip;
        </p>
      );
    }

    if (isUnverified) {
      return this._renderPrompt('error');
    }

    if (subscription.isActive) {
      return <SubscriptionStatus subscription={subscription} manageSubscriptionUrl={links.manageSubscriptionUrl} />;
    }

    if (!subscription.isEnabled) {
      if (subscription.state === 'active') {
        return this._renderPrompt('disabled');
      }

      return (
        <p className={styles.note}>
          {isNavigableUrl(links.marketplaceOfferUrl)
            ? `Turn on ${featureName} to subscribe in Microsoft Marketplace.`
            : `${featureName} requires an active subscription. Subscribing is not available here yet.`}
        </p>
      );
    }

    // Switched on, not usable: not bought yet, being processed, or suspended. A paid-for
    // answer cannot be both switched on and unusable unless the API is mid-update, which
    // reads best as "being processed".
    return this._renderPrompt(subscription.state === 'active' ? 'pending' : subscription.state);
  }

  private _renderPrompt(state: LockedSubscriptionState): React.ReactElement {
    return (
      <SubscriptionPrompt
        featureName={FEATURE_NAMES[this.props.feature]}
        state={state}
        links={this.props.subscriptionService.links}
        isAwaitingMarketplace={this.state.isAwaitingMarketplace}
        onSubscribe={this._onSubscribe}
        onCheckAgain={this._onCheckAgain}
      />
    );
  }

  private _load(refresh: boolean): void {
    const requestId: number = ++this._requestId;

    this.setState({ isLoading: true, isUnverified: false });

    this.props.subscriptionService.getFeatureSubscription(this.props.feature, { refresh })
      .then((subscription: IFeatureSubscription) => {
        if (this._isCurrent(requestId)) {
          this.setState((state: IFeatureSubscriptionSwitchState) => ({
            isLoading: false,
            subscription,
            // Nothing left to wait for once the API says it is usable.
            isAwaitingMarketplace: subscription.isActive ? false : state.isAwaitingMarketplace
          }));
        }
      })
      .catch(() => {
        // The service has already logged the failure in full; what is shown is fixed text.
        if (this._isCurrent(requestId)) {
          this.setState({ isLoading: false, isUnverified: true });
        }
      });
  }

  private _isCurrent(requestId: number): boolean {
    return this._isActive && requestId === this._requestId;
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` uses for its own click handler.

  private _onSwitchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const isEnabled: boolean = event.currentTarget.checked;
    const { subscription, isSaving } = this.state;
    const { feature, subscriptionService } = this.props;
    const offerUrl: string = subscriptionService.links.marketplaceOfferUrl;

    if (!subscription || isSaving) {
      return;
    }

    const opensCheckout: boolean = isEnabled && needsMarketplaceCheckout(subscription) && isNavigableUrl(offerUrl);

    if (opensCheckout) {
      // Opened here, inside the click, rather than once the save has answered - by then it
      // would no longer count as the user's action and the browser would block it.
      window.open(offerUrl, '_blank', 'noopener,noreferrer');
    }

    const requestId: number = ++this._requestId;

    this.setState((state: IFeatureSubscriptionSwitchState) => ({
      isSaving: true,
      pendingEnabled: isEnabled,
      isSaveFailed: false,
      isUnverified: false,
      isAwaitingMarketplace: opensCheckout || state.isAwaitingMarketplace
    }));

    subscriptionService.setFeatureEnabled(feature, isEnabled)
      .then((saved: IFeatureSubscription) => {
        if (this._isCurrent(requestId)) {
          this.setState({ isSaving: false, pendingEnabled: undefined, subscription: saved });
        }
      })
      .catch(() => {
        // The switch falls back to the stored value; the service has logged why.
        if (this._isCurrent(requestId)) {
          this.setState({ isSaving: false, pendingEnabled: undefined, isSaveFailed: true });
        }
      });
  };

  /**
   * The user followed Subscribe or Manage subscription to Marketplace in a new tab.
   *
   * Deliberately changes nothing about the subscription itself: clicking is not paying.
   */
  private _onSubscribe = (): void => {
    this.setState({ isAwaitingMarketplace: true });
  };

  private _onCheckAgain = (): void => {
    this._load(true);
  };

  /** Re-checks when the user switches back to this tab after going to Marketplace. */
  private _onVisibilityChange = (): void => {
    const { isAwaitingMarketplace, isLoading, isSaving } = this.state;

    if (document.visibilityState === 'visible' && isAwaitingMarketplace && !isLoading && !isSaving) {
      this._load(true);
    }
  };
}
