import * as React from 'react';

import styles from './Subscription.module.scss';
import type { ISubscriptionProps } from './ISubscriptionProps';
import type { ISubscriptionState } from './ISubscriptionState';
import FeatureSubscriptionSwitch from '../FeatureSubscriptionSwitch/FeatureSubscriptionSwitch';
import { getErrorMessage } from '../../../../models/ApiError';
import {
  FEATURE_CODES,
  FEATURE_NAMES,
  FeatureCode,
  IFeatureSubscription
} from '../../../../models/FeatureSubscription';

/**
 * Every Marketplace purchase token already handed to the Web API this session.
 *
 * Module-level rather than state, because it has to outlive the page: the router remounts
 * it, and "Change site" remounts the whole app while the token is still in the URL. A token
 * is added **before** the API is called, so a remount while the API is still answering
 * cannot hand it over a second time.
 */
const consumedTokens: Set<string> = new Set<string>();

/**
 * The Subscription page: where the tenant's paid features - sold through Microsoft
 * Marketplace - are switched on and off, and where a purchase is activated on the way back
 * from Marketplace.
 *
 * Everything per feature is `FeatureSubscriptionSwitch`, one for each of `FEATURE_CODES`,
 * so a new paid feature needs nothing here. What the page adds is the return leg of a
 * purchase: the Marketplace landing page sends the buyer back to
 * `#/subscription?subscription=returned&marketplaceToken=…`, and the page hands that token
 * to the Web API, which verifies it with Microsoft and records the purchase against the
 * caller's organisation. The token is handed over once, whatever the API answers; then the
 * switches are re-keyed, so they mount afresh and read what the API now holds, and the
 * token is taken out of the URL.
 *
 * Nothing here makes a feature usable: the activation's answer is shown, but only the API's
 * `isActive` - read again by every switch - is what the switches go by.
 */
export default class Subscription extends React.Component<ISubscriptionProps, ISubscriptionState> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;

  public constructor(props: ISubscriptionProps) {
    super(props);

    this.state = { activationGeneration: 0 };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._activate(this.props.marketplaceToken);
  }

  public componentDidUpdate(previous: ISubscriptionProps): void {
    if (this.props.marketplaceToken !== previous.marketplaceToken) {
      this._activate(this.props.marketplaceToken);
    }
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<ISubscriptionProps> {
    const { subscriptionService, canManage, isPurchaseAvailable, isReturningFromMarketplace } = this.props;
    const { activationGeneration } = this.state;

    return (
      <section className={styles.subscription}>
        <h3 className={styles.title}>Subscription</h3>

        {this._renderActivation()}

        {!canManage && <p className={styles.note}>Ask an administrator to subscribe</p>}

        <ul className={styles.features}>
          {FEATURE_CODES.map((feature: FeatureCode) => (
            <li key={`${feature}-${activationGeneration}`} className={styles.feature}>
              <h4 className={styles.featureName}>{FEATURE_NAMES[feature]}</h4>
              <FeatureSubscriptionSwitch
                feature={feature}
                subscriptionService={subscriptionService}
                canManage={canManage}
                isPurchaseAvailable={isPurchaseAvailable}
                isReturningFromMarketplace={isReturningFromMarketplace}
              />
            </li>
          ))}
        </ul>
      </section>
    );
  }

  private _renderActivation(): React.ReactNode {
    const { activationStatus, activationMessage } = this.state;

    switch (activationStatus) {
      case 'activating':
        return (
          <p className={styles.status} role="status" aria-live="polite">
            Activating your purchase&hellip;
          </p>
        );
      case 'activated':
        return (
          <p className={styles.success} role="status" aria-live="polite">
            {activationMessage}
          </p>
        );
      case 'refused':
        return (
          <p className={styles.failure} role="alert">
            Unable to activate your purchase. {activationMessage}
          </p>
        );
      default:
        return null;
    }
  }

  /** Hands a token to the Web API, unless it already has been. */
  private _activate(token: string | undefined): void {
    if (!token) {
      return;
    }

    if (consumedTokens.has(token)) {
      // Already handed over - before a remount, perhaps still being answered. Spent either
      // way, so out of the URL; on the next tick, because the router ignores a navigation
      // asked for while it is still mounting, which a remounted app is.
      Promise.resolve()
        .then(() => {
          if (this._isActive) {
            this.props.onTokenConsumed();
          }
        })
        .catch(() => undefined);
      return;
    }

    consumedTokens.add(token);
    this.setState({ activationStatus: 'activating', activationMessage: undefined });

    this.props.subscriptionService.activateMarketplacePurchase(token).then(
      (subscription: IFeatureSubscription) => {
        const featureName: string = FEATURE_NAMES[subscription.feature];

        // Only the API's `isActive` calls a feature active; an answer that is not yet
        // says so instead.
        this._settle(
          'activated',
          subscription.isActive
            ? `${featureName} is now active for your organisation`
            : `Your ${featureName} purchase has been received, but ${featureName} is not active yet.`
        );
      },
      (error: unknown) => {
        // The service has logged the failure in full; the API's own message says why.
        this._settle('refused', getErrorMessage(error, 'Please contact support.'));
      }
    );
  }

  /**
   * Shows the API's answer, re-keys the switches so they read again, and takes the spent
   * token out of the URL - whether the purchase was activated or refused.
   */
  private _settle(activationStatus: 'activated' | 'refused', activationMessage: string): void {
    if (!this._isActive) {
      return;
    }

    this.setState((state: ISubscriptionState) => ({
      activationGeneration: state.activationGeneration + 1,
      activationStatus,
      activationMessage
    }));
    this.props.onTokenConsumed();
  }
}
