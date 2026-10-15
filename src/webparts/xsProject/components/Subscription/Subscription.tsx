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
import { ISubscriptionService } from '../../../../services/SubscriptionService';

/** Said, instead of activating, to a user who may not buy (`canManage`). */
const NOT_PERMITTED_MESSAGE: string =
  'Only an administrator can complete this purchase. Ask an administrator to sign in and open the Subscription page.';

/**
 * Every Marketplace purchase token the Web API has activated a purchase with this session:
 * spent for good.
 *
 * Module-level rather than state - like the two records below - because it has to outlive
 * the page: the router remounts it, and "Change site" remounts the whole app while the token
 * is still in the URL.
 */
const activatedTokens: Set<string> = new Set<string>();

/**
 * Activations the API is still answering, by token. A token is in here **before** the API is
 * called, so a remount while the API is still answering waits for that same answer rather
 * than hands the token over a second time.
 */
const pendingActivations: Map<string, Promise<IFeatureSubscription>> = new Map<string, Promise<IFeatureSubscription>>();

/**
 * The last activation that failed - a timeout, a network error, the API temporarily
 * unavailable - kept so "Try again" can send the same token again rather than force a second
 * checkout. The API answers a repeat of the same token without asking Microsoft again.
 *
 * Only ever here: by then the token is out of the URL, and it is never put in state, shown or
 * logged. A new activation takes the slot; the token's own success leaves it empty.
 */
let failedActivation: { token: string; message: string } | undefined;

/**
 * Hands `token` to the Web API, unless it is already being answered, and keeps the records
 * above current whichever page - if any - is still mounted when the answer comes.
 */
function startActivation(service: ISubscriptionService, token: string): Promise<IFeatureSubscription> {
  const pending: Promise<IFeatureSubscription> | undefined = pendingActivations.get(token);

  if (pending) {
    return pending;
  }

  const request: Promise<IFeatureSubscription> = service.activateMarketplacePurchase(token);
  pendingActivations.set(token, request);
  failedActivation = undefined;

  // Registered before any page's own handlers, so the records are current when those run.
  request.then(
    () => {
      pendingActivations.delete(token);
      activatedTokens.add(token);
    },
    (error: unknown) => {
      // Not spent: kept to be tried again. The service has logged the failure in full; the
      // API's own message says why.
      pendingActivations.delete(token);
      failedActivation = { token, message: getErrorMessage(error, 'Please contact support.') };
    }
  );

  return request;
}

/**
 * Test seam: forgets every token and failed activation recorded this session, so one test's
 * records cannot show up in the next. Production code never calls it.
 */
export function forgetMarketplaceActivations(): void {
  activatedTokens.clear();
  pendingActivations.clear();
  failedActivation = undefined;
}

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
 * caller's organisation. Once the API has answered, the token is taken out of the URL and
 * the switches are re-keyed, so they mount afresh and read what the API now holds. A token
 * the API activated is spent; one whose activation failed is kept in memory for "Try
 * again", which sends that same token again. A user who may not buy is told who can, and
 * the token is not handed over at all - the API would refuse it.
 *
 * Nothing here makes a feature usable: the activation's answer is shown, but only the API's
 * `isActive` - read again by every switch - is what the switches go by.
 */
export default class Subscription extends React.Component<ISubscriptionProps, ISubscriptionState> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;

  public constructor(props: ISubscriptionProps) {
    super(props);

    // A failure from before a remount is still there to try again.
    this.state = failedActivation
      ? { activationGeneration: 0, activationStatus: 'refused', activationMessage: failedActivation.message, canRetry: true }
      : { activationGeneration: 0 };
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
    const { activationStatus, activationMessage, canRetry } = this.state;

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
          <div className={styles.failure} role="alert">
            <p className={styles.failureMessage}>Unable to activate your purchase. {activationMessage}</p>
            {canRetry && (
              <button type="button" className={styles.retry} onClick={this._onRetry}>
                Try again
              </button>
            )}
          </div>
        );
      case 'notPermitted':
        return (
          <p className={styles.failure} role="alert">
            {NOT_PERMITTED_MESSAGE}
          </p>
        );
      default:
        return null;
    }
  }

  /** Hands a token from the URL to the Web API, unless it already has been. */
  private _activate(token: string | undefined): void {
    if (!token) {
      return;
    }

    if (!this.props.canManage) {
      // The API refuses other roles: say who can complete the purchase rather than hand the
      // token over only to have it refused.
      this.setState({ activationStatus: 'notPermitted', activationMessage: undefined, canRetry: false });
      this._takeTokenOutOfUrlSoon();
      return;
    }

    if (activatedTokens.has(token)) {
      // Spent before a remount: out of the URL, and nothing more to say.
      this._takeTokenOutOfUrlSoon();
      return;
    }

    if (pendingActivations.has(token)) {
      // Handed over before a remount and still being answered: wait for that answer rather
      // than ask again. The URL is cleaned now, so the answer leaves it alone.
      this._takeTokenOutOfUrlSoon();
      this._await(token, false);
      return;
    }

    this._await(token, true);
  }

  /**
   * Shows "activating" until the API has answered `token`, then the answer.
   *
   * @param cleansUrl - whether the answer takes the token out of the URL: only for a token
   *   that is still there.
   */
  private _await(token: string, cleansUrl: boolean): void {
    this.setState({ activationStatus: 'activating', activationMessage: undefined, canRetry: false });

    startActivation(this.props.subscriptionService, token).then(
      (subscription: IFeatureSubscription) => {
        const featureName: string = FEATURE_NAMES[subscription.feature];

        // Only the API's `isActive` calls a feature active; an answer that is not yet
        // says so instead.
        this._settle(
          'activated',
          subscription.isActive
            ? `${featureName} is now active for your organisation`
            : `Your ${featureName} purchase has been received, but ${featureName} is not active yet.`,
          cleansUrl
        );
      },
      (error: unknown) => {
        // The service has logged the failure in full; the API's own message says why.
        this._settle('refused', getErrorMessage(error, 'Please contact support.'), cleansUrl);
      }
    );
  }

  /**
   * Takes the token out of the URL, re-keys the switches so they read again, and shows the
   * API's answer - whether the purchase was activated or refused.
   *
   * The URL goes first, so the switches mount afresh on a page that is no longer a return
   * from Marketplace and do not keep waiting for one.
   */
  private _settle(activationStatus: 'activated' | 'refused', activationMessage: string, cleansUrl: boolean): void {
    if (!this._isActive) {
      return;
    }

    if (cleansUrl) {
      this.props.onTokenConsumed();
    }

    this.setState((state: ISubscriptionState) => ({
      activationGeneration: state.activationGeneration + 1,
      activationStatus,
      activationMessage,
      // Offered while the failed token is still kept - not once a newer one has taken its place.
      canRetry: activationStatus === 'refused' && !!failedActivation
    }));
  }

  /**
   * Out of the URL on the next tick, because the router ignores a navigation asked for while
   * it is still mounting, which a remounted app is.
   */
  private _takeTokenOutOfUrlSoon(): void {
    Promise.resolve()
      .then(() => {
        if (this._isActive) {
          this.props.onTokenConsumed();
        }
      })
      .catch(() => undefined);
  }

  // Assigned as a property so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` uses for its own click handler.

  /** "Try again": the failed token once more. It is no longer in the URL, so that is left alone. */
  private _onRetry = (): void => {
    if (failedActivation) {
      this._await(failedActivation.token, false);
    }
  };
}
