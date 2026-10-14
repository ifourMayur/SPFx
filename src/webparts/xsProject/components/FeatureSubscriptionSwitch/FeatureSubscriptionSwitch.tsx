import * as React from 'react';

import styles from './FeatureSubscriptionSwitch.module.scss';
import type { IFeatureSubscriptionSwitchProps } from './IFeatureSubscriptionSwitchProps';
import type { IFeatureSubscriptionSwitchState } from './IFeatureSubscriptionSwitchState';
import SubscriptionPrompt from '../SubscriptionPrompt/SubscriptionPrompt';
import type { LockedSubscriptionState } from '../SubscriptionPrompt/ISubscriptionPromptProps';
import SubscriptionStatus from '../SubscriptionStatus/SubscriptionStatus';
import { getErrorMessage } from '../../../../models/ApiError';
import {
  FEATURE_NAMES,
  IFeatureSubscription,
  isNavigableUrl,
  needsMarketplaceCheckout
} from '../../../../models/FeatureSubscription';

/** Said instead of offering a purchase where this host cannot make one (Teams). */
const PURCHASE_UNAVAILABLE_MESSAGE: string = 'Open XSProject in SharePoint to subscribe';

/**
 * An on/off switch for one paid feature, for the whole tenant.
 *
 * Switching on does two things: it records the choice through `POST …/enable`, and - when
 * the feature has not been bought yet - sends a tab to Microsoft Marketplace, where the
 * plan is chosen and paid for. **Switching on is not paying.** The API answers with the
 * feature switched on but still locked, and it only becomes usable once Marketplace has
 * reported the purchase and the API says `isActive`. Until then the switch shows why, under
 * it: "complete your subscription", "being processed", "suspended".
 *
 * Switching off records that too (`POST …/disable`) and takes effect at once; billing is
 * Marketplace's, so a paid subscription keeps running there until it is cancelled there.
 *
 * **A purchase goes in a fixed order**, and Subscribe under the switch goes the same way:
 * 1. Inside the click, a blank tab is opened. One opened after the save had answered would
 *    no longer count as the user's own action, and the browser would block it.
 * 2. The choice is saved together with a return URL - this page without its route - which
 *    the Web API keeps as a pending checkout. That is how Marketplace's landing page later
 *    finds its way back here with the purchase.
 * 3. Only then is the tab sent to Marketplace. A refused save closes it instead, so nobody
 *    pays for a purchase the API has no checkout for; so does an answer with nothing left
 *    to buy. A tab the browser blocks anyway does not lose the purchase: the choice is
 *    still saved, and Marketplace is offered as a link.
 *
 * A new tab keeps the sign-in and the chosen site; switching back to this one re-checks.
 * Inside Teams no purchase is offered at all (`isPurchaseAvailable`): the page Marketplace
 * would return to only works inside Teams, so the switch records the choice and says to
 * subscribe from SharePoint.
 *
 * No setting is offered until the API has answered, and a user whose role may not change
 * it (`canManage`) sees the switch disabled and no Subscribe - the API refuses them
 * regardless.
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
      isPopupBlocked: false,
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
      this.setState({
        subscription: undefined,
        isPopupBlocked: false,
        isAwaitingMarketplace: !!this.props.isReturningFromMarketplace
      });
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
    const { isLoading, isSaving, pendingEnabled, subscription, saveError } = this.state;
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

        {saveError !== undefined && (
          <p className={styles.failure} role="alert">
            Unable to update {featureName}. {saveError}
          </p>
        )}

        {this._renderBlockedTab(subscription)}

        {this._renderDetail(subscription, featureName)}
      </div>
    );
  }

  /**
   * The browser blocked the Marketplace tab after the choice was saved: Marketplace as a
   * link instead, for as long as there is still something to buy there.
   */
  private _renderBlockedTab(subscription: IFeatureSubscription): React.ReactNode {
    const offerUrl: string = this.props.subscriptionService.links.marketplaceOfferUrl;

    if (!this.state.isPopupBlocked || !needsMarketplaceCheckout(subscription) || !isNavigableUrl(offerUrl)) {
      return null;
    }

    return (
      <p className={styles.note} role="alert">
        Your browser blocked the Microsoft Marketplace tab.{' '}
        <a href={offerUrl} target="_blank" rel="noopener noreferrer" onClick={this._onMarketplaceLinkFollowed}>
          Open Microsoft Marketplace
        </a>{' '}
        to complete your subscription.
      </p>
    );
  }

  /** What is said under the switch: the plan when usable, otherwise why it is not. */
  private _renderDetail(subscription: IFeatureSubscription, featureName: string): React.ReactNode {
    const { isLoading, isSaving, isUnverified } = this.state;
    const { isPurchaseAvailable } = this.props;
    const { links } = this.props.subscriptionService;
    const canBeBought: boolean = isNavigableUrl(links.marketplaceOfferUrl);

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

      let note: string = `${featureName} requires an active subscription. Subscribing is not available here yet.`;

      if (canBeBought) {
        note = isPurchaseAvailable
          ? `Turn on ${featureName} to subscribe in Microsoft Marketplace.`
          : PURCHASE_UNAVAILABLE_MESSAGE;
      }

      return <p className={styles.note}>{note}</p>;
    }

    // Switched on, not usable: not bought yet, being processed, or suspended. A paid-for
    // answer cannot be both switched on and unusable unless the API is mid-update, which
    // reads best as "being processed".
    const state: LockedSubscriptionState = subscription.state === 'active' ? 'pending' : subscription.state;

    if (state === 'inactive' && canBeBought && !isPurchaseAvailable) {
      return (
        <>
          {this._renderPrompt(state)}
          <p className={styles.note}>{PURCHASE_UNAVAILABLE_MESSAGE}</p>
        </>
      );
    }

    return this._renderPrompt(state);
  }

  private _renderPrompt(state: LockedSubscriptionState): React.ReactElement {
    const { feature, canManage, isPurchaseAvailable, subscriptionService } = this.props;

    return (
      <SubscriptionPrompt
        featureName={FEATURE_NAMES[feature]}
        state={state}
        links={subscriptionService.links}
        isAwaitingMarketplace={this.state.isAwaitingMarketplace}
        // Subscribe only where this user may buy and this host can come back from it.
        onSubscribe={canManage && isPurchaseAvailable ? this._onSubscribe : undefined}
        onManage={this._onManage}
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

  /**
   * Saves the switch - and, switching on, runs the purchase order from the class docblock.
   * Must be called synchronously from a click, so the tab it opens counts as the user's.
   */
  private _save(isEnabled: boolean): void {
    const { subscription, isSaving } = this.state;
    const { feature, subscriptionService, isPurchaseAvailable } = this.props;
    const offerUrl: string = subscriptionService.links.marketplaceOfferUrl;

    if (!subscription || isSaving) {
      return;
    }

    // Only a purchase says where to come back to: not switching off, not from a host
    // Marketplace could not return to, and not where no offer link is configured (the API then
    // writes no checkout, so there is nothing to come back from).
    const isPurchase: boolean = isEnabled && isPurchaseAvailable && isNavigableUrl(offerUrl);
    // A tab only when there may be something to buy: a feature that is paid for, pending or
    // suspended has nothing Marketplace should sell, and must not flash a tab open.
    const wantsTab: boolean = isPurchase && needsMarketplaceCheckout(subscription);
    // Blank, and without `noopener`: with it the browser hands back no tab to send on later.
    const tab: Window | null = wantsTab ? window.open('', '_blank') : null;
    const returnUrl: string | undefined = isPurchase
      ? FeatureSubscriptionSwitch._toReturnUrl(window.location.href)
      : undefined;
    const requestId: number = ++this._requestId;

    this.setState({
      isSaving: true,
      pendingEnabled: isEnabled,
      saveError: undefined,
      isUnverified: false,
      isPopupBlocked: false
    });

    // The two-argument form, not `.then(...).catch(...)`: only the service's own rejection may
    // close the tab and report a refused save. A fault in the success path, after the tab has
    // been sent to Marketplace, must not close a tab the buyer is already paying in.
    subscriptionService.setFeatureEnabled(feature, isEnabled, returnUrl)
      .then((saved: IFeatureSubscription) => {
        // Decided by the answer, not by the read the click started from: that may be stale.
        const opensCheckout: boolean = isPurchase && needsMarketplaceCheckout(saved) && isNavigableUrl(offerUrl);

        // Seen to even if this switch has gone meanwhile: a tab left blank is one nobody
        // asked for.
        if (tab) {
          if (opensCheckout) {
            // What the missing `noopener` would have done: Marketplace gets no handle on
            // this page through `window.opener`.
            tab.opener = null;
            tab.location.href = offerUrl;
          } else {
            tab.close();
          }
        }

        if (this._isCurrent(requestId)) {
          this.setState((state: IFeatureSubscriptionSwitchState) => ({
            isSaving: false,
            pendingEnabled: undefined,
            subscription: saved,
            isAwaitingMarketplace: (opensCheckout && !!tab) || state.isAwaitingMarketplace,
            isPopupBlocked: opensCheckout && wantsTab && !tab
          }));
        }
      }, (error: unknown) => {
        // Never sent on: there is no checkout behind a refused save to buy against.
        if (tab) {
          tab.close();
        }

        // The switch falls back to the stored value; the service has logged the failure in
        // full, and the API's own message says why - a refused return URL, say.
        if (this._isCurrent(requestId)) {
          this.setState({
            isSaving: false,
            pendingEnabled: undefined,
            saveError: getErrorMessage(error, 'Please try again.')
          });
        }
      });
  }

  /**
   * This page as Marketplace's landing page can send the buyer back to it: the query string
   * kept (the workbench's `?debug=…` included), the app's route in the fragment dropped -
   * the landing page appends its own.
   */
  private static _toReturnUrl(href: string): string {
    const fragmentStart: number = href.indexOf('#');
    return fragmentStart === -1 ? href : href.substring(0, fragmentStart);
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` uses for its own click handler.

  private _onSwitchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this._save(event.currentTarget.checked);
  };

  /** Subscribe: the same save as switching on, so the checkout exists before Marketplace. */
  private _onSubscribe = (): void => {
    this._save(true);
  };

  /**
   * The user followed Manage subscription to Marketplace in a new tab.
   *
   * Deliberately changes nothing about the subscription itself: clicking is not paying.
   */
  private _onManage = (): void => {
    this.setState({ isAwaitingMarketplace: true });
  };

  /** The user followed the link offered in place of a blocked tab. */
  private _onMarketplaceLinkFollowed = (): void => {
    this.setState({ isAwaitingMarketplace: true, isPopupBlocked: false });
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
