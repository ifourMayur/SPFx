import * as React from 'react';

import styles from './SubscriptionPrompt.module.scss';
import type { ISubscriptionPromptProps } from './ISubscriptionPromptProps';
import { isNavigableUrl } from '../../../../models/FeatureSubscription';

/**
 * What a locked paid feature shows instead of itself: why it is locked, and the one action
 * that can change that.
 *
 * Purely presentational. It never decides whether the feature is active - it is only
 * rendered once its owner has an answer that says it is not - and nothing it does can make
 * it active: Subscribe hands the purchase to its owner, where the plan is then chosen and
 * paid for in Microsoft Marketplace, and "Check again" asks the Web API.
 *
 * Subscribe is a **button, not a link to Marketplace**. A purchase can only be activated
 * when the Web API recorded a checkout before it - that is how the landing page finds its
 * way back to this tenant - so `onSubscribe` saves first and opens Marketplace after (see
 * `FeatureSubscriptionSwitch`). Without `onSubscribe` there is no Subscribe at all.
 *
 * Manage subscription is still a real link in a **new tab**: there is nothing to record
 * before it, and on a SharePoint page leaving would throw away the sign-in and the chosen
 * site. Keeping XSProject open makes "returning" a matter of switching back, which the
 * owner notices.
 */
export default class SubscriptionPrompt extends React.Component<ISubscriptionPromptProps> {
  public render(): React.ReactElement<ISubscriptionPromptProps> {
    const { state } = this.props;

    return (
      <div
        className={`${styles.prompt} ${state === 'error' ? styles.failure : ''}`}
        role={state === 'error' ? 'alert' : 'status'}
        aria-live="polite"
        data-subscription-state={state}
      >
        {this._renderMessage()}
        <div className={styles.actions}>{this._renderActions()}</div>
      </div>
    );
  }

  private _renderMessage(): React.ReactElement {
    const { state, featureName, isAwaitingMarketplace, links } = this.props;

    switch (state) {
      case 'pending':
        return (
          <>
            <p className={styles.message}>Your subscription is being processed.</p>
            <p className={styles.note}>Please try again in a few moments.</p>
          </>
        );

      case 'suspended':
        return (
          <>
            <p className={styles.message}>Your {featureName} subscription is suspended.</p>
            <p className={styles.note}>Manage your subscription to restore access.</p>
          </>
        );

      case 'disabled':
        return (
          <>
            <p className={styles.message}>{featureName} is turned off.</p>
            <p className={styles.note}>
              Your subscription stays active in Microsoft Marketplace until you cancel it there.
            </p>
          </>
        );

      case 'error':
        return (
          <>
            <p className={styles.message}>Unable to verify your subscription.</p>
            <p className={styles.note}>Please try again.</p>
          </>
        );

      default:
        return (
          <>
            <p className={styles.heading}>Subscription required</p>
            <p className={styles.message}>This feature requires an active subscription.</p>
            {/* Back from Marketplace and still not active: most often the purchase has
                not reached the Web API yet, so say so rather than inviting a second one. */}
            {isAwaitingMarketplace && (
              <p className={styles.note}>
                If you have just subscribed, your subscription is being processed. Please try
                again in a few moments.
              </p>
            )}
            {!isNavigableUrl(links.marketplaceOfferUrl) && (
              <p className={styles.note}>
                Subscribing is not available here yet. Please contact your administrator.
              </p>
            )}
          </>
        );
    }
  }

  private _renderActions(): React.ReactNode[] {
    const { state, isAwaitingMarketplace, links, onSubscribe } = this.props;
    const actions: React.ReactNode[] = [];

    if (state === 'inactive' && onSubscribe && isNavigableUrl(links.marketplaceOfferUrl)) {
      actions.push(
        <button key="subscribe" type="button" className={styles.primary} onClick={onSubscribe}>
          Subscribe
        </button>
      );
    }

    if ((state === 'suspended' || state === 'disabled') && isNavigableUrl(links.manageSubscriptionUrl)) {
      actions.push(
        <a
          key="manage"
          className={styles.primary}
          href={links.manageSubscriptionUrl}
          target="_blank"
          rel="noopener noreferrer"
          onClick={this.props.onManage}
        >
          Manage subscription
        </a>
      );
    }

    // A plain inactive answer with nothing in flight has nothing to re-check, nor has a
    // feature the tenant switched off itself; everything else may change on its own once
    // Marketplace and the Web API catch up.
    if ((state !== 'inactive' && state !== 'disabled') || isAwaitingMarketplace) {
      actions.push(
        <button key="check" type="button" className={styles.secondary} onClick={this.props.onCheckAgain}>
          {state === 'error' ? 'Try again' : 'Check again'}
        </button>
      );
    }

    return actions;
  }
}
