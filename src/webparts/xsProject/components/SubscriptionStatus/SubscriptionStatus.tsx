import * as React from 'react';

import styles from './SubscriptionStatus.module.scss';
import type { ISubscriptionStatusProps } from './ISubscriptionStatusProps';
import { isNavigableUrl, toPlanLabel } from '../../../../models/FeatureSubscription';

/**
 * The one-line summary shown above a paid feature the tenant has bought: its plan and
 * status as the Web API reported them, and where to change them.
 *
 * Plans are chosen, changed and cancelled in Microsoft Marketplace, never here, so "Manage
 * subscription" is a link out and the plan is display text only. A plan the API did not
 * name is left out rather than guessed.
 */
export default class SubscriptionStatus extends React.Component<ISubscriptionStatusProps> {
  public render(): React.ReactElement<ISubscriptionStatusProps> {
    const { subscription, manageSubscriptionUrl } = this.props;
    const plan: string | undefined = toPlanLabel(subscription);

    return (
      <div className={styles.subscriptionStatus} data-subscription-state={subscription.state}>
        {plan && (
          <span className={styles.item}>
            Plan: <strong>{plan}</strong>
          </span>
        )}
        <span className={styles.item}>
          Status: <strong>{subscription.subscriptionStatus || 'Active'}</strong>
        </span>
        {isNavigableUrl(manageSubscriptionUrl) && (
          <a className={styles.manage} href={manageSubscriptionUrl} target="_blank" rel="noopener noreferrer">
            Manage subscription
          </a>
        )}
      </div>
    );
  }
}
