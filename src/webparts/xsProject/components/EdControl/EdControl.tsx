import * as React from 'react';

import styles from './EdControl.module.scss';
import type { IEdControlProps } from './IEdControlProps';
import FeatureSubscriptionSwitch from '../FeatureSubscriptionSwitch/FeatureSubscriptionSwitch';
import { FEATURE_NAMES } from '../../../../models/FeatureSubscription';

/**
 * The ED Control page: where the tenant switches ED Control - a paid feature, sold through
 * Microsoft Marketplace - on and off.
 *
 * Everything below the title is `FeatureSubscriptionSwitch`. Switching on records the
 * choice and opens Marketplace to buy it; ED Control is only usable once the Web API
 * reports the purchase. The reference's own ED Control screen,
 * `Views/Building/Audits.cshtml`, is per project and is not ported, so this page has no
 * ED Control content of its own - only the setting.
 */
export default class EdControl extends React.Component<IEdControlProps> {
  public render(): React.ReactElement<IEdControlProps> {
    const { subscriptionService, canManage, isReturningFromMarketplace } = this.props;

    return (
      <section className={styles.edControl}>
        <h3 className={styles.title}>{FEATURE_NAMES.ED_CONTROL}</h3>

        <FeatureSubscriptionSwitch
          feature="ED_CONTROL"
          subscriptionService={subscriptionService}
          canManage={canManage}
          isReturningFromMarketplace={isReturningFromMarketplace}
        />
      </section>
    );
  }
}
