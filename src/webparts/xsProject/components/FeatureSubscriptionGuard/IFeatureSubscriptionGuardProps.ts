import * as React from 'react';

import { FeatureCode } from '../../../../models/FeatureSubscription';
import { ISubscriptionService } from '../../../../services/SubscriptionService';

export interface IFeatureSubscriptionGuardProps {
  /** The paid feature whose content this guards. */
  feature: FeatureCode;
  /** Reads the subscription from the Web API; also supplies the Marketplace links. */
  subscriptionService: ISubscriptionService;
  /**
   * True when the route was reached from the Marketplace landing page's redirect (see
   * `isMarketplaceReturn`). Skips the cache and words a not-yet-active answer as "being
   * processed". Never makes the feature active by itself.
   */
  isReturningFromMarketplace?: boolean;
  /** The feature itself, rendered only once the API reports it active. */
  children?: React.ReactNode;
}
