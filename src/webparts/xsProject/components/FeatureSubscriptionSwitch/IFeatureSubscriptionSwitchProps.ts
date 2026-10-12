import { FeatureCode } from '../../../../models/FeatureSubscription';
import { ISubscriptionService } from '../../../../services/SubscriptionService';

export interface IFeatureSubscriptionSwitchProps {
  /** The paid feature this switch turns on and off. */
  feature: FeatureCode;
  /** Reads and writes the feature's state; also supplies the Marketplace links. */
  subscriptionService: ISubscriptionService;
  /**
   * Whether this user may change the setting, from `canManageFeatureSubscriptions`. A
   * convenience only: `SubscriptionController` refuses other roles itself.
   */
  canManage: boolean;
  /**
   * True when the route was reached from the Marketplace landing page's redirect (see
   * `isMarketplaceReturn`). Skips the cache and words a not-yet-active answer as "being
   * processed". Never makes the feature active by itself.
   */
  isReturningFromMarketplace?: boolean;
}
