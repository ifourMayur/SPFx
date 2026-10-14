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
   * Whether this host can buy: false inside Teams, where the page Marketplace would send the
   * buyer back to only works inside Teams. Without it the switch still records the choice,
   * but opens no tab, sends no return URL and offers no Subscribe - it says to subscribe
   * from SharePoint instead.
   *
   * Required rather than defaulted, so that every place that mounts a switch has to say
   * which host it is in: a forgotten `true` in Teams would send a buyer to Marketplace with
   * no way back to activate the purchase.
   */
  isPurchaseAvailable: boolean;
  /**
   * True when the route was reached from the Marketplace landing page's redirect (see
   * `isMarketplaceReturn`). Skips the cache and words a not-yet-active answer as "being
   * processed". Never makes the feature active by itself.
   */
  isReturningFromMarketplace?: boolean;
}
