import { ISubscriptionService } from '../../../../services/SubscriptionService';

export interface IEdControlProps {
  /** Reads and switches ED Control through `FeatureSubscriptionSwitch`. */
  subscriptionService: ISubscriptionService;
  /** Whether this user may switch ED Control on or off, from `canManageFeatureSubscriptions`. */
  canManage: boolean;
  /** The route was reached from the Marketplace landing page's redirect. */
  isReturningFromMarketplace?: boolean;
}
