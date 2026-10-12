import { IFeatureSubscription } from '../../../../models/FeatureSubscription';

export interface IFeatureSubscriptionGuardState {
  /** True while the API is being asked; no action is offered until it has answered. */
  isLoading: boolean;
  /** The API's last answer for this feature. Only ever set from a service call. */
  subscription?: IFeatureSubscription;
  /** The subscription could not be verified - which is never read as "inactive". */
  isUnverified: boolean;
  /**
   * True once the user has gone to Microsoft Marketplace from here, or arrived back from
   * it. Makes the guard re-check when the tab regains focus and explains that a purchase
   * can take a moment to arrive - it changes what is said, never what is allowed.
   */
  isAwaitingMarketplace: boolean;
}
