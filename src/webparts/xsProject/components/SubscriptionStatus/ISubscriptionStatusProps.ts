import { IFeatureSubscription } from '../../../../models/FeatureSubscription';

export interface ISubscriptionStatusProps {
  /** An active subscription, as the API reported it. */
  subscription: IFeatureSubscription;
  /** Where the subscription is managed; the action is hidden when it is not configured. */
  manageSubscriptionUrl: string;
}
