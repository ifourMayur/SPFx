import { ISubscriptionLinks, SubscriptionState } from '../../../../models/FeatureSubscription';

/**
 * Why the feature is locked: the API's answer, `disabled` when it is paid for but switched
 * off, or `error` when there was no answer to read. `active` never reaches this component.
 */
export type LockedSubscriptionState = Exclude<SubscriptionState, 'active'> | 'disabled' | 'error';

export interface ISubscriptionPromptProps {
  /** The feature's display name, e.g. "ED Control". */
  featureName: string;
  state: LockedSubscriptionState;
  links: ISubscriptionLinks;
  /** The user has been to Marketplace from here, or has just come back from it. */
  isAwaitingMarketplace: boolean;
  /** The user followed the Subscribe link; the navigation itself is the link's own. */
  onSubscribe: () => void;
  /** Asks the API again, bypassing the cache - "Check again" and "Try again". */
  onCheckAgain: () => void;
}
