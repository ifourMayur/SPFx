import { IFeatureSubscription } from '../../../../models/FeatureSubscription';

export interface IFeatureSubscriptionSwitchState {
  /** True while the API is being asked; the switch is not offered until it first answers. */
  isLoading: boolean;
  /** True while a switch change is being saved. */
  isSaving: boolean;
  /** The value being saved, shown on the switch until the API answers. */
  pendingEnabled?: boolean;
  /** The API's last answer for this feature. Only ever set from a service call. */
  subscription?: IFeatureSubscription;
  /** The subscription could not be read - which is never read as "inactive". */
  isUnverified: boolean;
  /** The last switch change was not saved; the switch shows the stored value again. */
  isSaveFailed: boolean;
  /**
   * True once Marketplace was opened from here, or the user arrived back from it. Makes
   * the switch re-check when the tab regains focus, and explains that a purchase can take
   * a moment to arrive - it changes what is said, never what is allowed.
   */
  isAwaitingMarketplace: boolean;
}
