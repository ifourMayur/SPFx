import {
  ISaasActivationResult,
  ISaasEntitlements,
  ISaasErrorInfo,
  ISaasServiceRow,
  ISaasSubscription
} from '../../../../models/Saas';

/**
 * The banner above the list, about activating a Marketplace subscription.
 *
 * `activating` - the call is running; `done` - the server answered (success or a warning,
 * by `result.state`); `failed` - the call failed, with "Try again" when repeating it could
 * help; `returned` - the buyer came back from the Marketplace without a purchase to activate.
 */
export type ActivationBanner =
  | { kind: 'activating' }
  | { kind: 'done'; result: ISaasActivationResult }
  | { kind: 'failed'; subscriptionId: string; error: ISaasErrorInfo; canRetry: boolean }
  | { kind: 'returned' };

export interface ISubscriptionState {
  isLoading: boolean;
  /** A manual Refresh is running (the list stays on screen meanwhile). */
  isRefreshing: boolean;
  entitlements?: ISaasEntitlements;
  rows: ISaasServiceRow[];
  /** The entitlements could not be loaded - nothing reliable to show. */
  loadError?: ISaasErrorInfo;
  /** Only the catalog failed - the rows come from the entitlements alone. */
  catalogError?: ISaasErrorInfo;
  /** The tenant's Marketplace subscriptions. */
  subscriptions: ISaasSubscription[];
  /** Only the subscriptions failed - the services are still shown. */
  subscriptionsError?: ISaasErrorInfo;
  /** The service whose purchase is being started. */
  enablingKey?: string;
  /** Starting a purchase or refreshing failed. */
  actionError?: ISaasErrorInfo;
  /** The Marketplace subscription being activated; every action is disabled meanwhile. */
  activatingId?: string;
  banner?: ActivationBanner;
  /** The service the purchase was for, or the one a locked `FeatureGate` sent the user here for. */
  highlightKey?: string;
}
