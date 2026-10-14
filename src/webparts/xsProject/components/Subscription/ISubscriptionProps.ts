import { ISubscriptionService } from '../../../../services/SubscriptionService';

export interface ISubscriptionProps {
  /**
   * Reads, switches and activates the paid features. Also supplies the Marketplace links,
   * which every switch reads from it.
   */
  subscriptionService: ISubscriptionService;
  /**
   * Whether this user may switch features on or off and buy them, from
   * `canManageFeatureSubscriptions`. A convenience only: the Web API refuses other roles
   * itself.
   */
  canManage: boolean;
  /**
   * Whether this host can buy: false inside Teams, where the page Marketplace would send the
   * buyer back to only works inside Teams. Handed to every switch unchanged.
   */
  isPurchaseAvailable: boolean;
  /**
   * The one-time purchase token the Marketplace landing page sent the buyer back with
   * (`readMarketplaceToken`). Handed to the Web API once; never shown, never logged.
   */
  marketplaceToken?: string;
  /** The route was reached from the Marketplace landing page's redirect. */
  isReturningFromMarketplace: boolean;
  /**
   * Called once the API has answered a token - activated or refused - so the router can
   * take the spent token out of the URL.
   */
  onTokenConsumed: () => void;
}
