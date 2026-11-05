import { ISaasService } from '../../../../services/SaasService';

export interface ISubscriptionProps {
  saasService: ISaasService;
  /**
   * The route's query string, read once on mount: the return from a Marketplace purchase
   * (`?saas=returned&service=TENDER&subscriptionId={guid}`), or the service a locked
   * `FeatureGate` asked for (`?service=TENDER`).
   */
  query: string;
  /** Removes the query from the URL once it has been read, so a reload doesn't activate again. */
  onClearQuery: () => void;
  /** The page's current address, the base of the purchase return URL. Defaults to `window.location.href`. */
  getCurrentUrl?: () => string;
  /**
   * Leaves for the Marketplace offer, in the same tab. Defaults to
   * `window.location.assign`; injectable because jsdom cannot spy on `location`.
   */
  navigateToUrl?: (url: string) => void;
}
