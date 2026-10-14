/** Where handing a Marketplace purchase token to the Web API stands. */
export type ActivationStatus = 'activating' | 'activated' | 'refused';

export interface ISubscriptionState {
  /**
   * Bumped each time an activation settles. The switches are keyed on it, so they mount
   * afresh and read again rather than keep the answer they had before the purchase.
   */
  activationGeneration: number;
  /** Absent until a token is handed to the API on this page. */
  activationStatus?: ActivationStatus;
  /** What is said about the activation once the API has answered. */
  activationMessage?: string;
}
