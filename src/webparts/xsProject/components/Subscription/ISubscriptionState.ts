/**
 * Where handing a Marketplace purchase token to the Web API stands. `notPermitted` means it
 * was not handed over at all: this user may not buy.
 */
export type ActivationStatus = 'activating' | 'activated' | 'refused' | 'notPermitted';

export interface ISubscriptionState {
  /**
   * Bumped each time an activation settles. The switches are keyed on it, so they mount
   * afresh and read again rather than keep the answer they had before the purchase.
   */
  activationGeneration: number;
  /** Absent until a token is handed to the API on this page, or a failed one is kept. */
  activationStatus?: ActivationStatus;
  /** What is said about the activation once the API has answered. */
  activationMessage?: string;
  /**
   * A failed activation's token is kept, so "Try again" is offered. Only the flag lives in
   * state; the token itself never does.
   */
  canRetry?: boolean;
}
