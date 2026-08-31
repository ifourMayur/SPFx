import { AppView } from '../../../../models/Navigation';

export interface IXsProjectState {
  /** Set from `Login`'s `onLoginSucceeded`/`onLoginFailed` callbacks; gates the shared `Menu`. */
  isAuthenticated: boolean;
  /** View the shared `Menu` currently has selected. */
  activeView: AppView;
}
