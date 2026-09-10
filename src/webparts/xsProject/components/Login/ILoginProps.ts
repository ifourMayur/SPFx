import { ILoginFailure, ILoginResponse } from '../../../../models/Auth';
import { ILoginService } from '../../../../services/LoginService';

export interface ILoginProps {
  /**
   * Injected by the web part through `ServiceFactory`, ready for this component to call.
   * Keeps token acquisition and API access out of the component itself.
   */
  loginService: ILoginService;
  /** Runs the sign-in handshake as soon as the component mounts. Defaults to `true`. */
  autoLogin?: boolean;
  /**
   * Overrides `auth.domainUrl` from `src/config/environment.ts` for this instance.
   * Leave unset to use the configured value.
   */
  domainUrl?: string;
  /** Called with the Web API response after a successful sign-in. */
  onLoginSucceeded?: (response: ILoginResponse) => void;
  /** Called with the classified failure when sign-in does not complete. */
  onLoginFailed?: (failure: ILoginFailure) => void;
}
