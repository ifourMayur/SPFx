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
  /**
   * Renders the API response under the success message. Defaults to `true` as a
   * development aid; set it to `false` for production, where the payload carries the
   * application token and refresh token the API issued.
   */
  showResponse?: boolean;
  /** Called with the Web API response after a successful sign-in. */
  onLoginSucceeded?: (response: ILoginResponse) => void;
  /** Called with the classified failure when sign-in does not complete. */
  onLoginFailed?: (failure: ILoginFailure) => void;
}
