import { ILoginService } from '../../../../services/LoginService';
import { IProjectService } from '../../../../services/ProjectService';

export interface IXsProjectProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  /**
   * Injected by the web part through `ServiceFactory`, ready for this component tree to
   * call. Keeps API access out of the components themselves.
   */
  projectService: IProjectService;
  /** Passed on to the `Login` component, which runs the sign-in handshake. */
  loginService: ILoginService;
}
