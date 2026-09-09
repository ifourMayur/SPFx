import { IBuildingService } from '../../../../services/BuildingService';
import { ILoginService } from '../../../../services/LoginService';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectService } from '../../../../services/ProjectService';
import { ISharePointSiteService } from '../../../../services/SharePointSiteService';

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
  /**
   * Write path of the project form: `POST api/Building`. Called from this component,
   * because it owns the save callback the routed form is handed.
   */
  buildingService: IBuildingService;
  /** Reference data for the project form's dropdowns, passed down to `ProjectAddEdit`. */
  lookupService: ILookupService;
  /** Passed on to the `Login` component, which runs the sign-in handshake. */
  loginService: ILoginService;
  /**
   * Passed on to the `SharePointSites` component, which lists the sites the user may
   * choose between after signing in.
   */
  sharePointSiteService: ISharePointSiteService;
}
