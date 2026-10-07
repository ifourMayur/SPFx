import { IBuildingService } from '../../../../services/BuildingService';
import { ILoginService } from '../../../../services/LoginService';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectListingService } from '../../../../services/ProjectListingService';
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
   * The paged read behind the project overview table:
   * `Dashboard/GetListAllWithPagination`. Separate from `projectService`, which is the
   * REST-style `api/projects` the dashboard reads - only this one exposes per-phase status.
   */
  projectListingService: IProjectListingService;
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
