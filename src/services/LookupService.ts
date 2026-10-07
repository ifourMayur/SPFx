import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import { IFolderSelection, ILookupOption } from '../models/Building';
import { IApiEndpoints } from '../models/Environment';
import { FolderLevel } from '../models/FolderTree';
import { toLookupOptions } from '../models/Lookup';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/**
 * Reference data the project form binds its dropdowns to.
 *
 * Every method maps onto the endpoint the reference application already uses for that same
 * dropdown, so the React form is fed by exactly the same data as the Razor view:
 *
 * | Dropdown | Endpoint | Where the reference calls it |
 * | --- | --- | --- |
 * | Project Template | `ProjectTemplate/GetTemplateList` | `BuildingController.AddEdit` GET, into `ViewBag.ProjecTemplateList` |
 * | User List | `Users/GetAllList` | `BuildingController.AddEdit` GET, into `ViewBag.AccessRightsUserList` |
 * | Spatial Breakdown | `SpatialBreakdown/GetAllLookup` | `BuildingController.AddEdit` GET, into `ViewBag.spatialBreakdownViewModelsList` |
 * | Tender Templates | `TenderTemplate/GetTenderTemplateLookupItems` | `getTenderTemplate()` -> `Building/GetTenderTemplates` |
 * | Supplier | `Issuer/GetIssuerList` | `getSupplierList()` -> `Building/GetSupplierList` |
 * | Folder cascade | `Building/GetProjectFolders` and its three `…SubFolders` siblings | `getMOMProjectFolderList()` and friends |
 *
 * The MVC controller is only a proxy: its `DoActionForGet<List<LookupItem>>(query, route)`
 * calls the Web API route named above and hands back `ResponseDetail.data`. Calling the
 * Web API directly from here removes that hop rather than changing the data.
 *
 * The ED Controls tag list, ED Controls projects and KYP plannings are deliberately absent:
 * their sections are out of scope for this screen.
 */
export interface ILookupService {
  getProjectTemplates(): Promise<ILookupOption[]>;

  getUsers(): Promise<ILookupOption[]>;

  getSpatialBreakdowns(): Promise<ILookupOption[]>;

  getTenderTemplates(): Promise<ILookupOption[]>;

  getSuppliers(): Promise<ILookupOption[]>;

  /**
   * Top-level folders of one project.
   *
   * Every folder endpoint is scoped to a project, because a project's folder tree is built
   * from its template when the project is created. A project that does not exist yet
   * therefore has no folders, and the reference has the same gap: it passes
   * `projectId: $('#Id').val()`, which is `0` while adding.
   */
  getProjectFolders(projectId: number): Promise<ILookupOption[]>;

  getProjectSubFolders(projectId: number, folderId: number): Promise<ILookupOption[]>;

  getProjectSubSubFolders(
    projectId: number,
    folderId: number,
    subFolderId: number
  ): Promise<ILookupOption[]>;

  getProjectSubSubSubFolders(
    projectId: number,
    folderId: number,
    subFolderId: number,
    subSubFolderId: number
  ): Promise<ILookupOption[]>;

  /**
   * One level of the folder cascade, given what is selected above it.
   *
   * The single entry point the form uses, so the component never decides which endpoint a
   * level maps to or which ancestor ids it takes. Each level's endpoint takes the ids of
   * every level above it; they are strings in the form (a `<select>` value always is) and
   * integers on the wire, and this is where one becomes the other.
   *
   * @param projectId - the project being edited. Callers skip the call while adding
   *   (`0`), since a project that does not exist yet has no folders.
   * @param level - the level to load.
   * @param selection - the form's current cascade selection; only the levels above
   *   `level` are read.
   */
  getFolderLevel(projectId: number, level: FolderLevel, selection: IFolderSelection): Promise<ILookupOption[]>;

  /**
   * Every active folder in the system except BIM - the project overview table's phase
   * columns.
   *
   * The odd one out on this interface: not a dropdown, and not scoped to a project. It
   * lives here because it is the same `List<LookupItem>` reference data every other method
   * fetches, through the same unwrap and the same mapper, and `BuildingController.ProjectLoadData`
   * fetches it for exactly this purpose - it is the model of `_ProjectListingScreen`,
   * whose `<thead>` is built by looping over it.
   *
   * The ids are what `ProjectListing.resolveCell` matches each row's `lstDashboard` cell
   * on, so a column and its cells agree without either knowing the other's ordering.
   */
  getFolderColumns(): Promise<ILookupOption[]>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'LookupService';

/**
 * Controller actions, one per dropdown.
 *
 * Held here while the controller routes they hang off live in
 * `src/config/environment.ts`, exactly as `LoginService` splits `SPFX_LOGIN_ACTION` from
 * the configured `authenticate` route. `ApiService.resolveUrl` then supplies `baseUrl`, so
 * a complete endpoint is always `baseUrl` + controller + action.
 */
const ACTIONS = {
  projectTemplates: 'GetTemplateList',
  users: 'GetAllList',
  spatialBreakdowns: 'GetAllLookup',
  tenderTemplates: 'GetTenderTemplateLookupItems',
  suppliers: 'GetIssuerList',
  projectFolders: 'GetProjectFolders',
  projectSubFolders: 'GetProjectSubFolders',
  projectSubSubFolders: 'GetProjectSubSubFolders',
  projectSubSubSubFolders: 'GetProjectSubSubSubFolders',
  folderColumns: 'GetFolderWithoutBIM'
};

/**
 * Default {@link ILookupService}, built on `BaseApiService`.
 *
 * Each call is the same three steps: GET the route, unwrap the `ResponseDetail` envelope
 * (which reports failure with `success: false` on an HTTP 200, so the status alone proves
 * nothing), then map `LookupItem` onto `<option>` values. The first two are
 * `BaseApiService.getData` and the mapping is a pure function in the models layer, so this
 * class stays a thin, readable mapping of business operations onto routes.
 */
export class LookupService extends BaseApiService implements ILookupService {
  private readonly _endpoints: IApiEndpoints;

  /**
   * @param apiService - HTTP gateway used for every call. It supplies `baseUrl` and the
   *   `Authorization: Bearer` header carrying the application JWT that `LoginService`
   *   published on sign-in, so these `[Authorize]`d endpoints accept the call.
   * @param endpoints - controller routes; defaults to `api.endpoints` from the environment
   *   configuration, and can be overridden in tests. Mirrors `LoginService`'s `endpoint`
   *   parameter.
   */
  public constructor(apiService: IApiService, endpoints: IApiEndpoints = getEnvironment().api.endpoints) {
    super(apiService, LOG_SOURCE);
    this._endpoints = endpoints;
  }

  /** Project Template dropdown. `ViewBag.ProjecTemplateList` in the reference. */
  public async getProjectTemplates(): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.projectTemplates, ACTIONS.projectTemplates, 'the project templates');
  }

  /**
   * User List multi-select. `ViewBag.AccessRightsUserList` in the reference.
   *
   * The API scopes this to the caller and **excludes both the caller and every Admin**
   * (`UsersRepository.GetAllList`), so an empty result is normal on a tenant where the
   * signed-in user is the only non-admin account - it is not a failed call.
   */
  public async getUsers(): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.users, ACTIONS.users, 'the user list');
  }

  /**
   * Spatial Breakdown multi-select. `ViewBag.spatialBreakdownViewModelsList` in the
   * reference. Hidden altogether for the McD domain group.
   *
   * The API formats these as `Name (Code)`, so the code arrives inside `name`.
   */
  public async getSpatialBreakdowns(): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.spatialBreakdowns, ACTIONS.spatialBreakdowns, 'the spatial breakdowns');
  }

  /** Tender Templates dropdown, revealed by the Tender toggle. */
  public async getTenderTemplates(): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.tenderTemplates, ACTIONS.tenderTemplates, 'the tender templates');
  }

  /** Supplier dropdown beneath the MOM folder cascade. "Issuers" on the API side. */
  public async getSuppliers(): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.suppliers, ACTIONS.suppliers, 'the suppliers');
  }

  public async getProjectFolders(projectId: number): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.building, ACTIONS.projectFolders, 'the folders', { projectId });
  }

  /** Second cascade level: the children of `folderId`. */
  public async getProjectSubFolders(projectId: number, folderId: number): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.building, ACTIONS.projectSubFolders, 'the sub folders', { projectId, folderId });
  }

  /** Third cascade level. */
  public async getProjectSubSubFolders(
    projectId: number,
    folderId: number,
    subFolderId: number
  ): Promise<ILookupOption[]> {
    // The API names this parameter `subfolderId`, lower-cased, unlike `folderId`.
    return this._getLookup(this._endpoints.building, ACTIONS.projectSubSubFolders, 'the sub sub folders', {
      projectId,
      folderId,
      subfolderId: subFolderId
    });
  }

  /** Fourth cascade level, offered only when `isGrandChildFolderEnabled`. */
  public async getProjectSubSubSubFolders(
    projectId: number,
    folderId: number,
    subFolderId: number,
    subSubFolderId: number
  ): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.building, ACTIONS.projectSubSubSubFolders, 'the sub sub sub folders', {
      projectId,
      folderId,
      subfolderId: subFolderId,
      subsubfolderId: subSubFolderId
    });
  }

  public async getFolderLevel(
    projectId: number,
    level: FolderLevel,
    selection: IFolderSelection
  ): Promise<ILookupOption[]> {
    const folderId: number = Number(selection.folderId);
    const subFolderId: number = Number(selection.subFolderId);

    switch (level) {
      case 'folderId':
        return this.getProjectFolders(projectId);
      case 'subFolderId':
        return this.getProjectSubFolders(projectId, folderId);
      case 'subSubFolderId':
        return this.getProjectSubSubFolders(projectId, folderId, subFolderId);
      default:
        return this.getProjectSubSubSubFolders(projectId, folderId, subFolderId, Number(selection.subSubFolderId));
    }
  }

  /** Phase columns of the project overview table. */
  public async getFolderColumns(): Promise<ILookupOption[]> {
    return this._getLookup(this._endpoints.folder, ACTIONS.folderColumns, 'the folder columns');
  }

  /**
   * GET the route, unwrap the envelope, map to options - and say what happened.
   *
   * The outcome is logged unconditionally rather than behind `enableVerboseLogging`,
   * because "the dropdown is empty" has three very different causes that are otherwise
   * indistinguishable from the outside: the call failed, the API reported a failure inside
   * an HTTP 200, or it genuinely returned no rows. The log names which one, and for which
   * endpoint, in any environment.
   *
   * @param controller - configured controller route, e.g. `endpoints.users`.
   * @param action - the action on that controller, e.g. `GetAllList`.
   * @param operation - human-readable name of what is being fetched, used in the log line
   *   and in the error message when the API reports a failure without one of its own.
   * @param query - query-string parameters, used only by the folder cascade.
   */
  private async _getLookup(
    controller: string,
    action: string,
    operation: string,
    query?: { [name: string]: number }
  ): Promise<ILookupOption[]> {
    const relativeUrl: string = this.route(controller, action);
    const url: string = this.resolveUrl(relativeUrl, query);

    try {
      const options: ILookupOption[] = toLookupOptions(
        await this.getData<unknown>(relativeUrl, query, { operation })
      );

      Log.info(LOG_SOURCE, `${url} -> ${options.length} option(s) for ${operation}.`);

      return options;
    } catch (error) {
      this.logFailure(error, `${url} failed while loading ${operation}.`);

      throw error;
    }
  }
}
