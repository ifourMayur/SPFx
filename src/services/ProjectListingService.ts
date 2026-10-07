import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import { IApiEndpoints } from '../models/Environment';
import {
  IPaginationRequest,
  IProjectListPage,
  IProjectListQuery,
  toPaginationRequest,
  toProjectListPage
} from '../models/ProjectListing';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/**
 * The paged read behind the project overview table.
 *
 * Distinct from `IProjectService`, which is the REST-style `api/projects` the dashboard
 * reads, and from `IBuildingService`, which is the form's write path - don't merge them.
 * This one exists because the overview table needs something neither of those exposes:
 * per-phase status for every project, a page at a time.
 *
 * The reference reaches this through `BuildingController.GetAllProjectWithPagination`,
 * which is a proxy for DataTables' form encoding and nothing else. Calling
 * `Dashboard/GetListAllWithPagination` directly removes that hop rather than changing the
 * data, exactly as `LookupService` does for the dropdowns.
 *
 * The phase **columns** are not here: they are `ILookupService.getFolderColumns`, because
 * they are the same `List<LookupItem>` reference data every dropdown is built from. The
 * reference splits them the same way, fetching the folder list in `ProjectLoadData` and
 * the rows through this endpoint.
 */
export interface IProjectListingService {
  /**
   * One page of projects, plus how many there are in total.
   *
   * The total comes back stamped on the rows rather than beside them - see
   * `models/ProjectListing.ts` - so a caller must not compute it from `rows.length`.
   */
  getProjectPage(query: IProjectListQuery): Promise<IProjectListPage>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'ProjectListingService';

/** The action on the dashboard controller. Its route lives in `src/config/environment.ts`. */
const LIST_ACTION: string = 'GetListAllWithPagination';

/**
 * Default {@link IProjectListingService}, built on `BaseApiService`.
 *
 * Three steps, all of them thin: build the `PaginationModel`, POST it, unwrap the
 * `ResponseDetail` envelope. Both the request shape and the response mapping are pure
 * functions in the models layer, and the POST-and-unwrap is `BaseApiService.postData`, so
 * this class stays a readable mapping of one business operation onto one route.
 */
export class ProjectListingService extends BaseApiService implements IProjectListingService {
  private readonly _endpoints: IApiEndpoints;

  /**
   * @param apiService - HTTP gateway used for every call. It supplies `baseUrl` and the
   *   `Authorization: Bearer` header carrying the application JWT that `LoginService`
   *   published on sign-in - which is also where the API reads the caller's `UserId` and
   *   `UserRoleId` from, so neither is sent in the body.
   * @param endpoints - controller routes; defaults to `api.endpoints` from the environment
   *   configuration, and can be overridden in tests.
   */
  public constructor(apiService: IApiService, endpoints: IApiEndpoints = getEnvironment().api.endpoints) {
    super(apiService, LOG_SOURCE);
    this._endpoints = endpoints;
  }

  public async getProjectPage(query: IProjectListQuery): Promise<IProjectListPage> {
    const relativeUrl: string = this.route(this._endpoints.dashboard, LIST_ACTION);
    const url: string = this.resolveUrl(relativeUrl);
    const request: IPaginationRequest = toPaginationRequest(query);

    try {
      // An HTTP 200 proves nothing here: a refusal arrives as `success: false` inside the
      // envelope, and passing that payload through would render an empty table.
      const page: IProjectListPage = toProjectListPage(
        await this.postData<unknown, IPaginationRequest>(relativeUrl, request, {
          operation: 'the project list'
        })
      );

      // Logged unconditionally, like `LookupService`: "the table is empty" has three
      // causes - the call failed, the API refused inside a 200, or the tenant genuinely
      // has no projects - and nothing else distinguishes them from the outside.
      Log.info(
        LOG_SOURCE,
        `${url} -> ${page.rows.length} row(s) of ${page.recordsTotal} for page ${query.page}.`
      );

      return page;
    } catch (error) {
      this.logFailure(error, `${url} failed while loading the project list.`);

      throw error;
    }
  }
}
