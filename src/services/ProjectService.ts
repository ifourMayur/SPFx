import { getEnvironment } from '../config/environment';
import { IQueryParameters } from '../models/ApiResponse';
import { IProject, IProjectQuery } from '../models/Project';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'ProjectService';

/**
 * Business operations available for the `Project` resource.
 *
 * Components depend on this interface only, which keeps them free of HTTP concerns
 * and makes them easy to test with a stub implementation.
 */
export interface IProjectService {
  getProjects(query?: IProjectQuery): Promise<IProject[]>;
}

/**
 * Maps `IProjectService` operations onto the REST endpoints exposed by the
 * `ProjectsController` of the .NET Core Web API.
 *
 * The one Web API service that does not unwrap the `ResponseDetail` envelope: it is written
 * against a plain REST controller answering bare resources, so it calls `apiService`
 * directly rather than `getData`/`postData`.
 */
export class ProjectService extends BaseApiService implements IProjectService {
  private readonly _endpoint: string;

  /**
   * @param apiService - HTTP gateway used for every call.
   * @param endpoint - controller route; defaults to `api.endpoints.projects` from the
   *   environment configuration, and can be overridden in tests.
   */
  public constructor(apiService: IApiService, endpoint: string = getEnvironment().api.endpoints.projects) {
    super(apiService, LOG_SOURCE);
    this._endpoint = endpoint;
  }

  public async getProjects(query?: IProjectQuery): Promise<IProject[]> {
    const projects: IProject[] = await this.apiService.get<IProject[]>(
      this._endpoint,
      this._toQueryParameters(query)
    );

    // Guard against an API that returns an empty body instead of an empty collection.
    return projects || [];
  }

  private _toQueryParameters(query?: IProjectQuery): IQueryParameters | undefined {
    if (!query) {
      return undefined;
    }

    return {
      search: query.search,
      status: query.status
    };
  }
}
