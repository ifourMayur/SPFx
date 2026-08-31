import { getEnvironment } from '../config/environment';
import { IQueryParameters } from '../models/ApiResponse';
import { ICreateProjectRequest, IProject, IProjectQuery, IUpdateProjectRequest } from '../models/Project';
import { IApiService } from './ApiService';

/**
 * Business operations available for the `Project` resource.
 *
 * Components depend on this interface only, which keeps them free of HTTP concerns
 * and makes them easy to test with a stub implementation.
 */
export interface IProjectService {
  getProjects(query?: IProjectQuery): Promise<IProject[]>;

  getProjectById(id: number): Promise<IProject>;

  createProject(request: ICreateProjectRequest): Promise<IProject>;

  updateProject(id: number, request: IUpdateProjectRequest): Promise<IProject>;

  deleteProject(id: number): Promise<void>;
}

/**
 * Maps `IProjectService` operations onto the REST endpoints exposed by the
 * `ProjectsController` of the .NET Core Web API.
 */
export class ProjectService implements IProjectService {
  private readonly _apiService: IApiService;
  private readonly _endpoint: string;

  /**
   * @param apiService - HTTP gateway used for every call.
   * @param endpoint - controller route; defaults to `api.endpoints.projects` from the
   *   environment configuration, and can be overridden in tests.
   */
  public constructor(apiService: IApiService, endpoint: string = getEnvironment().api.endpoints.projects) {
    this._apiService = apiService;
    this._endpoint = endpoint;
  }

  public async getProjects(query?: IProjectQuery): Promise<IProject[]> {
    const projects: IProject[] = await this._apiService.get<IProject[]>(
      this._endpoint,
      this._toQueryParameters(query)
    );

    // Guard against an API that returns an empty body instead of an empty collection.
    return projects || [];
  }

  public async getProjectById(id: number): Promise<IProject> {
    return this._apiService.get<IProject>(`${this._endpoint}/${id}`);
  }

  public async createProject(request: ICreateProjectRequest): Promise<IProject> {
    return this._apiService.post<IProject, ICreateProjectRequest>(this._endpoint, request);
  }

  public async updateProject(id: number, request: IUpdateProjectRequest): Promise<IProject> {
    return this._apiService.put<IProject, IUpdateProjectRequest>(`${this._endpoint}/${id}`, request);
  }

  public async deleteProject(id: number): Promise<void> {
    return this._apiService.delete(`${this._endpoint}/${id}`);
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
