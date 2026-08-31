/**
 * Data contracts for the `Project` resource exposed by the .NET Core Web API,
 * plus the view state shape shared by the project components.
 *
 * Part of the models layer: no imports from `src/services` or `src/components`.
 */

/** Lifecycle state of a project, matching the API's string enum. */
export type ProjectStatus = 'NotStarted' | 'InProgress' | 'OnHold' | 'Completed' | 'Cancelled';

/** A project as returned by `GET {baseUrl}/projects`. */
export interface IProject {
  id: number;
  name: string;
  description?: string;
  status: ProjectStatus;
  /** ISO 8601 date string, as serialized by System.Text.Json. */
  startDate?: string;
  /** ISO 8601 date string, as serialized by System.Text.Json. */
  endDate?: string;
  ownerEmail?: string;
  budget?: number;
}

/** Body sent to `POST {baseUrl}/projects`. */
export interface ICreateProjectRequest {
  name: string;
  description?: string;
  status: ProjectStatus;
  startDate?: string;
  endDate?: string;
  ownerEmail?: string;
  budget?: number;
}

/** Body sent to `PUT {baseUrl}/projects/{id}`; every field is optional. */
export type IUpdateProjectRequest = Partial<ICreateProjectRequest>;

/** Filter applied when listing projects. Serialized into the query string. */
export interface IProjectQuery {
  search?: string;
  status?: ProjectStatus;
}

/**
 * View state used by components that render a list of projects.
 * Kept in the models layer so any component can reuse it.
 */
export interface IProjectListState {
  projects: IProject[];
  isLoading: boolean;
  errorMessage?: string;
}
