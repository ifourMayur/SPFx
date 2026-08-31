/**
 * Transport-agnostic data contracts used when talking to the .NET Core Web API.
 *
 * This file is part of the bottom layer of the dependency flow:
 *
 *   models  ->  services  ->  components
 *
 * Models MUST NOT import anything from `src/services` or `src/components`, and they
 * intentionally avoid importing SPFx packages so that they stay reusable and testable.
 */

/** HTTP verbs supported by the API layer. */
export type HttpMethod = 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';

/** Query string values appended to a request URL. `undefined` values are skipped. */
export interface IQueryParameters {
  [name: string]: string | number | boolean | undefined;
}

/** Additional request headers. */
export interface IRequestHeaders {
  [name: string]: string;
}

/**
 * A single API call description.
 *
 * @typeParam TBody - shape of the JSON request body, or `undefined` when there is none.
 */
export interface IApiRequest<TBody = undefined> {
  method: HttpMethod;
  /** Path relative to {@link IApiServiceConfiguration.baseUrl}, e.g. `projects/12`. */
  relativeUrl: string;
  query?: IQueryParameters;
  headers?: IRequestHeaders;
  body?: TBody;
}

/**
 * A successful API call result.
 *
 * @typeParam TData - shape of the deserialized response payload.
 */
export interface IApiResponse<TData> {
  ok: boolean;
  status: number;
  statusText: string;
  data: TData;
}

/**
 * ASP.NET Core `ProblemDetails` / `ValidationProblemDetails` payload
 * (RFC 7807), returned by the Web API for failed requests.
 */
export interface IProblemDetails {
  type?: string;
  title?: string;
  status?: number;
  detail?: string;
  instance?: string;
  traceId?: string;
  /** Field name -> validation messages, produced by `ValidationProblemDetails`. */
  errors?: { [field: string]: string[] };
  /**
   * Not part of RFC 7807: the `AuthenticateController` reports failures as
   * `{ "message": "..." }` rather than as `ProblemDetails`, so the reason is read from
   * here too instead of being replaced by a generic status message.
   */
  message?: string;
}

/**
 * Per-instance overrides for the API service.
 *
 * Every field is optional: the API service reads its settings from the environment
 * configuration (`src/config/environment.ts`) and only consults these values to
 * replace individual settings for one component instance.
 */
export interface IApiServiceConfiguration {
  /**
   * Absolute base address of the Web API, e.g. `https://contoso-api.azurewebsites.net/api`.
   * Omit to use `api.baseUrl` from the environment configuration.
   */
  baseUrl?: string;
  /**
   * Application ID URI (or client ID) of the Entra ID app registration that fronts the API.
   * When omitted, the API is called anonymously with the SPFx `HttpClient`.
   */
  resourceUri?: string;
  /** Headers applied to every request issued through the service. */
  defaultHeaders?: IRequestHeaders;
  /** Abandon a request after this many milliseconds. Omit or use `0` to disable. */
  timeoutMs?: number;
  /** When true, requests and failures are written to the SPFx log. */
  enableVerboseLogging?: boolean;
}
