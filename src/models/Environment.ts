import { IRequestHeaders } from './ApiResponse';

/**
 * Contracts for the solution-wide environment configuration.
 *
 * Only the shapes live here; the actual values live in `src/config/environment.ts`.
 * Part of the models layer: no imports from `src/config`, `src/services` or `src/components`.
 */

/** Deployment environments this solution is built for. */
export type EnvironmentName = 'local' | 'development' | 'test' | 'production';

/** Route names of the .NET Core Web API controllers, relative to the API base URL. */
export interface IApiEndpoints {
  readonly projects: string;
  /** `AuthenticateController` route; the sign-in actions hang off it, e.g. `SPFxLogin`. */
  readonly authenticate: string;
}

/** Everything the API layer needs to talk to the Web API in a given environment. */
export interface IApiEnvironmentSettings {
  /** Absolute base address of the .NET Core Web API, e.g. `https://contoso-api.azurewebsites.net/api`. */
  readonly baseUrl: string;
  /**
   * Application ID URI (or client ID) of the Entra ID app registration that protects the API.
   * Leave undefined to call the API anonymously.
   */
  readonly resourceUri?: string;
  /** Abandon a request after this many milliseconds. Use `0` to disable the timeout. */
  readonly requestTimeoutMs: number;
  /** Headers applied to every request, e.g. an API version header. */
  readonly defaultHeaders?: IRequestHeaders;
  readonly endpoints: IApiEndpoints;
}

/** Everything the sign-in handshake needs in a given environment. */
export interface IAuthEnvironmentSettings {
  /**
   * Fallback for `domainUrl` in the SPFxLogin request body, used only when
   * `context.pageContext.web.absoluteUrl` is unavailable. Normally `ServiceFactory`
   * resolves `domainUrl` dynamically from the current SharePoint site instead of this
   * value - see `ServiceFactory._resolveDomainUrl`.
   */
  readonly domainUrl: string;
  /**
   * Resource the SPFx access token is requested for, for example
   * `https://graph.microsoft.com` or the Application ID URI of the Entra ID app
   * registration that fronts the Web API.
   */
  readonly tokenResourceUri: string;
}

/** Presentation defaults shared by the components. */
export interface IUiEnvironmentSettings {
  /** Default heading for the project list; a web part instance may override it. */
  readonly projectListTitle: string;
  /** Upper bound on how many projects are rendered at once. */
  readonly maxProjectsToDisplay: number;
}

/** The resolved configuration for the environment the solution is currently running in. */
export interface IEnvironmentConfiguration {
  readonly name: EnvironmentName;
  readonly api: IApiEnvironmentSettings;
  readonly auth: IAuthEnvironmentSettings;
  readonly ui: IUiEnvironmentSettings;
  /** When true, the API layer writes request/response diagnostics to the SPFx log. */
  readonly enableVerboseLogging: boolean;
}

/** Hints used to decide which environment the solution is running in. */
export interface IEnvironmentResolutionOptions {
  /** Forces a specific environment and skips all detection. */
  environmentName?: EnvironmentName;
  /** Host name of the current site, e.g. `contoso.sharepoint.com`. */
  hostName?: string;
  /** Pass `context.isServedFromLocalhost` so `heft start` always resolves to `local`. */
  isServedFromLocalhost?: boolean;
}
