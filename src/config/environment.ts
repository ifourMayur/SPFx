import {
  EnvironmentName,
  IApiEndpoints,
  IAuthEnvironmentSettings,
  IEnvironmentConfiguration,
  IEnvironmentResolutionOptions,
  IUiEnvironmentSettings
} from '../models/Environment';

/**
 * ============================================================================
 *  SOLUTION ENVIRONMENT CONFIGURATION - edit this file to change settings
 * ============================================================================
 *
 * This is the single place where the .NET Core Web API address and other
 * cross-cutting settings are defined. Everything else reads them through
 * `getEnvironment()`:
 *
 *   getEnvironment().api.baseUrl   ->  ServiceFactory  ->  ApiService
 *   getEnvironment().api.endpoints ->  ProjectService / LoginService
 *   getEnvironment().auth          ->  ServiceFactory  ->  LoginService
 *   getEnvironment().ui            ->  components
 *
 * SPFx bundles run in the browser, so there is no `process.env` / `.env` file to
 * read at runtime. Configuration therefore lives in this typed module, and the
 * active environment is resolved from the host name (plus the `DEBUG` flag that
 * the SPFx webpack build defines). Add a tenant below to onboard a new
 * environment; no other file needs to change.
 */

/**
 * Web API controller routes. Shared by every environment because the API surface is the
 * same; only `baseUrl` differs.
 *
 * These are controller routes, not complete endpoints: a service appends the action it
 * needs (`Authenticate` + `SPFxLogin`, `Users` + `GetAllList`) and `ApiService.resolveUrl`
 * prefixes `baseUrl`. Keeping the controller here means a renamed route is a one-line
 * change in this file rather than a hunt through the services.
 */
const ENDPOINTS: IApiEndpoints = {
  projects: 'projects',
  authenticate: 'Authenticate',
  building: 'Building',
  users: 'Users',
  projectTemplates: 'ProjectTemplate',
  spatialBreakdowns: 'SpatialBreakdown',
  tenderTemplates: 'TenderTemplate',
  suppliers: 'Issuer',
  document: 'Document'
};

/**
 * Sign-in defaults.
 *
 * `domainUrl` is the address of the XSProject application the user is signed in to; the
 * Web API echoes it back into the session it creates. `ServiceFactory` resolves this
 * dynamically per request from `context.pageContext.web.absoluteUrl` (the current
 * SharePoint site), so the value below is only a fallback for the rare case that site URL
 * is unavailable.
 *
 * `tokenResourceUri` is the resource the SPFx access token is issued for. Microsoft Graph
 * is used because the Web API only needs to confirm that the caller is a signed-in
 * Microsoft 365 user. Point it at the Application ID URI of the app registration in front
 * of the API once one exists, and add a matching `webApiPermissionRequests` entry in
 * `config/package-solution.json`.
 */
const AUTH_DEFAULTS: IAuthEnvironmentSettings = {
  domainUrl: 'https://localhost:44316/',
  tokenResourceUri: 'https://graph.microsoft.com'
};

/** Presentation defaults. Shared by every environment unless overridden below. */
const UI_DEFAULTS: IUiEnvironmentSettings = {
  projectListTitle: 'Projects',
  maxProjectsToDisplay: 50
};

/** Requests are abandoned after this many milliseconds. Use `0` to disable. */
const DEFAULT_REQUEST_TIMEOUT_MS: number = 30000;

/**
 * Per-environment settings.
 *
 * `resourceUri` is the Application ID URI of the Entra ID app registration in front of
 * the API. Leave it as an empty string while the API is called anonymously; remember
 * that a non-empty value also needs a matching `webApiPermissionRequests` entry in
 * `config/package-solution.json`.
 */
const ENVIRONMENTS: { [name in EnvironmentName]: IEnvironmentConfiguration } = {
  local: {
    name: 'local',
    api: {
      // Matches BMDesk.API's **IIS Express** profile (`sslPort: 44396` in its
      // launchSettings.json), which is also the address its own appsettings.json records.
      // Its other profile - the one `dotnet run` and the "project" launch option use -
      // listens on https://localhost:5001 instead, and nothing here would reach it. If the
      // API is started that way, either switch to IIS Express or put
      // `https://localhost:5001/api` in the web part's "API base URL" property, which
      // overrides this without a rebuild.
      baseUrl: 'https://localhost:44396/api',
      resourceUri: undefined,
      requestTimeoutMs: DEFAULT_REQUEST_TIMEOUT_MS,
      defaultHeaders: { 'X-Api-Version': '1.0' },
      endpoints: ENDPOINTS
    },
    auth: AUTH_DEFAULTS,
    ui: UI_DEFAULTS,
    enableVerboseLogging: true
  },

  development: {
    name: 'development',
    api: {
      baseUrl: 'https://xsproject-api-dev.azurewebsites.net/api',
      resourceUri: undefined,
      requestTimeoutMs: DEFAULT_REQUEST_TIMEOUT_MS,
      defaultHeaders: { 'X-Api-Version': '1.0' },
      endpoints: ENDPOINTS
    },
    auth: AUTH_DEFAULTS,
    ui: UI_DEFAULTS,
    enableVerboseLogging: true
  },

  test: {
    name: 'test',
    api: {
      baseUrl: 'https://xsproject-api-test.azurewebsites.net/api',
      resourceUri: undefined,
      requestTimeoutMs: DEFAULT_REQUEST_TIMEOUT_MS,
      defaultHeaders: { 'X-Api-Version': '1.0' },
      endpoints: ENDPOINTS
    },
    auth: AUTH_DEFAULTS,
    ui: UI_DEFAULTS,
    enableVerboseLogging: false
  },

  production: {
    name: 'production',
    api: {
      baseUrl: 'https://xsproject-api.azurewebsites.net/api',
      resourceUri: undefined,
      requestTimeoutMs: DEFAULT_REQUEST_TIMEOUT_MS,
      defaultHeaders: { 'X-Api-Version': '1.0' },
      endpoints: ENDPOINTS
    },
    auth: AUTH_DEFAULTS,
    ui: UI_DEFAULTS,
    enableVerboseLogging: false
  }
};

/**
 * Maps a SharePoint host name to an environment, so one package can be deployed to
 * several tenants. Keys are compared case-insensitively. Unlisted hosts fall back to
 * {@link FALLBACK_ENVIRONMENT}.
 */
const HOST_NAME_TO_ENVIRONMENT: { [hostName: string]: EnvironmentName } = {
  'ifourtechnolab.sharepoint.com': 'production'
};

/** Environment used when the host name is not listed above. */
const FALLBACK_ENVIRONMENT: EnvironmentName = 'production';

let currentEnvironment: IEnvironmentConfiguration | undefined;

/** True while running under `heft start` / `heft build` (the SPFx build defines `DEBUG`). */
function isDebugBuild(): boolean {
  return typeof DEBUG !== 'undefined' && DEBUG;
}

/** Extracts the host name from an absolute URL, or from `window.location` as a fallback. */
function resolveHostName(hostName?: string): string {
  if (hostName) {
    const match: RegExpMatchArray | null = /^https?:\/\/([^/:]+)/i.exec(hostName);
    return (match ? match[1] : hostName).toLowerCase();
  }

  return typeof window !== 'undefined' && window.location ? window.location.hostname.toLowerCase() : '';
}

/**
 * Decides which environment applies, in priority order:
 * an explicit name, localhost / a debug build, the host name map, then the fallback.
 */
export function resolveEnvironmentName(options?: IEnvironmentResolutionOptions): EnvironmentName {
  if (options?.environmentName) {
    return options.environmentName;
  }

  const hostName: string = resolveHostName(options?.hostName);

  if (options?.isServedFromLocalhost || hostName === 'localhost' || (!hostName && isDebugBuild())) {
    return 'local';
  }

  return HOST_NAME_TO_ENVIRONMENT[hostName] || FALLBACK_ENVIRONMENT;
}

/**
 * Resolves and caches the active configuration. Call this once from a web part's
 * `onInit`, passing the SPFx context hints so detection is exact:
 *
 * ```ts
 * initializeEnvironment({
 *   hostName: this.context.pageContext.web.absoluteUrl,
 *   isServedFromLocalhost: this.context.isServedFromLocalhost
 * });
 * ```
 */
export function initializeEnvironment(options?: IEnvironmentResolutionOptions): IEnvironmentConfiguration {
  currentEnvironment = ENVIRONMENTS[resolveEnvironmentName(options)];
  return currentEnvironment;
}

/**
 * The active configuration, for use anywhere in the solution.
 *
 * Falls back to auto-detection when {@link initializeEnvironment} has not run, so it is
 * always safe to call from models-adjacent code, services and components alike.
 */
export function getEnvironment(): IEnvironmentConfiguration {
  if (!currentEnvironment) {
    currentEnvironment = ENVIRONMENTS[resolveEnvironmentName()];
  }

  return currentEnvironment;
}

/** Clears the cached configuration. Intended for unit tests. */
export function resetEnvironment(): void {
  currentEnvironment = undefined;
}
