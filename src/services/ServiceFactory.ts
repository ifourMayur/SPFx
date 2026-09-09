import { BaseComponentContext } from '@microsoft/sp-component-base';

import { getEnvironment } from '../config/environment';
import { IApiServiceConfiguration } from '../models/ApiResponse';
import { AadAccessTokenProvider, IAccessTokenProvider } from './AccessTokenProvider';
import { AadApiHttpClient, AnonymousApiHttpClient, IApiHttpClient } from './ApiHttpClient';
import { ApiService, IApiService, IResolvedApiConfiguration, resolveApiConfiguration } from './ApiService';
import { BuildingService, IBuildingService } from './BuildingService';
import { ILoginService, LoginService } from './LoginService';
import { ILookupService, LookupService } from './LookupService';
import { IProjectService, ProjectService } from './ProjectService';
import { DocumentStorageService, IDocumentStorageService } from './DocumentStorageService';
import { IProjectTemplateService, ProjectTemplateService } from './ProjectTemplateService';
import { ISharePointFolderService, SharePointFolderService } from './SharePointFolderService';
import { ISpHttpClient, SpRestHttpClient } from './SharePointHttpClient';
import { ISharePointSiteService, SharePointSiteService } from './SharePointSiteService';

/** All services available to the components, resolved as interfaces. */
export interface IServiceContainer {
  readonly apiService: IApiService;
  readonly loginService: ILoginService;
  readonly projectService: IProjectService;
  /** Write path of the project add/edit form: `POST api/Building`. */
  readonly buildingService: IBuildingService;
  /** What a project template defines - today its folder tree. */
  readonly projectTemplateService: IProjectTemplateService;
  /**
   * Records which SharePoint folder each of a project's folders became, once they have
   * been created: `Document/AddDocumentStorageDetails`.
   */
  readonly documentStorageService: IDocumentStorageService;
  /** Reference data the project form binds its dropdowns to. */
  readonly lookupService: ILookupService;
  /**
   * The SharePoint sites the user may work in, read from SharePoint itself rather than
   * from the XSProject Web API.
   */
  readonly sharePointSiteService: ISharePointSiteService;
}

/**
 * Single place where concrete service classes are wired to the SPFx runtime.
 *
 * `ApiService` reads the API address and related settings from the environment
 * configuration itself, so callers normally just do:
 *
 * ```ts
 * const services = ServiceFactory.getServices(this.context);
 * ```
 *
 * A caller may pass `overrides` to replace individual settings for one component
 * instance (for example a base URL typed into the property pane).
 *
 * The container is cached per resolved configuration, so re-rendering a web part reuses
 * the same service instances - and therefore the same cached Entra ID token - instead of
 * rebuilding the API layer.
 */
export class ServiceFactory {
  private static _container: IServiceContainer | undefined;
  private static _cacheKey: string | undefined;

  public static getServices(
    context: BaseComponentContext,
    overrides?: IApiServiceConfiguration
  ): IServiceContainer {
    // Resolved here only to pick the transport and build a stable cache key; the service
    // itself re-reads the environment configuration through `resolveApiConfiguration`.
    const configuration: IResolvedApiConfiguration = resolveApiConfiguration(overrides);
    const cacheKey: string = ServiceFactory._buildCacheKey(context, configuration);

    if (!ServiceFactory._container || ServiceFactory._cacheKey !== cacheKey) {
      ServiceFactory._container = ServiceFactory._createContainer(context, configuration, overrides);
      ServiceFactory._cacheKey = cacheKey;
    }

    return ServiceFactory._container;
  }

  /** Drops the cached container. Useful in unit tests and when disposing a component. */
  public static reset(): void {
    ServiceFactory._container = undefined;
    ServiceFactory._cacheKey = undefined;
  }

  private static _createContainer(
    context: BaseComponentContext,
    configuration: IResolvedApiConfiguration,
    overrides?: IApiServiceConfiguration
  ): IServiceContainer {
    const httpClient: IApiHttpClient = configuration.resourceUri
      ? new AadApiHttpClient(context.aadHttpClientFactory, configuration.resourceUri)
      : new AnonymousApiHttpClient(context.httpClient);

    const apiService: IApiService = new ApiService(httpClient, overrides);

    // The sign-in handshake needs the raw user token in the request body rather than in an
    // `Authorization` header, so it uses the token provider directly instead of the
    // Entra ID aware transport above.
    const tokenProvider: IAccessTokenProvider = new AadAccessTokenProvider(
      context.aadTokenProviderFactory,
      getEnvironment().auth.tokenResourceUri
    );

    const projectTemplateService: IProjectTemplateService = new ProjectTemplateService(
      apiService,
      getEnvironment().api.endpoints.projectTemplates
    );

    // One SharePoint transport for both SharePoint services: it holds no per-call state, and
    // sharing it keeps "which client talks to SharePoint" a single answer.
    const spHttpClient: ISpHttpClient = new SpRestHttpClient(context.spHttpClient);
    const folderService: ISharePointFolderService = new SharePointFolderService(spHttpClient);
    const documentStorageService: IDocumentStorageService = new DocumentStorageService(
      apiService,
      getEnvironment().api.endpoints.document
    );

    return {
      apiService,
      projectTemplateService,
      documentStorageService,
      loginService: new LoginService(
        apiService,
        tokenProvider,
        getEnvironment().api.endpoints.authenticate,
        ServiceFactory._resolveDomainUrl(context)
      ),
      projectService: new ProjectService(apiService, getEnvironment().api.endpoints.projects),
      // Given the template service rather than reaching for it, so the save's own
      // sequencing stays testable without a second transport.
      buildingService: new BuildingService(
        apiService,
        projectTemplateService,
        folderService,
        documentStorageService,
        getEnvironment().api.endpoints.building
      ),
      lookupService: new LookupService(apiService),
      // Talks to SharePoint, not to the Web API, so it takes the SharePoint transport and
      // the current site - whose `_api` endpoint the tenant-wide search query is addressed
      // to.
      sharePointSiteService: new SharePointSiteService(
        spHttpClient,
        ServiceFactory._resolveWebUrl(context)
      )
    };
  }

  /**
   * The XSProject application address sent as `domainUrl` in the sign-in request.
   *
   * Resolved from the current SharePoint site (`context.pageContext.web.absoluteUrl`) so
   * the Web API sees which site/tenant the session belongs to, rather than the static
   * placeholder in `auth.domainUrl` - which is kept only as a fallback for the rare case
   * the site URL is not available.
   */
  private static _resolveDomainUrl(context: BaseComponentContext): string {
    return context.pageContext?.web?.absoluteUrl || getEnvironment().auth.domainUrl;
  }

  /**
   * The SharePoint site whose `_api` endpoint SharePoint requests are addressed to.
   *
   * Kept apart from {@link _resolveDomainUrl}, whose fallback is the XSProject application
   * address - not a SharePoint site, and so not somewhere `_api/search/query` exists. A
   * search query is tenant-wide whichever site receives it, so the page's own origin is a
   * safe fallback if the site URL is ever unavailable.
   */
  private static _resolveWebUrl(context: BaseComponentContext): string {
    return context.pageContext?.web?.absoluteUrl || window.location.origin;
  }

  private static _buildCacheKey(
    context: BaseComponentContext,
    configuration: IResolvedApiConfiguration
  ): string {
    return [context.instanceId, configuration.baseUrl, configuration.resourceUri || ''].join('|');
  }
}
