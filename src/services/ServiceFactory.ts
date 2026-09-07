import { BaseComponentContext } from '@microsoft/sp-component-base';

import { getEnvironment } from '../config/environment';
import { IApiServiceConfiguration } from '../models/ApiResponse';
import { AadAccessTokenProvider, IAccessTokenProvider } from './AccessTokenProvider';
import { AadApiHttpClient, AnonymousApiHttpClient, IApiHttpClient } from './ApiHttpClient';
import { ApiService, IApiService, IResolvedApiConfiguration, resolveApiConfiguration } from './ApiService';
import { ILoginService, LoginService } from './LoginService';
import { ILookupService, LookupService } from './LookupService';
import { IProjectService, ProjectService } from './ProjectService';

/** All services available to the components, resolved as interfaces. */
export interface IServiceContainer {
  readonly apiService: IApiService;
  readonly loginService: ILoginService;
  readonly projectService: IProjectService;
  /** Reference data the project form binds its dropdowns to. */
  readonly lookupService: ILookupService;
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

    return {
      apiService,
      loginService: new LoginService(
        apiService,
        tokenProvider,
        getEnvironment().api.endpoints.authenticate,
        ServiceFactory._resolveDomainUrl(context)
      ),
      projectService: new ProjectService(apiService, getEnvironment().api.endpoints.projects),
      lookupService: new LookupService(apiService)
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

  private static _buildCacheKey(
    context: BaseComponentContext,
    configuration: IResolvedApiConfiguration
  ): string {
    return [context.instanceId, configuration.baseUrl, configuration.resourceUri || ''].join('|');
  }
}
