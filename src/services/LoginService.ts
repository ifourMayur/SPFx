import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import { ApiError } from '../models/ApiError';
import { ILoginResponse, ISPFxLoginRequest } from '../models/Auth';
import { IAccessTokenProvider } from './AccessTokenProvider';
import { IApiTokenStore, apiTokenStore } from './ApiTokenStore';
import { IApiService } from './ApiService';

/** Source name used for SPFx log entries emitted by the sign-in flow. */
const LOG_SOURCE: string = 'LoginService';

/** Action on the `Authenticate` controller that exchanges an SPFx token for a session. */
const SPFX_LOGIN_ACTION: string = 'SPFxLogin';

/**
 * Sign-in operations exposed by the `AuthenticateController` of the .NET Core Web API.
 *
 * Components depend on this interface only, which keeps HTTP concerns and SPFx token
 * acquisition out of the UI and makes the components easy to test with a stub.
 */
export interface ILoginService {
  /**
   * Signs in with a token the caller already holds.
   *
   * `POST {baseUrl}/Authenticate/SPFxLogin` with `{ token, domainUrl }`.
   *
   * @throws ApiError - for any non-successful response, including 401 and 403, and for a
   *   success status that carries `errorMessage` or no application token.
   */
  login(request: ISPFxLoginRequest): Promise<ILoginResponse>;

  /**
   * Acquires the access token of the signed-in Microsoft 365 user and signs in with it.
   *
   * This is the entry point components use: the token is obtained, passed to
   * {@link ILoginService.login} and discarded, so it never reaches component state.
   *
   * @param domainUrl - overrides `auth.domainUrl` from `src/config/environment.ts`.
   * @throws TokenAcquisitionError - when SPFx cannot issue a token for the current user.
   * @throws ApiError - as for {@link ILoginService.login}.
   */
  loginAsCurrentUser(domainUrl?: string): Promise<ILoginResponse>;
}

/**
 * Maps `ILoginService` operations onto the sign-in endpoints of the Web API.
 *
 * URL building, serialization and transport error translation stay in `ApiService`; token
 * acquisition stays in the injected `IAccessTokenProvider`. This class orchestrates the
 * two, applies the configured defaults and decides whether a 200 actually signed the user
 * in, so components see one uniform failure path.
 */
export class LoginService implements ILoginService {
  private readonly _apiService: IApiService;
  private readonly _tokenProvider: IAccessTokenProvider;
  private readonly _endpoint: string;
  private readonly _domainUrl: string;
  private readonly _tokenStore: IApiTokenStore;

  /**
   * @param apiService - HTTP gateway used for the call.
   * @param tokenProvider - source of the signed-in user's access token.
   * @param endpoint - controller route; defaults to `api.endpoints.authenticate` from the
   *   environment configuration, and can be overridden in tests.
   * @param domainUrl - default `domainUrl` sent in the request body; defaults to
   *   `auth.domainUrl` from the environment configuration.
   * @param tokenStore - where the application JWT is published on success; defaults to the
   *   process-wide store `ApiService` reads when building request headers.
   */
  public constructor(
    apiService: IApiService,
    tokenProvider: IAccessTokenProvider,
    endpoint: string = getEnvironment().api.endpoints.authenticate,
    domainUrl: string = getEnvironment().auth.domainUrl,
    tokenStore: IApiTokenStore = apiTokenStore
  ) {
    this._apiService = apiService;
    this._tokenProvider = tokenProvider;
    this._endpoint = endpoint;
    this._domainUrl = domainUrl;
    this._tokenStore = tokenStore;
  }

  /** Default `domainUrl` sent to the API, from `src/config/environment.ts`. */
  public get domainUrl(): string {
    return this._domainUrl;
  }

  public async login(request: ISPFxLoginRequest): Promise<ILoginResponse> {
    const relativeUrl: string = `${this._endpoint}/${SPFX_LOGIN_ACTION}`;

    // Typed as possibly undefined because an API that answers 200 with an empty body is a
    // runtime possibility the compiler cannot see.
    const response: ILoginResponse | undefined = await this._apiService.post<
      ILoginResponse | undefined,
      ISPFxLoginRequest
    >(relativeUrl, request);

    this._logResponse(request.domainUrl, response);

    const url: string = this._apiService.resolveUrl(relativeUrl);

    if (!response) {
      throw new ApiError('The Web API returned an empty response to SPFxLogin.', { status: 0, url });
    }

    // The API reports business failures - a locked account, for example - as a success
    // status with `errorMessage` set, so the status alone does not mean the user is in.
    if (response.errorMessage) {
      throw new ApiError(response.errorMessage, { status: 200, statusText: 'OK', url });
    }

    if (!response.token) {
      throw new ApiError('The Web API signed in without issuing an application token.', {
        status: 200,
        statusText: 'OK',
        url
      });
    }

    // Published once, here, so every later call through `ApiService` carries
    // `Authorization: Bearer <token>` without any component having to pass it along. Done
    // only after the checks above, so a failed sign-in never leaves a stale token behind.
    this._tokenStore.setToken(response.token);

    return response;
  }

  public async loginAsCurrentUser(domainUrl: string = this._domainUrl): Promise<ILoginResponse> {
    const token: string = await this._tokenProvider.getAccessToken();

    return this.login({ token, domainUrl });
  }

  /**
   * Writes the complete response to the browser console, which is how the handshake is
   * verified during development.
   *
   * The request body is still never logged: it carries the user's Microsoft 365 access
   * token. The response is printed as received, including the application token and
   * refresh token the API issues, so remove this dump before shipping to production.
   */
  private _logResponse(domainUrl: string, response: ILoginResponse | undefined): void {
    Log.info(LOG_SOURCE, `${SPFX_LOGIN_ACTION} completed for domainUrl '${domainUrl}'.`);
    console.log(`[${LOG_SOURCE}] ${SPFX_LOGIN_ACTION} response:`, response);
  }
}
