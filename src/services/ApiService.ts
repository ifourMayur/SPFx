import { Log } from '@microsoft/sp-core-library';
import { HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import { getEnvironment } from '../config/environment';
import { ApiError } from '../models/ApiError';
import {
  HttpMethod,
  IApiRequest,
  IApiResponse,
  IApiServiceConfiguration,
  IProblemDetails,
  IQueryParameters,
  IRequestHeaders
} from '../models/ApiResponse';
import { IEnvironmentConfiguration } from '../models/Environment';
import { IApiHttpClient } from './ApiHttpClient';

/** Source name used for SPFx log entries emitted by the API layer. */
const LOG_SOURCE: string = 'ApiService';

/** Matches a relativeUrl that is in fact already absolute. */
const ABSOLUTE_URL_PATTERN: RegExp = /^https?:\/\//i;

/** Effective API settings after merging overrides onto `src/config/environment.ts`. */
export interface IResolvedApiConfiguration {
  baseUrl: string;
  resourceUri: string | undefined;
  defaultHeaders: IRequestHeaders | undefined;
  timeoutMs: number;
  enableVerboseLogging: boolean;
}

/**
 * Returns the trimmed override when it actually holds text, otherwise the environment
 * value. Whitespace-only overrides (an all-spaces property pane field) are treated as
 * unset rather than blanking out a configured setting.
 */
function coalesceText(override: string | undefined, fallback: string | undefined): string | undefined {
  const trimmed: string = (override || '').trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

/**
 * Merges optional per-instance overrides onto the environment configuration.
 *
 * The environment file is the source of truth; empty or omitted override values fall
 * back to it, so an unset property pane field never blanks out a configured setting.
 * Explicit `0` and `false` are honoured, since they are meaningful values.
 */
export function resolveApiConfiguration(overrides?: IApiServiceConfiguration): IResolvedApiConfiguration {
  const environment: IEnvironmentConfiguration = getEnvironment();

  return {
    baseUrl: coalesceText(overrides?.baseUrl, environment.api.baseUrl) || '',
    resourceUri: coalesceText(overrides?.resourceUri, environment.api.resourceUri),
    defaultHeaders: overrides?.defaultHeaders || environment.api.defaultHeaders,
    timeoutMs: overrides?.timeoutMs !== undefined ? overrides.timeoutMs : environment.api.requestTimeoutMs,
    enableVerboseLogging:
      overrides?.enableVerboseLogging !== undefined
        ? overrides.enableVerboseLogging
        : environment.enableVerboseLogging
  };
}

/**
 * Generic HTTP gateway to the .NET Core Web API.
 *
 * Resource specific services (for example `ProjectService`) depend on this
 * interface, so URL building, JSON serialization and error translation live in
 * exactly one place and never leak into React components.
 */
export interface IApiService {
  /** Base address every relative URL is resolved against, from `src/config/environment.ts`. */
  readonly baseUrl: string;

  /** Completes an endpoint: `baseUrl` + `relativeUrl` + query string. */
  resolveUrl(relativeUrl: string, query?: IQueryParameters): string;

  get<TResponse>(relativeUrl: string, query?: IQueryParameters, headers?: IRequestHeaders): Promise<TResponse>;

  post<TResponse, TBody = unknown>(relativeUrl: string, body: TBody, headers?: IRequestHeaders): Promise<TResponse>;

  put<TResponse, TBody = unknown>(relativeUrl: string, body: TBody, headers?: IRequestHeaders): Promise<TResponse>;

  patch<TResponse, TBody = unknown>(relativeUrl: string, body: TBody, headers?: IRequestHeaders): Promise<TResponse>;

  delete(relativeUrl: string, headers?: IRequestHeaders): Promise<void>;

  /** Escape hatch for calls that need the status code or a non-standard verb. */
  request<TResponse, TBody = undefined>(request: IApiRequest<TBody>): Promise<IApiResponse<TResponse>>;
}

/**
 * Default implementation of `IApiService` on top of an `IApiHttpClient`.
 *
 * Settings come straight from `src/config/environment.ts`, so the service can be
 * created with nothing but a transport:
 *
 * ```ts
 * const apiService = new ApiService(httpClient);
 * ```
 *
 * Pass `overrides` to replace individual settings for one instance.
 */
export class ApiService implements IApiService {
  private readonly _httpClient: IApiHttpClient;
  private readonly _overrides: IApiServiceConfiguration | undefined;
  private _resolved: IResolvedApiConfiguration | undefined;

  public constructor(httpClient: IApiHttpClient, overrides?: IApiServiceConfiguration) {
    this._httpClient = httpClient;
    this._overrides = overrides;
  }

  public get baseUrl(): string {
    return this._configuration.baseUrl;
  }

  /**
   * Completes an endpoint by merging the configured base URL with a relative URL and
   * the query string:
   *
   * ```text
   * baseUrl      https://localhost:5001/api      <- src/config/environment.ts
   * relativeUrl  projects/12
   * query        { search: "roof" }
   * ------------------------------------------------------------------------
   * endpoint     https://localhost:5001/api/projects/12?search=roof
   * ```
   *
   * Duplicate and missing slashes between the two halves are normalised. A
   * `relativeUrl` that is already absolute is used as-is, which lets a caller reach a
   * second host without going through the environment configuration. An existing query
   * string on `relativeUrl` is preserved and extended with `&`.
   */
  public resolveUrl(relativeUrl: string, query?: IQueryParameters): string {
    const path: string = (relativeUrl || '').trim();
    let endpoint: string;

    if (ABSOLUTE_URL_PATTERN.test(path)) {
      endpoint = path.replace(/\/+$/, '');
    } else {
      const base: string = this.baseUrl.trim();
      if (!base) {
        throw new ApiError(
          'The Web API base URL is not configured. Set api.baseUrl in src/config/environment.ts.',
          { status: 0 }
        );
      }

      const trimmedBase: string = base.replace(/\/+$/, '');
      const trimmedPath: string = path.replace(/^\/+/, '');
      endpoint = trimmedPath.length > 0 ? `${trimmedBase}/${trimmedPath}` : trimmedBase;
    }

    const queryString: string = this._buildQueryString(query);
    if (!queryString) {
      return endpoint;
    }

    return `${endpoint}${endpoint.indexOf('?') >= 0 ? '&' : '?'}${queryString}`;
  }

  /**
   * Effective settings, resolved on first use so that the initializeEnvironment() call
   * made by the hosting web part is always picked up regardless of construction order.
   */
  private get _configuration(): IResolvedApiConfiguration {
    if (!this._resolved) {
      this._resolved = resolveApiConfiguration(this._overrides);
    }

    return this._resolved;
  }

  public async get<TResponse>(
    relativeUrl: string,
    query?: IQueryParameters,
    headers?: IRequestHeaders
  ): Promise<TResponse> {
    const response: IApiResponse<TResponse> = await this.request<TResponse>({
      method: 'GET',
      relativeUrl,
      query,
      headers
    });

    return response.data;
  }

  public async post<TResponse, TBody = unknown>(
    relativeUrl: string,
    body: TBody,
    headers?: IRequestHeaders
  ): Promise<TResponse> {
    return this._send<TResponse, TBody>('POST', relativeUrl, body, headers);
  }

  public async put<TResponse, TBody = unknown>(
    relativeUrl: string,
    body: TBody,
    headers?: IRequestHeaders
  ): Promise<TResponse> {
    return this._send<TResponse, TBody>('PUT', relativeUrl, body, headers);
  }

  public async patch<TResponse, TBody = unknown>(
    relativeUrl: string,
    body: TBody,
    headers?: IRequestHeaders
  ): Promise<TResponse> {
    return this._send<TResponse, TBody>('PATCH', relativeUrl, body, headers);
  }

  public async delete(relativeUrl: string, headers?: IRequestHeaders): Promise<void> {
    await this.request<void>({ method: 'DELETE', relativeUrl, headers });
  }

  public async request<TResponse, TBody = undefined>(request: IApiRequest<TBody>): Promise<IApiResponse<TResponse>> {
    const url: string = this.resolveUrl(request.relativeUrl, request.query);
    const hasBody: boolean = request.body !== undefined;
    const timeout: ITimeout = this._startTimeout();
    const options: IHttpClientOptions = {
      method: request.method,
      headers: this._buildHeaders(request.headers, hasBody),
      body: hasBody ? JSON.stringify(request.body) : undefined,
      signal: timeout.signal
    };

    this._logVerbose(`${request.method} ${url}`);

    let response: HttpClientResponse;
    try {
      response = await this._httpClient.fetch(url, options);
    } catch (error) {
      // Network failure, CORS rejection, timeout or token acquisition failure:
      // the request never completed.
      const reason: string = timeout.hasExpired
        ? `The request timed out after ${this._timeoutMs}ms.`
        : error instanceof Error
          ? error.message
          : '';
      const apiError: ApiError = new ApiError(
        `Unable to reach the Web API at ${url}. ${reason}`.trim(),
        { status: 0, url }
      );
      this._logError(apiError);
      throw apiError;
    } finally {
      timeout.clear();
    }

    const payload: unknown = await this._readPayload(response);

    if (!response.ok) {
      const apiError: ApiError = this._toApiError(response, url, payload);
      this._logError(apiError);
      throw apiError;
    }

    this._logVerbose(`${request.method} ${url} -> ${response.status}`);

    return {
      ok: true,
      status: response.status,
      statusText: response.statusText,
      data: payload as TResponse
    };
  }

  private async _send<TResponse, TBody>(
    method: HttpMethod,
    relativeUrl: string,
    body: TBody,
    headers?: IRequestHeaders
  ): Promise<TResponse> {
    const response: IApiResponse<TResponse> = await this.request<TResponse, TBody>({
      method,
      relativeUrl,
      body,
      headers
    });

    return response.data;
  }

  private _buildQueryString(query?: IQueryParameters): string {
    if (!query) {
      return '';
    }

    const parts: string[] = [];
    Object.keys(query).forEach((name: string): void => {
      const value: string | number | boolean | undefined = query[name];
      if (value !== undefined) {
        parts.push(`${encodeURIComponent(name)}=${encodeURIComponent(String(value))}`);
      }
    });

    return parts.join('&');
  }

  private _buildHeaders(headers: IRequestHeaders | undefined, hasBody: boolean): IRequestHeaders {
    const result: IRequestHeaders = { Accept: 'application/json' };

    if (hasBody) {
      result['Content-Type'] = 'application/json';
    }

    const defaults: IRequestHeaders | undefined = this._configuration.defaultHeaders;
    if (defaults) {
      Object.keys(defaults).forEach((name: string): void => {
        result[name] = defaults[name];
      });
    }

    if (headers) {
      Object.keys(headers).forEach((name: string): void => {
        result[name] = headers[name];
      });
    }

    return result;
  }

  /** Reads the response body once, tolerating empty (204/205) and non-JSON payloads. */
  private async _readPayload(response: HttpClientResponse): Promise<unknown> {
    let text: string;
    try {
      text = await response.text();
    } catch {
      return undefined;
    }

    if (!text) {
      return undefined;
    }

    try {
      return JSON.parse(text);
    } catch {
      return text;
    }
  }

  private _toApiError(response: HttpClientResponse, url: string, payload: unknown): ApiError {
    const problemDetails: IProblemDetails | undefined = this._toProblemDetails(payload);
    const message: string =
      problemDetails?.detail ||
      problemDetails?.title ||
      problemDetails?.message ||
      (typeof payload === 'string' && payload.length > 0 ? payload : '') ||
      `The Web API returned ${response.status} ${response.statusText}.`;

    return new ApiError(message, {
      status: response.status,
      statusText: response.statusText,
      url,
      problemDetails
    });
  }

  private _toProblemDetails(payload: unknown): IProblemDetails | undefined {
    return typeof payload === 'object' && payload !== null ? (payload as IProblemDetails) : undefined;
  }

  private get _timeoutMs(): number {
    return this._configuration.timeoutMs;
  }

  /**
   * Starts the configured request timeout.
   *
   * Returns a no-op handle when the timeout is disabled or when `AbortController` is
   * unavailable, so the request still runs unbounded rather than failing outright.
   */
  private _startTimeout(): ITimeout {
    const timeoutMs: number = this._timeoutMs;

    if (timeoutMs <= 0 || typeof AbortController === 'undefined') {
      return { signal: undefined, hasExpired: false, clear: (): void => undefined };
    }

    const controller: AbortController = new AbortController();
    let handle: number = 0;
    const timeout: ITimeout = {
      signal: controller.signal,
      hasExpired: false,
      clear: (): void => clearTimeout(handle)
    };

    handle = setTimeout((): void => {
      timeout.hasExpired = true;
      controller.abort();
    }, timeoutMs) as unknown as number;

    return timeout;
  }

  private _logVerbose(message: string): void {
    if (this._configuration.enableVerboseLogging) {
      Log.verbose(LOG_SOURCE, message);
    }
  }

  private _logError(error: ApiError): void {
    if (this._configuration.enableVerboseLogging) {
      Log.error(LOG_SOURCE, error);
    }
  }
}

/** Handle over an in-flight request timeout. */
interface ITimeout {
  signal: AbortSignal | undefined;
  hasExpired: boolean;
  clear: () => void;
}
