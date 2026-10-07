import { Log } from '@microsoft/sp-core-library';

import { unwrapResponseDetail } from '../models/ApiEnvelope';
import { ApiError } from '../models/ApiError';
import { IQueryParameters } from '../models/ApiResponse';
import { IApiService } from './ApiService';

/**
 * How {@link BaseApiService.getData} and {@link BaseApiService.postData} treat one call.
 */
export interface IEnvelopeCallOptions {
  /**
   * What is being fetched or written, in words - "the project templates". Used in the
   * error when the API refuses inside an HTTP 200 without a message of its own.
   */
  operation: string;

  /**
   * Prints the exchange to the browser console, **before** the envelope is unwrapped.
   *
   * That ordering is the point: `unwrapResponseDetail` throws on `success: false`, so a
   * dump placed after it would stay silent on exactly the responses worth reading. Omit
   * it and nothing is printed. This is the dev-only dump the save path relies on during
   * development; remove the callers' `print` options before shipping to production.
   */
  print?: {
    /** Tells apart two calls that share a URL, e.g. the insert and the rebind. */
    label?: string;
    /** Also prints the request body. Never set this on a body that carries a token. */
    includeRequest?: boolean;
  };
}

/**
 * Common base of every service that talks to the BMDesk Web API.
 *
 * The layering is:
 *
 * ```text
 * Component  ->  domain service (extends BaseApiService)  ->  IApiService  ->  IApiHttpClient
 * ```
 *
 * `IApiService` owns the HTTP: URL building, headers, the bearer token, the timeout and
 * turning a failed **status** into an `ApiError`. What it cannot know is that this API
 * answers HTTP 200 even when it refused the call - the `ResponseDetail` envelope - and
 * that is what this class adds: one place that calls `controller/action`, unwraps the
 * envelope, raises "refused inside a 200" as an `ApiError`, and logs consistently.
 *
 * It holds no business operations. Endpoints, parameters and response mapping stay in
 * the domain service that extends it, so `BaseApiService.getProject()` can never exist.
 *
 * Composition is kept underneath: a subclass is still constructed from an `IApiService`,
 * so tests keep running the real `ApiService` over a fake transport.
 */
export abstract class BaseApiService {
  /** HTTP gateway every call goes through. Exposed for the calls that are not enveloped. */
  protected readonly apiService: IApiService;

  /** Source name for SPFx log entries and the console dump - the subclass's own name. */
  protected readonly logSource: string;

  protected constructor(apiService: IApiService, logSource: string) {
    this.apiService = apiService;
    this.logSource = logSource;
  }

  /**
   * `controller/action`, the relative URL of one Web API action.
   *
   * The controller route comes from `src/config/environment.ts` and the action is the
   * subclass's own constant; `ApiService.resolveUrl` then prefixes `baseUrl`. Without an
   * action this is the controller route itself (`api/Building` routes on the verb).
   */
  protected route(controller: string, action?: string): string {
    return action ? `${controller}/${action}` : controller;
  }

  /** The complete URL of a relative one, for log lines and error details. */
  protected resolveUrl(relativeUrl: string, query?: IQueryParameters): string {
    return this.apiService.resolveUrl(relativeUrl, query);
  }

  /**
   * GET an enveloped action and return its `data`.
   *
   * @returns the payload, or `undefined` when a successful envelope carries no `data`.
   * @throws ApiError - for a failed status (from `ApiService`) and for `success: false`.
   */
  protected async getData<T>(
    relativeUrl: string,
    query: IQueryParameters | undefined,
    options: IEnvelopeCallOptions
  ): Promise<T | undefined> {
    const payload: unknown = await this.apiService.get<unknown>(relativeUrl, query);

    if (options.print) {
      this._printExchange('GET', this.resolveUrl(relativeUrl, query), options.print.label, undefined, payload);
    }

    return unwrapResponseDetail<T>(payload, options.operation);
  }

  /**
   * POST to an enveloped action and return its `data`.
   *
   * @returns the payload, or `undefined` when a successful envelope carries no `data`.
   * @throws ApiError - for a failed status (from `ApiService`) and for `success: false`.
   */
  protected async postData<T, TBody>(
    relativeUrl: string,
    body: TBody,
    options: IEnvelopeCallOptions
  ): Promise<T | undefined> {
    return unwrapResponseDetail<T>(await this.postRaw<TBody>(relativeUrl, body, options), options.operation);
  }

  /**
   * POST and hand back the raw body, envelope included, after the same optional dump.
   *
   * For the one endpoint whose envelope lies: `Document/AddDocumentStorageDetails` answers
   * `success: true` with `messageType: Error`, so its caller has to read the envelope
   * itself after unwrapping it.
   */
  protected async postRaw<TBody>(relativeUrl: string, body: TBody, options: IEnvelopeCallOptions): Promise<unknown> {
    const payload: unknown = await this.apiService.post<unknown, TBody>(relativeUrl, body);

    if (options.print) {
      this._printExchange(
        'POST',
        this.resolveUrl(relativeUrl),
        options.print.label,
        options.print.includeRequest ? body : undefined,
        payload
      );
    }

    return payload;
  }

  /**
   * The error for a call the API refused while answering HTTP 200.
   *
   * Status `200 OK` is carried because that genuinely is what the transport saw, which is
   * what tells this apart from a failed status further up.
   */
  protected refusal(message: string, relativeUrl: string, query?: IQueryParameters): ApiError {
    return new ApiError(message, { status: 200, statusText: 'OK', url: this.resolveUrl(relativeUrl, query) });
  }

  /** Logs a caught failure, describing it with `fallback` when it is not an `Error`. */
  protected logFailure(error: unknown, fallback: string): void {
    Log.error(this.logSource, error instanceof Error ? error : new Error(fallback));
  }

  /**
   * Writes one exchange to the browser console, exactly as the API sent it - the
   * `ResponseDetail` envelope included, since that is where a refusal is reported.
   */
  private _printExchange(
    method: string,
    url: string,
    label: string | undefined,
    request: unknown,
    payload: unknown
  ): void {
    const call: string = `${method} ${url}${label ? ` (${label})` : ''}`;

    Log.info(this.logSource, `${call} answered.`);

    if (request !== undefined) {
      console.log(`[${this.logSource}] ${call} request:`, request);
    }

    console.log(`[${this.logSource}] ${call} response:`, payload);
  }
}
