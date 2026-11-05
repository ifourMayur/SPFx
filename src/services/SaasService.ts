import { Log } from '@microsoft/sp-core-library';

import { ApiError, isApiError } from '../models/ApiError';
import {
  ISaasActivationResult,
  ISaasCatalog,
  ISaasEntitlements,
  ISaasPurchaseIntent,
  ISaasPurchaseIntentRequest,
  ISaasSubscription,
  toSaasActivationResult,
  toSaasCatalog,
  toSaasEntitlements,
  toSaasSubscriptions
} from '../models/Saas';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/** Actions of the `api/saas` controller. */
const ENTITLEMENTS_ACTION: string = 'entitlements';
const CATALOG_ACTION: string = 'catalog';
const SUBSCRIPTIONS_ACTION: string = 'subscriptions';
const PURCHASE_INTENTS_ACTION: string = 'purchase-intents';
/** Sub-action of one subscription: `subscriptions/{marketplaceSubscriptionId}/activate`. */
const ACTIVATE_ACTION: string = 'activate';

/** Prefix of the sessionStorage key the entitlements are cached under. */
const CACHE_KEY_PREFIX: string = 'xsProject.saas.entitlements';

/** How {@link ISaasService.getEntitlements} treats the cache. */
export interface IGetEntitlementsOptions {
  /**
   * Skip the cache and ask the server - after a purchase, on the Subscription page, or on
   * an explicit Retry. The fresh answer replaces the cached one.
   */
  refresh?: boolean;
}

/**
 * Microsoft Marketplace SaaS subscriptions: what the caller's tenant may use, what it can
 * buy, starting a purchase and activating what was bought.
 *
 * The answers are the server's (`api/saas`, evaluated from the caller's verified token).
 * Components use them to decide what to show; the Web API enforces the same answer on
 * every gated endpoint, so nothing here is a security boundary.
 */
export interface ISaasService {
  /**
   * `GET saas/entitlements`, cached in sessionStorage for the response's `cacheSeconds`.
   * Concurrent calls share one request, and a failure is never cached.
   */
  getEntitlements(options?: IGetEntitlementsOptions): Promise<ISaasEntitlements>;

  /** `GET saas/catalog` - every service, its features and plans. Not cached. */
  getCatalog(): Promise<ISaasCatalog>;

  /** `GET saas/subscriptions` - the tenant's Marketplace subscriptions. Not cached. */
  getSubscriptions(): Promise<ISaasSubscription[]>;

  /**
   * `POST saas/purchase-intents`: records who is buying what and where to come back to,
   * and answers the Marketplace offer URL to send the buyer to. Admins only (server-side).
   */
  createPurchaseIntent(serviceKey: string, returnUrl: string): Promise<ISaasPurchaseIntent>;

  /**
   * `POST saas/subscriptions/{marketplaceSubscriptionId}/activate`: activates a purchased
   * Marketplace subscription for the caller's tenant. Tenant admins only (server-side);
   * repeating it for an active subscription does nothing twice (`AlreadyActive`).
   *
   * The answer carries the tenant's fresh entitlements, which replace the cached ones at
   * once, so every gate sees the service unlocked without another round trip. The caller's
   * token is unchanged: the Web API checks entitlements per request.
   */
  activateSubscription(marketplaceSubscriptionId: string): Promise<ISaasActivationResult>;

  /** Forgets the cached entitlements, so the next read asks the server. */
  invalidate(): void;
}

/** The part of the Web Storage API the cache uses. */
export interface ISaasCacheStorage {
  getItem(key: string): string | null; // eslint-disable-line @rushstack/no-new-null -- Web Storage API signature
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export interface ISaasServiceOptions {
  /**
   * Who the cache belongs to - the tenant id and user, from the SPFx page context. Part
   * of the storage key together with the API base URL, so a cached answer can never be
   * read back for another tenant, user or API.
   */
  cacheScope?: string;
  /**
   * Where entitlements are cached across page loads. Defaults to `window.sessionStorage`;
   * `null` keeps the cache in memory only. Storage that throws (blocked cookies, private
   * mode quota) degrades to the in-memory cache rather than failing the call.
   */
  storage?: ISaasCacheStorage | null; // eslint-disable-line @rushstack/no-new-null -- explicit "no storage"
  /** Clock, in milliseconds; injectable for tests. */
  now?: () => number;
}

/** What is cached, in memory and in sessionStorage. */
interface ICachedEntitlements {
  key: string;
  expiresAt: number;
  value: ISaasEntitlements;
}

/** The request currently in flight, shared by every caller that asks meanwhile. */
interface IInFlightRequest {
  key: string;
  isRefresh: boolean;
  promise: Promise<ISaasEntitlements>;
}

/**
 * Default `ISaasService`, over the `api/saas` controller.
 *
 * These endpoints are **not** enveloped: success is a 2xx with the DTO as the body, and a
 * failure is a non-2xx with `{ code, message, correlationId }`, which `ApiService` turns
 * into an `ApiError` carrying `code` and `correlationId`. So this service calls
 * `apiService` directly instead of `getData`/`postData` - unwrapping would be harmless
 * today, but a DTO that ever gained a `data` or `success` property would be misread as an
 * envelope.
 */
export class SaasService extends BaseApiService implements ISaasService {
  private readonly _controller: string;
  private readonly _cacheScope: string;
  private readonly _now: () => number;
  private readonly _storageOverride: ISaasCacheStorage | null | undefined;
  private _memory: ICachedEntitlements | undefined;
  private _inFlight: IInFlightRequest | undefined;
  /**
   * Bumped by every request, by an activation and by `invalidate`; only the newest request
   * may write the cache, so a slow answer from before a refresh or an activation can't
   * overwrite the fresher one.
   */
  private _generation: number = 0;

  public constructor(apiService: IApiService, controller: string, options?: ISaasServiceOptions) {
    super(apiService, 'SaasService');

    this._controller = controller;
    this._cacheScope = options?.cacheScope || '';
    this._now = options?.now || ((): number => Date.now());
    this._storageOverride = options?.storage;
  }

  public getEntitlements(options?: IGetEntitlementsOptions): Promise<ISaasEntitlements> {
    const key: string = this._cacheKey();
    const isRefresh: boolean = !!options?.refresh;
    const inFlight: IInFlightRequest | undefined =
      this._inFlight && this._inFlight.key === key ? this._inFlight : undefined;

    if (!isRefresh) {
      const cached: ISaasEntitlements | undefined = this._readCache(key);
      if (cached) {
        return Promise.resolve(cached);
      }
      if (inFlight) {
        return inFlight.promise;
      }
    } else if (inFlight && inFlight.isRefresh) {
      // Already asking the server afresh; a second refresh would learn nothing new.
      return inFlight.promise;
    }

    const generation: number = ++this._generation;
    const promise: Promise<ISaasEntitlements> = this._fetchEntitlements(key, generation);
    const request: IInFlightRequest = { key, isRefresh, promise };
    const settle = (): void => {
      if (this._inFlight === request) {
        this._inFlight = undefined;
      }
    };

    this._inFlight = request;
    promise.then(settle, settle);

    return promise;
  }

  public async getCatalog(): Promise<ISaasCatalog> {
    const relativeUrl: string = this.route(this._controller, CATALOG_ACTION);
    const payload: unknown = await this._call('load the service catalog', () =>
      this.apiService.get<unknown>(relativeUrl)
    );
    const catalog: ISaasCatalog | undefined = toSaasCatalog(payload);

    if (!catalog) {
      throw this._unexpected('the service catalog', relativeUrl);
    }

    return catalog;
  }

  public async getSubscriptions(): Promise<ISaasSubscription[]> {
    const relativeUrl: string = this.route(this._controller, SUBSCRIPTIONS_ACTION);
    const payload: unknown = await this._call('load the subscriptions', () =>
      this.apiService.get<unknown>(relativeUrl)
    );

    const subscriptions: ISaasSubscription[] | undefined = toSaasSubscriptions(payload);

    if (!subscriptions) {
      throw this._unexpected('the subscriptions', relativeUrl);
    }

    return subscriptions;
  }

  public async createPurchaseIntent(serviceKey: string, returnUrl: string): Promise<ISaasPurchaseIntent> {
    const relativeUrl: string = this.route(this._controller, PURCHASE_INTENTS_ACTION);
    const body: ISaasPurchaseIntentRequest = { serviceKey, returnUrl };
    const payload: unknown = await this._call(`start the purchase of ${serviceKey}`, () =>
      this.apiService.post<unknown, ISaasPurchaseIntentRequest>(relativeUrl, body)
    );
    const intent: ISaasPurchaseIntent | undefined =
      typeof payload === 'object' && payload !== null ? (payload as ISaasPurchaseIntent) : undefined;

    if (!intent || typeof intent.marketplaceOfferUrl !== 'string' || !intent.marketplaceOfferUrl) {
      throw this._unexpected('the purchase', relativeUrl);
    }

    return intent;
  }

  public async activateSubscription(marketplaceSubscriptionId: string): Promise<ISaasActivationResult> {
    const key: string = this._cacheKey();
    const relativeUrl: string = this.route(
      this._controller,
      `${SUBSCRIPTIONS_ACTION}/${encodeURIComponent(marketplaceSubscriptionId)}/${ACTIVATE_ACTION}`
    );
    const payload: unknown = await this._call(`activate subscription ${marketplaceSubscriptionId}`, () =>
      this.apiService.post<unknown, Record<string, never>>(relativeUrl, {})
    );
    const result: ISaasActivationResult | undefined = toSaasActivationResult(payload);

    if (!result) {
      throw this._unexpected('the activation', relativeUrl);
    }

    // The newest answer there is: an entitlements read still in flight was evaluated before
    // the activation, so it is superseded and may no longer write the cache.
    this._generation++;
    this._inFlight = undefined;
    this._writeCache(key, result.entitlements);

    return result;
  }

  public invalidate(): void {
    const key: string = this._cacheKey();

    this._generation++;
    this._memory = undefined;
    this._inFlight = undefined;
    this._withStorage((storage: ISaasCacheStorage) => storage.removeItem(key));
  }

  private async _fetchEntitlements(key: string, generation: number): Promise<ISaasEntitlements> {
    const relativeUrl: string = this.route(this._controller, ENTITLEMENTS_ACTION);
    const payload: unknown = await this._call('load the entitlements', () =>
      this.apiService.get<unknown>(relativeUrl)
    );
    const entitlements: ISaasEntitlements | undefined = toSaasEntitlements(payload);

    if (!entitlements) {
      throw this._unexpected('the entitlements', relativeUrl);
    }

    if (generation === this._generation) {
      this._writeCache(key, entitlements);
    }

    return entitlements;
  }

  /** Runs one call, logging a failure (with its code and reference) before rethrowing it. */
  private async _call(operation: string, send: () => Promise<unknown>): Promise<unknown> {
    try {
      return await send();
    } catch (error) {
      this._logSaasFailure(operation, error);
      throw error;
    }
  }

  /**
   * Logs a failure with what support needs - status, `code`, `correlationId` - and never
   * the request headers, so the bearer token cannot end up in the log. A 4xx is the
   * server answering as designed (not an admin, nothing to buy), so it is a warning.
   */
  private _logSaasFailure(operation: string, error: unknown): void {
    if (!isApiError(error)) {
      Log.error(this.logSource, new Error(`Could not ${operation}: ${error instanceof Error ? error.message : 'unknown error'}`));
      return;
    }

    const message: string =
      `Could not ${operation}: status ${error.status}, code ${error.code || '-'}, ` +
      `correlationId ${error.correlationId || '-'}. ${error.message}`;

    if (error.status >= 400 && error.status < 500) {
      Log.warn(this.logSource, message);
    } else {
      Log.error(this.logSource, new Error(message));
    }
  }

  /** The error for a 2xx whose body is not the promised DTO. */
  private _unexpected(what: string, relativeUrl: string): ApiError {
    const error: ApiError = new ApiError(`The Web API answered ${what} in an unexpected format.`, {
      status: 200,
      statusText: 'OK',
      url: this.resolveUrl(relativeUrl)
    });

    Log.error(this.logSource, error);
    return error;
  }

  private _cacheKey(): string {
    return `${CACHE_KEY_PREFIX}|${this.apiService.baseUrl}|${this._cacheScope}`;
  }

  private _readCache(key: string): ISaasEntitlements | undefined {
    const now: number = this._now();

    if (this._memory && this._memory.key === key && this._memory.expiresAt > now) {
      return this._memory.value;
    }

    let stored: ICachedEntitlements | undefined;
    this._withStorage((storage: ISaasCacheStorage): void => {
      const raw: string | null = storage.getItem(key);
      if (raw) {
        const parsed: ICachedEntitlements = JSON.parse(raw);
        const value: ISaasEntitlements | undefined = toSaasEntitlements(parsed?.value);
        if (value && typeof parsed.expiresAt === 'number' && parsed.expiresAt > now) {
          stored = { key, expiresAt: parsed.expiresAt, value };
        }
      }
    });

    this._memory = stored;
    return stored ? stored.value : undefined;
  }

  private _writeCache(key: string, value: ISaasEntitlements): void {
    const entry: ICachedEntitlements = { key, expiresAt: this._now() + value.cacheSeconds * 1000, value };

    this._memory = entry;
    this._withStorage((storage: ISaasCacheStorage) => storage.setItem(key, JSON.stringify(entry)));
  }

  /**
   * Runs `action` against the cache storage, if there is one. Every failure - storage
   * unavailable, quota exceeded, a corrupt entry - is swallowed: the cache is an
   * optimisation, and the in-memory copy still works without it.
   */
  private _withStorage(action: (storage: ISaasCacheStorage) => void): void {
    try {
      const storage: ISaasCacheStorage | null | undefined =
        this._storageOverride !== undefined
          ? this._storageOverride
          : typeof window !== 'undefined'
            ? window.sessionStorage
            : undefined;

      if (storage) {
        action(storage);
      }
    } catch {
      // Deliberately ignored - see above.
    }
  }
}
