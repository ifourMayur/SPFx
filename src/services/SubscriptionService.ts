import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import {
  FeatureCode,
  IFeatureSubscription,
  ISubscriptionLinks,
  toActivatedFeatureSubscription,
  toFeatureSubscription,
  toFeatureSubscriptions
} from '../models/FeatureSubscription';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/** How {@link ISubscriptionService.getFeatureSubscription} treats one call. */
export interface IFeatureSubscriptionRequestOptions {
  /**
   * Skips the session cache and asks the API again - for "Check again", and for the user
   * coming back from Microsoft Marketplace, when the cached answer is exactly the one
   * that is out of date.
   */
  refresh?: boolean;
}

/**
 * Which paid features the signed-in tenant has bought through Microsoft Marketplace.
 *
 * The API is the source of truth. Payment and plan selection happen in Marketplace, and
 * the API verifies the purchase with Microsoft; this client asks what the API concluded,
 * switches a feature on or off, and hands the API the one-time purchase token Marketplace
 * returned the buyer with, which the API verifies itself - the client never claims a
 * feature is active. The tenant is never sent: the API derives it from the application JWT
 * `ApiService` attaches to every call, so the browser cannot ask about - or claim -
 * another tenant's subscription.
 */
export interface ISubscriptionService {
  /** Where Subscribe and Manage subscription lead in this environment. */
  readonly links: ISubscriptionLinks;

  /**
   * Every paid feature's subscription - `GET api/Subscription/features`.
   *
   * Always a fresh read and never cached: it is the overview, not the gate. Features this
   * build does not know are left out.
   */
  getSubscriptionStatus(): Promise<IFeatureSubscription[]>;

  /**
   * One feature's subscription - `GET api/Subscription/features/{featureCode}`.
   *
   * Cached for the session per feature, and concurrent calls for the same feature share
   * one request, so several gated screens - or one remounted by the router - do not each
   * ask again. A failed call is not cached, so a retry really retries.
   *
   * @throws ApiError - when the call fails, and when the API answers without a
   *   subscription for `feature`. Neither is ever reported as "inactive": a feature the
   *   client could not verify is not a feature the tenant did not buy.
   */
  getFeatureSubscription(
    feature: FeatureCode,
    options?: IFeatureSubscriptionRequestOptions
  ): Promise<IFeatureSubscription>;

  /**
   * Switches a feature on or off for the tenant - `POST api/Subscription/features/{code}/
   * enable` or `…/disable` - and returns the API's answer, which also replaces the cached
   * one.
   *
   * Switching on records the tenant's choice and nothing more: it never marks the feature
   * paid for. An unpaid feature comes back still locked, with `needsMarketplaceCheckout`
   * true, until Microsoft Marketplace reports the purchase to the API.
   *
   * @param returnUrl - sent on switch-on only: the page Marketplace should send the buyer
   *   back to. The API validates it; omitted, the body is `{}`.
   * @throws ApiError - when the call fails or is refused (e.g. a role not allowed to
   *   change features), and when the API answers about a different feature.
   */
  setFeatureEnabled(feature: FeatureCode, isEnabled: boolean, returnUrl?: string): Promise<IFeatureSubscription>;

  /**
   * Hands the API the purchase token Marketplace returned the buyer with -
   * `POST api/Subscription/marketplace/activate`. The API verifies it with Microsoft and
   * answers with the activated feature's subscription, which also replaces that feature's
   * cached answer.
   *
   * The token travels in the body only - never in the URL, never logged.
   *
   * @throws ApiError - when the call fails or is refused, and when the API answers without
   *   a subscription for a feature this build knows.
   */
  activateMarketplacePurchase(token: string): Promise<IFeatureSubscription>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'SubscriptionService';

/**
 * The action on the subscription controller; a feature's own status is one segment below
 * it. The controller route lives in `src/config/environment.ts`.
 */
const FEATURES_ACTION: string = 'features';

/** The actions below one feature that switch it on and off. */
const ENABLE_ACTION: string = 'enable';
const DISABLE_ACTION: string = 'disable';

/** The action that turns a Marketplace purchase token into an active subscription. */
const ACTIVATE_ACTION: string = 'marketplace/activate';

/** Default {@link ISubscriptionService}, built on `BaseApiService`. */
export class SubscriptionService extends BaseApiService implements ISubscriptionService {
  public readonly links: ISubscriptionLinks;

  private readonly _endpoint: string;

  /**
   * In-flight and settled reads, keyed by feature code.
   *
   * Per instance, not module-level: `ServiceFactory` builds one container per web part
   * instance and API, and every entry is the API's answer for the JWT of that session -
   * never a flag the client set.
   */
  private readonly _cache: { [feature: string]: Promise<IFeatureSubscription> } = {};

  /**
   * @param apiService - HTTP gateway used for every call.
   * @param links - where Subscribe and Manage subscription lead; defaults to the
   *   environment configuration's `subscription`.
   * @param endpoint - controller route; defaults to `api.endpoints.subscription`.
   */
  public constructor(
    apiService: IApiService,
    links: ISubscriptionLinks = getEnvironment().subscription,
    endpoint: string = getEnvironment().api.endpoints.subscription
  ) {
    super(apiService, LOG_SOURCE);
    this.links = links;
    this._endpoint = endpoint;
  }

  public async getSubscriptionStatus(): Promise<IFeatureSubscription[]> {
    const relativeUrl: string = this.route(this._endpoint, FEATURES_ACTION);
    const payload: unknown = await this.getData<unknown>(relativeUrl, undefined, {
      operation: 'the feature subscriptions'
    });

    return toFeatureSubscriptions(payload);
  }

  public getFeatureSubscription(
    feature: FeatureCode,
    options?: IFeatureSubscriptionRequestOptions
  ): Promise<IFeatureSubscription> {
    const cached: Promise<IFeatureSubscription> | undefined = this._cache[feature];

    if (cached && !options?.refresh) {
      return cached;
    }

    const request: Promise<IFeatureSubscription> = this._readFeature(feature);
    this._cache[feature] = request;

    // Evicted only if it is still the entry: a refresh started meanwhile owns the slot.
    request.catch((): void => {
      if (this._cache[feature] === request) {
        delete this._cache[feature];
      }
    });

    return request;
  }

  public async setFeatureEnabled(
    feature: FeatureCode,
    isEnabled: boolean,
    returnUrl?: string
  ): Promise<IFeatureSubscription> {
    const relativeUrl: string = this.route(
      this._endpoint,
      `${FEATURES_ACTION}/${encodeURIComponent(feature)}/${isEnabled ? ENABLE_ACTION : DISABLE_ACTION}`
    );

    try {
      const payload: unknown = await this.postData<unknown, { returnUrl?: string }>(
        relativeUrl,
        isEnabled && returnUrl ? { returnUrl } : {},
        {
          operation: `the ${feature} setting`
        }
      );
      const subscription: IFeatureSubscription | undefined = toFeatureSubscription(payload, feature);

      if (!subscription) {
        throw this.refusal(`The API did not report the ${feature} subscription.`, relativeUrl);
      }

      // The answer to a write is the freshest there is; a later read should not undo it.
      this._cache[feature] = Promise.resolve(subscription);
      Log.info(LOG_SOURCE, `${feature} switched ${isEnabled ? 'on' : 'off'}: ${subscription.state}, usable: ${subscription.isActive}.`);

      return subscription;
    } catch (error) {
      this.logFailure(error, `Could not switch ${feature} ${isEnabled ? 'on' : 'off'}.`);
      throw error;
    }
  }

  public async activateMarketplacePurchase(token: string): Promise<IFeatureSubscription> {
    const relativeUrl: string = this.route(this._endpoint, ACTIVATE_ACTION);

    try {
      // No `print`: the request body carries the purchase token.
      const payload: unknown = await this.postData<unknown, { token: string }>(relativeUrl, { token }, {
        operation: 'the Marketplace purchase'
      });
      const subscription: IFeatureSubscription | undefined = toActivatedFeatureSubscription(payload);

      if (!subscription) {
        throw this.refusal('The API did not report the activated subscription.', relativeUrl);
      }

      // Seeded, not merely cleared: an in-flight read could otherwise cache the stale answer.
      this._cache[subscription.feature] = Promise.resolve(subscription);
      Log.info(LOG_SOURCE, `${subscription.feature} activated: ${subscription.state}, usable: ${subscription.isActive}.`);

      return subscription;
    } catch (error) {
      this.logFailure(error, 'Could not activate the Marketplace purchase.');
      throw error;
    }
  }

  private async _readFeature(feature: FeatureCode): Promise<IFeatureSubscription> {
    const relativeUrl: string = this.route(this._endpoint, `${FEATURES_ACTION}/${encodeURIComponent(feature)}`);

    try {
      const payload: unknown = await this.getData<unknown>(relativeUrl, undefined, {
        operation: `the ${feature} subscription`
      });
      const subscription: IFeatureSubscription | undefined = toFeatureSubscription(payload, feature);

      if (!subscription) {
        throw this.refusal(`The API did not report the ${feature} subscription.`, relativeUrl);
      }

      Log.info(LOG_SOURCE, `${feature} subscription: ${subscription.state} (${subscription.subscriptionStatus || 'no status'}).`);

      return subscription;
    } catch (error) {
      this.logFailure(error, `Could not read the ${feature} subscription.`);
      throw error;
    }
  }
}
