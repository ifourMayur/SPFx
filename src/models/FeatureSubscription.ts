/* eslint-disable @rushstack/no-new-null -- the Web API serializes with System.Text.Json
   defaults, so an absent value arrives as an explicit `null` rather than being omitted. */

/**
 * Paid features, and whether the signed-in tenant may use them.
 *
 * Microsoft Marketplace is the payment provider and the XSProject Web API is the source of
 * truth: the API verifies a purchase with Microsoft and stores it against that tenant, and
 * this client only ever *reads* the result. The client does forward the one-time Marketplace
 * token the API's landing page hands back, but only so the API can verify it; it never
 * decides activation. Nothing here can make a feature active - there is no setter, no flag
 * kept in storage, and no reading of the Marketplace return URL as proof of purchase. A UI gate is a convenience,
 * never a security boundary: the API endpoints behind a paid feature must check the
 * subscription themselves.
 *
 * Part of the models layer: no imports from `config`, `services` or `components`.
 */

/**
 * Every feature that can be sold separately, as the Web API names it.
 *
 * Adding a paid feature is one entry here and one in {@link FEATURE_NAMES}; the service,
 * the cache and `FeatureSubscriptionGuard` are all keyed on the code and need no change.
 */
export const FEATURE_CODES = ['ED_CONTROL', 'KYP', 'MINUTES', 'TENDER'] as const;

/** A feature that can be sold separately. */
export type FeatureCode = (typeof FEATURE_CODES)[number];

/** The name a feature is shown under. */
export const FEATURE_NAMES: { readonly [code in FeatureCode]: string } = {
  ED_CONTROL: 'ED Control',
  KYP: 'KYP',
  MINUTES: 'Minutes of Meetings',
  TENDER: 'Tender'
};

/** True when `value` is one of {@link FEATURE_CODES}, compared exactly. */
export function isFeatureCode(value: unknown): value is FeatureCode {
  return typeof value === 'string' && (FEATURE_CODES as readonly string[]).indexOf(value) > -1;
}

/**
 * Where the Marketplace subscription stands, reduced from the API's own status text.
 *
 * Separate from whether the feature may be *used* - that is `IFeatureSubscription.isActive`,
 * which also needs the tenant's own switch to be on. A feature can be paid for and still
 * switched off.
 *
 * - `active` - paid for.
 * - `pending` - bought, not yet activated (Marketplace's `PendingFulfillmentStart`, or the
 *   API still synchronising). Offering Subscribe here would invite a second purchase.
 * - `suspended` - bought, but Marketplace suspended it, usually over payment. The way back
 *   is managing the existing subscription, not buying another.
 * - `inactive` - never bought, or cancelled.
 */
export type SubscriptionState = 'active' | 'pending' | 'suspended' | 'inactive';

/**
 * `GET api/Subscription/features/{featureCode}` - one feature's subscription for the
 * tenant the caller's JWT belongs to. `POST …/enable` and `…/disable` answer the same.
 *
 * The tenant is deliberately not a parameter: the API derives it from the token, so a
 * browser cannot ask about - or claim - another tenant's subscription.
 */
export interface IFeatureSubscriptionResponse {
  feature?: string | null;
  /**
   * The authority: paid for **and** switched on. Only an explicit `true` makes a feature
   * usable.
   */
  isActive?: boolean | null;
  /** The tenant's own on/off switch, whatever the Marketplace status. */
  isEnabled?: boolean | null;
  /** Status text, e.g. `Active`, `Inactive`, `Pending`, `Suspended`. Display only. */
  subscriptionStatus?: string | null;
  /** Marketplace plan id, e.g. `EDCONTROL_BASIC`. */
  planId?: string | null;
  /** The plan's display name, e.g. `Professional`, when the API knows it. */
  planName?: string | null;
  requiresSubscription?: boolean | null;
}

/** One feature's subscription, as the components read it. */
export interface IFeatureSubscription {
  feature: FeatureCode;
  /** Where the Marketplace subscription stands; never used to decide access. */
  state: SubscriptionState;
  /** Whether the feature may be used: the API's `isActive`, and nothing else. */
  isActive: boolean;
  /** The tenant's own switch. On does not mean paid for - see {@link needsMarketplaceCheckout}. */
  isEnabled: boolean;
  /** The API's own status text, for display; never used to decide access. */
  subscriptionStatus: string;
  planId?: string;
  planName?: string;
}

/**
 * Where the Subscribe and Manage subscription actions lead, from the environment
 * configuration. An empty string means "not configured here" and hides that action rather
 * than sending the user somewhere broken.
 */
export interface ISubscriptionLinks {
  /** The XSProject SaaS offer in Microsoft Marketplace, where plans are chosen and bought. */
  readonly marketplaceOfferUrl: string;
  /** Where an existing subscription is changed or cancelled. */
  readonly manageSubscriptionUrl: string;
  /**
   * Optional: a Marketplace link per feature - each feature's own offer - so switching on KYP
   * opens a checkout that offers KYP alone (Marketplace's checkout lists every plan of an
   * offer). A feature without one (or with a link that could not be opened) falls back to
   * `marketplaceOfferUrl`; see {@link offerUrlFor}. The link only chooses where the buyer
   * starts: which feature a purchase unlocks is still decided by the Web API from the offer
   * and plan actually bought.
   */
  readonly featureOfferUrls?: Readonly<Partial<Record<FeatureCode, string>>>;
}

/**
 * Status texts that mean "bought, not active yet", compared case-insensitively.
 * `pendingfulfillmentstart` is Marketplace's own SaaS status for exactly that.
 */
/** Status texts that mean "paid for", compared case-insensitively. Marketplace says `Subscribed`. */
const ACTIVE_STATUSES: string[] = ['active', 'subscribed'];

const PENDING_STATUSES: string[] = ['pending', 'pendingfulfillmentstart', 'pendingactivation', 'processing', 'provisioning'];

/** Status texts that mean "bought, then suspended", compared case-insensitively. */
const SUSPENDED_STATUSES: string[] = ['suspended'];

/**
 * Maps the API's answer for one feature onto {@link IFeatureSubscription}.
 *
 * Two rules keep a paid feature from being switched on by mistake:
 *
 * - **`isActive` decides, and only an explicit `true` counts.** A status text of `Active`
 *   with `isActive` missing or `false` is not usable - paid for, perhaps, but switched off;
 *   the text only says where the subscription stands.
 * - **The answer must be about the feature that was asked for.** A response naming a
 *   different feature - or none - is rejected rather than reused, so an active ED Control
 *   can never unlock Minutes through a mis-routed or mis-cached call.
 *
 * @param response - the payload, already out of its `ResponseDetail` envelope.
 * @param requested - the feature the call asked about.
 * @returns the subscription, or `undefined` when `response` is not an answer about
 *   `requested` - the caller treats that as a failure to verify, never as "inactive".
 */
export function toFeatureSubscription(
  response: unknown,
  requested: FeatureCode
): IFeatureSubscription | undefined {
  if (!response || typeof response !== 'object' || Array.isArray(response)) {
    return undefined;
  }

  const body: IFeatureSubscriptionResponse = response as IFeatureSubscriptionResponse;

  // Compared case-insensitively because ASP.NET routes the same way. Absent is rejected too:
  // the API always names the feature, so a body that does not - a proxy's `{}`, say - is not
  // an answer, and read as "not bought" it would invite a second purchase.
  const feature: string = typeof body.feature === 'string' ? body.feature.trim().toUpperCase() : '';

  if (feature !== requested) {
    return undefined;
  }

  const subscriptionStatus: string = (body.subscriptionStatus || '').trim();
  const isActive: boolean = body.isActive === true;

  return {
    feature: requested,
    state: toSubscriptionState(isActive, subscriptionStatus),
    isActive,
    // A feature in use is switched on, whatever an older API leaves out.
    isEnabled: isActive || body.isEnabled === true,
    subscriptionStatus,
    planId: (body.planId || '').trim() || undefined,
    planName: (body.planName || '').trim() || undefined
  };
}

/**
 * Maps the answer to `POST api/Subscription/marketplace/activate`, which - unlike the
 * per-feature reads - is not addressed to a feature: the purchase decides which one it
 * unlocked, so the feature is read from the answer itself.
 *
 * The answer must name a feature this build knows (any case); otherwise it is rejected
 * rather than guessed at. Everything else is {@link toFeatureSubscription}'s rules, so only
 * an explicit `isActive: true` makes the feature usable.
 *
 * @returns the subscription, or `undefined` when `response` is not an answer about a known
 *   feature - a failure to verify, never "inactive".
 */
export function toActivatedFeatureSubscription(response: unknown): IFeatureSubscription | undefined {
  if (!response || typeof response !== 'object' || Array.isArray(response)) {
    return undefined;
  }

  const code: string = String((response as IFeatureSubscriptionResponse).feature || '').toUpperCase();

  return isFeatureCode(code) ? toFeatureSubscription(response, code) : undefined;
}

/**
 * Maps every entry of `GET api/Subscription/features` the client recognises.
 *
 * An entry for a feature this build does not know is dropped rather than kept under a
 * code nothing can ask for, and so is one that does not parse.
 */
export function toFeatureSubscriptions(response: unknown): IFeatureSubscription[] {
  if (!Array.isArray(response)) {
    return [];
  }

  const subscriptions: IFeatureSubscription[] = [];

  response.forEach((item: unknown): void => {
    const code: string = item && typeof item === 'object'
      ? String((item as IFeatureSubscriptionResponse).feature || '').toUpperCase()
      : '';
    const subscription: IFeatureSubscription | undefined = isFeatureCode(code)
      ? toFeatureSubscription(item, code)
      : undefined;

    if (subscription) {
      subscriptions.push(subscription);
    }
  });

  return subscriptions;
}

function toSubscriptionState(isActive: boolean, subscriptionStatus: string): SubscriptionState {
  if (isActive) {
    return 'active';
  }

  const status: string = subscriptionStatus.toLowerCase();

  if (ACTIVE_STATUSES.indexOf(status) > -1) {
    return 'active';
  }

  if (PENDING_STATUSES.indexOf(status) > -1) {
    return 'pending';
  }

  return SUSPENDED_STATUSES.indexOf(status) > -1 ? 'suspended' : 'inactive';
}

/**
 * Whether switching a feature on has to send the user to Microsoft Marketplace to pay.
 *
 * Only when nothing has been bought: a paid-for feature that was switched off just comes
 * back on, a pending purchase is already on its way, and a suspended one is mended through
 * Manage subscription rather than by buying it again.
 */
export function needsMarketplaceCheckout(subscription: IFeatureSubscription): boolean {
  return subscription.state === 'inactive';
}

/** The plan to show beside an active subscription: its name, else its id. */
export function toPlanLabel(subscription: IFeatureSubscription): string | undefined {
  return subscription.planName || subscription.planId;
}

/**
 * The query parameter the Marketplace landing page adds when it sends the user back, e.g.
 * `…/XSProject.aspx#/subscription?subscription=returned`.
 *
 * It only changes what the UI *says* - "being processed" rather than "subscribe" while the
 * API catches up - and makes the guard skip its cache. It never makes a feature active:
 * anyone can type it into a URL.
 */
export const SUBSCRIPTION_RETURN_PARAMETER: string = 'subscription';

/** The value of {@link SUBSCRIPTION_RETURN_PARAMETER} that marks a return from Marketplace. */
export const SUBSCRIPTION_RETURN_VALUE: string = 'returned';

/**
 * Whether a location's query string marks a return from Microsoft Marketplace.
 *
 * @param search - `location.search` of the routed location, with or without the `?`.
 */
export function isMarketplaceReturn(search: string): boolean {
  const query: string = (search || '').replace(/^\?/, '');

  return query.split('&').some((pair: string): boolean => {
    const separator: number = pair.indexOf('=');
    const name: string = separator > -1 ? pair.substring(0, separator) : pair;
    const value: string = separator > -1 ? pair.substring(separator + 1) : '';

    return safeDecode(name) === SUBSCRIPTION_RETURN_PARAMETER
      && safeDecode(value).toLowerCase() === SUBSCRIPTION_RETURN_VALUE;
  });
}

/**
 * The query parameter the Marketplace landing page puts the one-time purchase token in,
 * e.g. `#/subscription?subscription=returned&marketplaceToken=…`.
 */
export const MARKETPLACE_TOKEN_PARAMETER: string = 'marketplaceToken';

/**
 * The Marketplace token in a location's query string, for the API to verify.
 *
 * Decoded with `decodeURIComponent` only: Marketplace tokens are base64-like, so a literal
 * `+` is data and must not become a space the way a form-encoded query would have it. A
 * value that is not valid percent-encoding is returned as it came.
 *
 * @param search - `location.search` of the routed location, with or without the `?`.
 * @returns the token, or `undefined` when absent or empty. Never log it.
 */
export function readMarketplaceToken(search: string): string | undefined {
  const query: string = (search || '').replace(/^\?/, '');

  for (const pair of query.split('&')) {
    const separator: number = pair.indexOf('=');
    const name: string = separator > -1 ? pair.substring(0, separator) : pair;

    if (safeDecode(name) === MARKETPLACE_TOKEN_PARAMETER) {
      const value: string = separator > -1 ? pair.substring(separator + 1) : '';
      let token: string;

      try {
        token = decodeURIComponent(value);
      } catch {
        token = value;
      }

      return token || undefined;
    }
  }

  return undefined;
}

function safeDecode(value: string): string {
  try {
    return decodeURIComponent(value.replace(/\+/g, ' '));
  } catch {
    return value;
  }
}

/**
 * Whether a configured link can be offered to the user: an absolute `https` URL.
 *
 * Anything else - empty, relative, `http`, `javascript:` - hides the action instead, so a
 * missing or mistyped configuration shows as "not available here" rather than as a button
 * that goes nowhere.
 */
export function isNavigableUrl(url: string | undefined): boolean {
  return /^https:\/\/[^\s/]+/i.test((url || '').trim());
}

/**
 * Where buying `feature` starts: its own plan's link when one is configured and can be
 * opened, otherwise the offer's. An empty answer means neither is configured, which every
 * caller already treats as "subscribing is not available here".
 */
export function offerUrlFor(links: ISubscriptionLinks, feature: FeatureCode): string {
  const featureUrl: string | undefined = links.featureOfferUrls?.[feature];

  return isNavigableUrl(featureUrl) ? (featureUrl as string) : links.marketplaceOfferUrl;
}
