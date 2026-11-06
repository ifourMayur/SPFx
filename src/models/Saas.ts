/**
 * Microsoft Marketplace SaaS subscriptions: the `api/saas` wire contracts and the pure
 * rules the Subscription page and `FeatureGate` share.
 *
 * Part of the models layer: no imports outside `models/`, so every rule here is testable
 * under plain jsdom without a transport or a component.
 *
 * **The client never decides entitlement.** `GET api/saas/entitlements` is evaluated on
 * the server from the caller's verified token, and the Web API enforces the same answer on
 * every gated endpoint (`[RequireSaasFeature]`). Everything below only decides what to
 * *show* - a locked panel instead of a section, an Enable button for an admin - so a wrong
 * answer here can hide or offer a button, never grant access.
 *
 * Contracts: `D:\Projects\TaskXS\BMDeskV2\.claude\docs\Marketplace\saas-implementation-plan.md`,
 * "JSON contracts" -> "SPFx".
 */

/* eslint-disable @rushstack/no-new-null -- These interfaces describe the BMDesk Web API's
   wire format, which is the "legacy API" case the rule exempts: it serializes with
   System.Text.Json defaults, so absent values arrive as an explicit `null` rather than
   being omitted. Declaring them as `| null` is what makes the null checks below type-check
   instead of looking redundant. */

import { ApiError, isApiError } from './ApiError';
import { ROUTE_PATHS } from './Navigation';

/** Feature keys the SPFx screens gate on (`SaasFeatureKeys` on the server). */
export const SAAS_FEATURE_KEYS = {
  edControl: 'EDCONTROL.CORE',
  kyp: 'KYP.CORE',
  minutes: 'MINUTES.CORE',
  tender: 'TENDER.CORE'
} as const;

/** Service keys - what a tenant buys, one Marketplace offer each (`SaasServiceKeys`). */
export const SAAS_SERVICE_KEYS = {
  edControl: 'ED_CONTROL',
  kyp: 'KYP',
  minutes: 'MINUTES',
  tender: 'TENDER'
} as const;

/** How long entitlements are cached when the response does not say (`cacheSeconds`). */
export const DEFAULT_ENTITLEMENT_CACHE_SECONDS: number = 300;

/** A service ending within this many days is flagged as expiring. */
export const EXPIRY_WARNING_DAYS: number = 14;

/** Where a tenant's billing administrator manages Marketplace SaaS subscriptions. */
export const SAAS_MANAGE_BILLING_URL: string = 'https://admin.microsoft.com/Adminportal/Home#/subscriptions';

/**
 * A service's status for the caller's tenant.
 *
 * `NotPurchased` - nothing bought; `Unavailable` - the service is switched off globally on
 * the XSProject side, so it cannot be bought or used by anyone.
 */
export type SaasServiceStatus = 'Active' | 'Pending' | 'Suspended' | 'Inactive' | 'NotPurchased' | 'Unavailable';

/** Where an entitlement comes from. `Included` = the tenant is not billed (billing not enforced). */
export type SaasEntitlementSource = 'Marketplace' | 'Manual' | 'Included';

/** One service in `GET api/saas/entitlements` -> `details`. */
export interface ISaasServiceStatus {
  serviceKey: string;
  name: string;
  description: string | null;
  /** Typed as a string too: an unknown future status must still render, not break. */
  status: SaasServiceStatus | string;
  source: SaasEntitlementSource | string | null;
  validTo: string | null;
  termEnd: string | null;
  isFreeTrial: boolean;
  autoRenew: boolean | null;
  /** The server's answer to "may this caller buy this service now". */
  canPurchase: boolean;
  marketplaceSubscriptionId: string | null;
}

/** `GET api/saas/entitlements`. */
export interface ISaasEntitlements {
  tenantKnown: boolean;
  entraTenantId: string | null;
  /** Global/SharePoint Administrator or XSProject Admin - decided by the server from the token. */
  isTenantAdmin: boolean;
  isBillingEnforced: boolean;
  /** Keys of the services enabled for the tenant. */
  services: string[];
  /** Keys of the features enabled for the tenant - what gates check. */
  features: string[];
  details: ISaasServiceStatus[];
  cacheSeconds: number;
  evaluatedOn: string;
}

export interface ISaasCatalogFeature {
  featureKey: string;
  name: string;
}

export interface ISaasCatalogPlan {
  offerId: string;
  planId: string;
  name: string;
  pricingModel: string | null;
  billingTerm: string | null;
}

export interface ISaasCatalogService {
  serviceKey: string;
  name: string;
  description: string | null;
  displayOrder: number;
  isAvailable: boolean;
  hasOffer: boolean;
  features: ISaasCatalogFeature[];
  plans: ISaasCatalogPlan[];
}

/** `GET api/saas/catalog`. */
export interface ISaasCatalog {
  services: ISaasCatalogService[];
}

/** Body of `POST api/saas/purchase-intents`. */
export interface ISaasPurchaseIntentRequest {
  serviceKey: string;
  /** Where the landing page sends the buyer back to - this page, on one of the tenant's hosts. */
  returnUrl: string;
}

/** Answer of `POST api/saas/purchase-intents`. */
export interface ISaasPurchaseIntent {
  intentId: string;
  marketplaceOfferUrl: string;
  expiresOn: string;
}

/**
 * A Marketplace subscription's status, in Microsoft's own words (`SaasSubscriptionStatus`
 * on the server). `PendingFulfillmentStart` - bought but not activated yet, so nothing is
 * billed or unlocked until it is activated (automatically, on the return from the Marketplace).
 */
export type SaasSubscriptionStatus = 'NotStarted' | 'PendingFulfillmentStart' | 'Subscribed' | 'Suspended' | 'Unsubscribed';

/** One row of `GET api/saas/subscriptions`. */
export interface ISaasSubscription {
  marketplaceSubscriptionId: string;
  subscriptionName: string | null;
  offerId: string;
  planId: string;
  planName: string | null;
  quantity: number | null;
  /** Typed as a string too: an unknown future status must still render, not break. */
  status: SaasSubscriptionStatus | string;
  services: string[];
  termStart: string | null;
  termEnd: string | null;
  termUnit: string | null;
  autoRenew: boolean | null;
  isFreeTrial: boolean;
}

/**
 * What `POST api/saas/subscriptions/{id}/activate` did.
 *
 * `Activated` - this call activated it; `AlreadyActive` - it already was, nothing was done
 * twice; `Suspended` / `Unsubscribed` - Microsoft says so, nothing was activated.
 */
export type SaasActivationState = 'Activated' | 'AlreadyActive' | 'Suspended' | 'Unsubscribed';

/** Answer of `POST api/saas/subscriptions/{marketplaceSubscriptionId}/activate`. */
export interface ISaasActivationResult {
  /** Typed as a string too: an unknown future state is shown as a warning, not dropped. */
  state: SaasActivationState | string;
  subscription: ISaasSubscription;
  /** Keys of the services the subscription's plan grants. */
  services: string[];
  /** The tenant's entitlements after the call - the shape of `GET api/saas/entitlements`. */
  entitlements: ISaasEntitlements;
  /** The server's user-facing text for `state`. */
  message: string;
  correlationId: string | null;
}

/** One row of the Subscription page: a catalog service merged with its status. */
export interface ISaasServiceRow {
  serviceKey: string;
  name: string;
  description: string;
  displayOrder: number;
  status: SaasServiceStatus | string;
  source: string | null;
  validTo: string | null;
  termEnd: string | null;
  isFreeTrial: boolean;
  autoRenew: boolean | null;
  canPurchase: boolean;
  plans: ISaasCatalogPlan[];
}

/** Visual weight of a status badge or banner. */
export type SaasTone = 'success' | 'warning' | 'error' | 'neutral';

/** A failure, reduced to what the UI shows. */
export interface ISaasErrorInfo {
  message: string;
  code?: string;
  /** The server's reference for the failure, shown so support can find it. */
  correlationId?: string;
}

/** What the Subscription route's hash query says. */
export interface ISaasReturnQuery {
  /** `saas=returned` - the buyer came back from the Marketplace, with or without buying. */
  isPurchaseReturn: boolean;
  /** The service the purchase was for, or the one a locked `FeatureGate` asked for. */
  serviceKey: string | undefined;
  /**
   * The Marketplace subscription the landing page resolved the purchase to (a GUID), read
   * only on a purchase return. Absent when the buyer did not complete the checkout.
   */
  subscriptionId: string | undefined;
}

// ---------------------------------------------------------------------------------------
// Wire normalisation
// ---------------------------------------------------------------------------------------

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function text(value: unknown): string | null {
  return typeof value === 'string' ? value : null;
}

function flag(value: unknown): boolean {
  return value === true;
}

function textList(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item: unknown): item is string => typeof item === 'string') : [];
}

function objectList(value: unknown): Record<string, unknown>[] {
  return Array.isArray(value) ? value.filter(isObject) : [];
}

function toServiceStatus(raw: Record<string, unknown>): ISaasServiceStatus {
  const serviceKey: string = text(raw.serviceKey) || '';

  return {
    serviceKey,
    name: text(raw.name) || serviceKey,
    description: text(raw.description),
    status: text(raw.status) || 'NotPurchased',
    source: text(raw.source),
    validTo: text(raw.validTo),
    termEnd: text(raw.termEnd),
    isFreeTrial: flag(raw.isFreeTrial),
    autoRenew: typeof raw.autoRenew === 'boolean' ? raw.autoRenew : null,
    canPurchase: flag(raw.canPurchase),
    marketplaceSubscriptionId: text(raw.marketplaceSubscriptionId)
  };
}

/**
 * The entitlements response with every field defaulted, or `undefined` when the body is
 * not an object at all (an HTML error page answered with a 200, say).
 *
 * Defaults are the **closed** answer: a missing `isTenantAdmin` is `false`, missing
 * `features` is none. A missing or non-positive `cacheSeconds` becomes
 * {@link DEFAULT_ENTITLEMENT_CACHE_SECONDS}.
 */
export function toSaasEntitlements(payload: unknown): ISaasEntitlements | undefined {
  if (!isObject(payload)) {
    return undefined;
  }

  const cacheSeconds: unknown = payload.cacheSeconds;

  return {
    tenantKnown: flag(payload.tenantKnown),
    entraTenantId: text(payload.entraTenantId),
    isTenantAdmin: flag(payload.isTenantAdmin),
    isBillingEnforced: flag(payload.isBillingEnforced),
    services: textList(payload.services),
    features: textList(payload.features),
    details: objectList(payload.details).map(toServiceStatus).filter((detail: ISaasServiceStatus) => !!detail.serviceKey),
    cacheSeconds:
      typeof cacheSeconds === 'number' && isFinite(cacheSeconds) && cacheSeconds > 0
        ? cacheSeconds
        : DEFAULT_ENTITLEMENT_CACHE_SECONDS,
    evaluatedOn: text(payload.evaluatedOn) || ''
  };
}

/** The catalog response with every field defaulted, or `undefined` when it is not an object. */
export function toSaasCatalog(payload: unknown): ISaasCatalog | undefined {
  if (!isObject(payload)) {
    return undefined;
  }

  return {
    services: objectList(payload.services)
      .map((raw: Record<string, unknown>): ISaasCatalogService => {
        const serviceKey: string = text(raw.serviceKey) || '';
        return {
          serviceKey,
          name: text(raw.name) || serviceKey,
          description: text(raw.description),
          displayOrder: typeof raw.displayOrder === 'number' ? raw.displayOrder : 0,
          isAvailable: flag(raw.isAvailable),
          hasOffer: flag(raw.hasOffer),
          features: objectList(raw.features).map((feature: Record<string, unknown>) => ({
            featureKey: text(feature.featureKey) || '',
            name: text(feature.name) || ''
          })),
          plans: objectList(raw.plans).map((plan: Record<string, unknown>) => ({
            offerId: text(plan.offerId) || '',
            planId: text(plan.planId) || '',
            name: text(plan.name) || text(plan.planId) || '',
            pricingModel: text(plan.pricingModel),
            billingTerm: text(plan.billingTerm)
          }))
        };
      })
      .filter((service: ISaasCatalogService) => !!service.serviceKey)
  };
}

function toSaasSubscription(raw: Record<string, unknown>): ISaasSubscription {
  return {
    marketplaceSubscriptionId: text(raw.marketplaceSubscriptionId) || '',
    subscriptionName: text(raw.subscriptionName),
    offerId: text(raw.offerId) || '',
    planId: text(raw.planId) || '',
    planName: text(raw.planName),
    quantity: typeof raw.quantity === 'number' ? raw.quantity : null,
    status: text(raw.status) || '',
    services: textList(raw.services),
    termStart: text(raw.termStart),
    termEnd: text(raw.termEnd),
    termUnit: text(raw.termUnit),
    autoRenew: typeof raw.autoRenew === 'boolean' ? raw.autoRenew : null,
    isFreeTrial: flag(raw.isFreeTrial)
  };
}

/**
 * The subscriptions response with every field defaulted, or `undefined` when it is not an
 * array. A row without a `marketplaceSubscriptionId` is dropped: there is nothing to show
 * or activate for it.
 */
export function toSaasSubscriptions(payload: unknown): ISaasSubscription[] | undefined {
  if (!Array.isArray(payload)) {
    return undefined;
  }

  return objectList(payload)
    .map(toSaasSubscription)
    .filter((subscription: ISaasSubscription) => !!subscription.marketplaceSubscriptionId);
}

/**
 * The activation response with every field defaulted, or `undefined` when it is not the
 * promised shape: an object with a `state`, a `subscription` naming its id, and
 * `entitlements` that are an object. Those are what the page and the entitlements cache act
 * on, so a body missing any of them is rejected rather than half-applied.
 */
export function toSaasActivationResult(payload: unknown): ISaasActivationResult | undefined {
  if (!isObject(payload) || !isObject(payload.subscription)) {
    return undefined;
  }

  const state: string | null = text(payload.state);
  const subscription: ISaasSubscription = toSaasSubscription(payload.subscription);
  const entitlements: ISaasEntitlements | undefined = toSaasEntitlements(payload.entitlements);

  if (!state || !subscription.marketplaceSubscriptionId || !entitlements) {
    return undefined;
  }

  return {
    state,
    subscription,
    services: textList(payload.services),
    entitlements,
    message: text(payload.message) || '',
    correlationId: text(payload.correlationId)
  };
}

/** Whether the subscription is active after the call - activated by it or already before. */
export function isActivationSuccessful(result: Pick<ISaasActivationResult, 'state'>): boolean {
  return result.state === 'Activated' || result.state === 'AlreadyActive';
}

// ---------------------------------------------------------------------------------------
// Gating
// ---------------------------------------------------------------------------------------

function sameKey(left: string, right: string): boolean {
  return left.toUpperCase() === right.toUpperCase();
}

/**
 * Whether the server reported `featureKey` as enabled for the tenant.
 *
 * `false` for no entitlements at all (not loaded, or the load failed): a gate fails
 * closed. Compared case-insensitively, as the server's keys are.
 */
export function isFeatureEnabled(entitlements: ISaasEntitlements | undefined, featureKey: string): boolean {
  if (!entitlements || !featureKey) {
    return false;
  }

  return entitlements.features.some((key: string) => sameKey(key, featureKey));
}

/** The status of one service, or `undefined` when the server did not report it. */
export function findServiceStatus(
  entitlements: ISaasEntitlements | undefined,
  serviceKey: string | undefined
): ISaasServiceStatus | undefined {
  if (!entitlements || !serviceKey) {
    return undefined;
  }

  return entitlements.details.filter((detail: ISaasServiceStatus) => sameKey(detail.serviceKey, serviceKey))[0];
}

// ---------------------------------------------------------------------------------------
// Presentation
// ---------------------------------------------------------------------------------------

const STATUS_LABELS: { [status: string]: string } = {
  Active: 'Active',
  Pending: 'Not activated',
  Suspended: 'Suspended',
  Inactive: 'Inactive',
  NotPurchased: 'Not enabled',
  Unavailable: 'Unavailable'
};

const STATUS_TONES: { [status: string]: SaasTone } = {
  Active: 'success',
  Pending: 'warning',
  Suspended: 'error',
  Inactive: 'neutral',
  NotPurchased: 'neutral',
  Unavailable: 'neutral'
};

/** Badge text for a status. An unknown status is shown as the server wrote it. */
export function getServiceStatusLabel(status: string): string {
  return STATUS_LABELS[status] || status || 'Unknown';
}

/** Badge colour for a status. */
export function getServiceStatusTone(status: string): SaasTone {
  return STATUS_TONES[status] || 'neutral';
}

const SOURCE_LABELS: { [source: string]: string } = {
  Marketplace: 'Microsoft Marketplace',
  Manual: 'Granted by XSProject',
  Included: 'Included in your XSProject agreement'
};

/** Where an entitlement comes from, in words; `undefined` when there is none. */
export function getSourceLabel(source: string | null | undefined): string | undefined {
  if (!source) {
    return undefined;
  }

  return SOURCE_LABELS[source] || source;
}

const SUBSCRIPTION_STATUS_LABELS: { [status: string]: string } = {
  NotStarted: 'Not started',
  PendingFulfillmentStart: 'Awaiting activation',
  Subscribed: 'Active',
  Suspended: 'Suspended',
  Unsubscribed: 'Cancelled'
};

const SUBSCRIPTION_STATUS_TONES: { [status: string]: SaasTone } = {
  NotStarted: 'neutral',
  PendingFulfillmentStart: 'warning',
  Subscribed: 'success',
  Suspended: 'error',
  Unsubscribed: 'neutral'
};

/** Badge text for a Marketplace subscription's status; an unknown one is shown as the server wrote it. */
export function getSubscriptionStatusLabel(status: string): string {
  return SUBSCRIPTION_STATUS_LABELS[status] || status || 'Unknown';
}

/** Badge colour for a Marketplace subscription's status. */
export function getSubscriptionStatusTone(status: string): SaasTone {
  return SUBSCRIPTION_STATUS_TONES[status] || 'neutral';
}

/** A date from the wire, or `undefined` when absent or unparseable. */
export function parseSaasDate(value: string | null | undefined): Date | undefined {
  if (!value) {
    return undefined;
  }

  const date: Date = new Date(value);
  return isNaN(date.getTime()) ? undefined : date;
}

/** A wire date for display, e.g. `12 Nov 2026`; empty when absent or unparseable. */
export function formatSaasDate(value: string | null | undefined): string {
  const date: Date | undefined = parseSaasDate(value);
  return date ? date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '';
}

/**
 * The date a service stops, when one is coming.
 *
 * `validTo` always ends access. `termEnd` only does when the subscription will not
 * renew - an auto-renewing term end is a renewal date, not an end, and flagging it as
 * "expiring" every month would train admins to ignore the warning.
 */
export function getServiceEndDate(
  detail: Pick<ISaasServiceStatus, 'validTo' | 'termEnd' | 'autoRenew'>
): Date | undefined {
  return parseSaasDate(detail.validTo) || (detail.autoRenew === true ? undefined : parseSaasDate(detail.termEnd));
}

/** Whether an active service ends within {@link EXPIRY_WARNING_DAYS} of `now`. */
export function isExpiringSoon(
  detail: Pick<ISaasServiceStatus, 'status' | 'validTo' | 'termEnd' | 'autoRenew'>,
  now: Date
): boolean {
  if (detail.status !== 'Active') {
    return false;
  }

  const end: Date | undefined = getServiceEndDate(detail);
  if (!end) {
    return false;
  }

  const remaining: number = end.getTime() - now.getTime();
  return remaining >= 0 && remaining <= EXPIRY_WARNING_DAYS * 24 * 60 * 60 * 1000;
}

/**
 * The sentence explaining a service's state, or `undefined` when its badge says enough.
 *
 * Shared by the Subscription page and the locked `FeatureGate`, so the two never explain
 * the same state differently.
 */
export function getServiceStatusMessage(
  detail: Pick<ISaasServiceStatus, 'name' | 'status' | 'validTo' | 'termEnd' | 'autoRenew'>,
  now: Date
): string | undefined {
  switch (detail.status) {
    case 'Suspended':
      return `${detail.name} is suspended because of a payment problem. A billing administrator can fix the payment method for this subscription in the Microsoft 365 admin center; the service comes back as soon as Microsoft confirms the payment.`;
    case 'Inactive':
      return `${detail.name} is no longer active - its subscription was cancelled or has ended. Your data is kept; enable the service again to use it.`;
    case 'Unavailable':
      return `${detail.name} is currently unavailable. Please try again later or contact XSProject support.`;
    case 'Pending':
      return `${detail.name} was bought but never activated, so it is not billed and cannot be used yet. An administrator can enable it again; the new purchase is activated automatically when you come back from Microsoft Marketplace.`;
    case 'Active':
      if (isExpiringSoon(detail, now)) {
        const end: Date = getServiceEndDate(detail) as Date;
        return `${detail.name} ends on ${formatSaasDate(end.toISOString())}. Renew it before then to keep using it.`;
      }
      return undefined;
    default:
      return undefined;
  }
}

/**
 * The Subscription page's rows: catalog services in `displayOrder`, each merged with its
 * status. A service the server reported without a catalog entry is still listed (after
 * the catalog ones), and a catalog service without a status is `NotPurchased` - or
 * `Unavailable` when the catalog says so - and never purchasable, since only the server
 * decides `canPurchase`.
 */
export function mergeServiceRows(
  catalog: ISaasCatalog | undefined,
  entitlements: ISaasEntitlements | undefined
): ISaasServiceRow[] {
  const rows: ISaasServiceRow[] = [];
  const seen: { [key: string]: boolean } = {};
  const catalogServices: ISaasCatalogService[] = catalog
    ? catalog.services.slice().sort((a: ISaasCatalogService, b: ISaasCatalogService) => a.displayOrder - b.displayOrder)
    : [];

  catalogServices.forEach((service: ISaasCatalogService): void => {
    const detail: ISaasServiceStatus | undefined = findServiceStatus(entitlements, service.serviceKey);
    seen[service.serviceKey.toUpperCase()] = true;

    rows.push({
      serviceKey: service.serviceKey,
      name: detail?.name || service.name,
      description: service.description || detail?.description || '',
      displayOrder: service.displayOrder,
      status: detail ? detail.status : service.isAvailable ? 'NotPurchased' : 'Unavailable',
      source: detail ? detail.source : null,
      validTo: detail ? detail.validTo : null,
      termEnd: detail ? detail.termEnd : null,
      isFreeTrial: detail ? detail.isFreeTrial : false,
      autoRenew: detail ? detail.autoRenew : null,
      canPurchase: detail ? detail.canPurchase : false,
      plans: service.plans
    });
  });

  (entitlements ? entitlements.details : []).forEach((detail: ISaasServiceStatus): void => {
    if (seen[detail.serviceKey.toUpperCase()]) {
      return;
    }

    rows.push({
      serviceKey: detail.serviceKey,
      name: detail.name,
      description: detail.description || '',
      displayOrder: Number.MAX_SAFE_INTEGER,
      status: detail.status,
      source: detail.source,
      validTo: detail.validTo,
      termEnd: detail.termEnd,
      isFreeTrial: detail.isFreeTrial,
      autoRenew: detail.autoRenew,
      canPurchase: detail.canPurchase,
      plans: []
    });
  });

  return rows;
}

// ---------------------------------------------------------------------------------------
// The purchase round trip
// ---------------------------------------------------------------------------------------

/**
 * The URL the landing page sends the buyer back to: this page, on the Subscription route,
 * marked as a return for `serviceKey`.
 *
 * The page's own path and query are kept (SharePoint owns them) and only the fragment is
 * replaced - the app routes in the fragment. Once Microsoft has sent the buyer to the
 * XSProject landing page, it resolves the purchase and redirects here with
 * `subscriptionId={marketplace subscription id}` added to the same fragment query; the
 * Subscription page then activates that subscription itself.
 */
export function buildPurchaseReturnUrl(currentUrl: string, serviceKey: string): string {
  const hashIndex: number = currentUrl.indexOf('#');
  const base: string = hashIndex >= 0 ? currentUrl.substring(0, hashIndex) : currentUrl;

  return `${base}#${ROUTE_PATHS.subscription}?saas=returned&service=${encodeURIComponent(serviceKey)}`;
}

function decodeQueryPart(value: string): string {
  try {
    return decodeURIComponent(value.replace(/\+/g, ' '));
  } catch {
    return value;
  }
}

const GUID_PATTERN: RegExp = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Reads the Subscription route's query: `?saas=returned&service=TENDER&subscriptionId={guid}`
 * from the landing page, `?saas=returned&service=TENDER` when the buyer came back without
 * buying (browser Back, a cancelled checkout), or `?service=TENDER` from a locked
 * `FeatureGate`.
 *
 * The first value of each parameter wins. `subscriptionId` is read only on a purchase
 * return and only when it is a GUID - anything else is ignored rather than sent to the
 * activate endpoint.
 */
export function parseSaasReturnQuery(search: string | undefined): ISaasReturnQuery {
  let isPurchaseReturn: boolean = false;
  let serviceKey: string | undefined;
  let subscriptionId: string | undefined;

  (search || '')
    .replace(/^\?/, '')
    .split('&')
    .forEach((part: string): void => {
      const separator: number = part.indexOf('=');
      const name: string = decodeQueryPart(separator >= 0 ? part.substring(0, separator) : part);
      const value: string = decodeQueryPart(separator >= 0 ? part.substring(separator + 1) : '').trim();

      if (name === 'saas' && value === 'returned') {
        isPurchaseReturn = true;
      } else if (name === 'service' && value && !serviceKey) {
        serviceKey = value;
      } else if (name === 'subscriptionId' && GUID_PATTERN.test(value) && !subscriptionId) {
        subscriptionId = value;
      }
    });

  return { isPurchaseReturn, serviceKey, subscriptionId: isPurchaseReturn ? subscriptionId : undefined };
}

/**
 * Whether a URL fragment is the return from a Marketplace purchase - the Subscription
 * route with `saas=returned` in its query, e.g.
 * `#/subscription?saas=returned&service=TENDER&subscriptionId={guid}`.
 *
 * The return reloads the SharePoint page, so the web part signs in and asks for a site
 * again before its router (and the fragment's route) is mounted; `XsProject` uses this to
 * reopen the site the purchase was started from instead.
 */
export function isSaasReturnFragment(hash: string | undefined): boolean {
  const fragment: string = (hash || '').replace(/^#/, '');
  const queryIndex: number = fragment.indexOf('?');

  if (queryIndex < 0 || fragment.substring(0, queryIndex) !== ROUTE_PATHS.subscription) {
    return false;
  }

  return parseSaasReturnQuery(fragment.substring(queryIndex)).isPurchaseReturn;
}

/** Only an `https:` offer URL is navigated to; anything else is reported, not followed. */
export function isSafeOfferUrl(url: string | null | undefined): boolean {
  return typeof url === 'string' && /^https:\/\/[^\s]+$/i.test(url);
}

// ---------------------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------------------

/** Friendly text for the `SaasErrorCodes` the SPFx screens can meet. */
const ERROR_MESSAGES: { [code: string]: string } = {
  SAAS_NOT_AUTHENTICATED: 'Your XSProject session has expired. Reload the page to sign in again.',
  SAAS_NOT_TENANT_ADMIN:
    'Only a Global Administrator or SharePoint Administrator of your organisation can enable, activate or change services - ask one of them to open this page.',
  SAAS_VALIDATION: 'XSProject could not accept this request. Reload the page and try again; if it keeps happening, contact XSProject support.',
  SAAS_WRONG_TENANT:
    'This Microsoft Marketplace purchase belongs to another organisation, so it cannot be activated here. Sign in with an account of the organisation that bought it.',
  SAAS_NOT_FOUND:
    'Microsoft does not know this subscription. Check the purchase in the Microsoft 365 admin center, or contact XSProject support.',
  SAAS_UNKNOWN_PLAN: 'The plan bought for this subscription is not known to XSProject yet. Please contact XSProject support.',
  SAAS_ACTIVATION_FAILED: 'The subscription could not be activated. Please try again in a moment.',
  SAAS_MARKETPLACE_ERROR: 'Microsoft Marketplace did not answer as expected. Please try again in a moment.',
  SAAS_SERVICE_NOT_PURCHASABLE:
    'This service cannot be bought from Microsoft Marketplace yet. Please contact XSProject support.',
  SAAS_SERVICE_UNKNOWN: 'This service is not known to XSProject. Please reload the page and try again.',
  SAAS_RETURN_URL_INVALID:
    "This page's address is not registered for your organisation, so Microsoft could not send you back here after the purchase. Open XSProject from your organisation's SharePoint site and try again.",
  SAAS_TENANT_UNKNOWN:
    'Your organisation is not registered with XSProject yet. Reload the page to sign in again, or contact XSProject support.',
  SAAS_FEATURE_NOT_ENTITLED: 'This feature is not enabled for your organisation.',
  SAAS_UNAVAILABLE: 'The subscription service is temporarily unavailable. Please try again in a few minutes.',
  SAAS_NOT_CONFIGURED: 'Subscriptions are not set up for XSProject yet. Please contact XSProject support.',
  SAAS_INTERNAL: 'Something went wrong in XSProject. Please try again; if it keeps happening, contact XSProject support.'
};

/**
 * Reduces any failure to the message, code and reference the UI shows.
 *
 * A known `code` wins; otherwise the transport decides - unreachable (status 0), 401, 403
 * - and only then the server's own message.
 */
export function getSaasErrorInfo(error: unknown): ISaasErrorInfo {
  if (!isApiError(error)) {
    return {
      message:
        error instanceof Error && error.message
          ? error.message
          : 'Something went wrong while checking your subscription. Please try again.'
    };
  }

  const apiError: ApiError = error;
  const known: string | undefined = apiError.code ? ERROR_MESSAGES[apiError.code] : undefined;
  let message: string;

  if (known) {
    message = known;
  } else if (apiError.status === 0) {
    message = 'The XSProject service could not be reached. Check your connection and try again in a few minutes.';
  } else if (apiError.status === 401) {
    message = ERROR_MESSAGES.SAAS_NOT_AUTHENTICATED;
  } else if (apiError.status === 403) {
    message = 'Your account is not allowed to view or change subscriptions.';
  } else if (apiError.status >= 500) {
    message = ERROR_MESSAGES.SAAS_INTERNAL;
  } else {
    message = apiError.message || 'The subscription request could not be completed.';
  }

  return { message, code: apiError.code, correlationId: apiError.correlationId };
}

/** Codes for a failure that may well succeed if the same call is simply repeated. */
const RETRYABLE_CODES: string[] = ['SAAS_ACTIVATION_FAILED', 'SAAS_MARKETPLACE_ERROR', 'SAAS_UNAVAILABLE', 'SAAS_INTERNAL'];

/**
 * Whether repeating the call that failed could help - a temporary Marketplace or server
 * failure, or an API that could not be reached - so the UI offers "Try again". A refusal
 * (not an admin, another organisation's purchase, an unknown plan) or a service that is
 * not set up would fail the same way again, so it is not offered one.
 */
export function isRetryableSaasError(error: unknown): boolean {
  if (!isApiError(error)) {
    return true;
  }

  if (error.code) {
    return RETRYABLE_CODES.indexOf(error.code) >= 0;
  }

  return error.status === 0 || error.status >= 500;
}
