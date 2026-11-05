import * as React from 'react';

import styles from './Subscription.module.scss';
import type { ISubscriptionProps } from './ISubscriptionProps';
import type { ActivationBanner, ISubscriptionState } from './ISubscriptionState';
import {
  ISaasActivationResult,
  ISaasCatalog,
  ISaasCatalogPlan,
  ISaasEntitlements,
  ISaasErrorInfo,
  ISaasPurchaseIntent,
  ISaasReturnQuery,
  ISaasServiceRow,
  ISaasServiceStatus,
  ISaasSubscription,
  SAAS_MANAGE_BILLING_URL,
  SaasTone,
  buildPurchaseReturnUrl,
  canActivateSubscription,
  findServiceStatus,
  formatSaasDate,
  getSaasErrorInfo,
  getServiceStatusLabel,
  getServiceStatusMessage,
  getServiceStatusTone,
  getSourceLabel,
  getSubscriptionStatusLabel,
  getSubscriptionStatusTone,
  isActivationSuccessful,
  isRetryableSaasError,
  isSafeOfferUrl,
  mergeServiceRows,
  parseSaasReturnQuery
} from '../../../../models/Saas';

const TONE_CLASSES: { [tone in SaasTone]: string } = {
  success: styles.toneSuccess,
  warning: styles.toneWarning,
  error: styles.toneError,
  neutral: styles.toneNeutral
};

/**
 * The Subscription page (`/subscription`): which Microsoft Marketplace services the
 * organisation has, its Marketplace subscriptions, and - for admins - enabling a service
 * and activating what was bought.
 *
 * Every service row is a catalog service merged with the server's status for this tenant
 * (`mergeServiceRows`). The page always asks the server afresh on load rather than using
 * the cached entitlements: it is the one screen whose job is the current status.
 *
 * **Enabling** asks the Web API for a purchase intent, with this page as the return URL
 * (`#/subscription?saas=returned&service={key}`), then leaves for the Marketplace offer in
 * the same tab. After the checkout Microsoft sends the buyer to the XSProject landing page,
 * which only resolves the purchase and redirects back here with
 * `subscriptionId={marketplace subscription id}` added.
 *
 * **Activating** - on a return carrying `subscriptionId`, the page activates that
 * subscription itself, with the SPFx sign-in it already has (`activateSubscription`), and
 * shows the outcome: success naming the services, a warning with the server's message when
 * Microsoft reports the subscription suspended or cancelled, or the failure with its
 * reference and - when repeating could help - "Try again". The answer carries the fresh
 * entitlements, which the service also writes into its cache, so every `FeatureGate`
 * unlocks at once. A return without `subscriptionId` (browser Back, a cancelled checkout)
 * activates nothing. A subscription still awaiting activation - bought outside SPFx, or an
 * activation that failed - is listed with an "Activate" button for tenant admins, which
 * calls the same endpoint. The query is removed once read, so a reload doesn't activate
 * again.
 *
 * Who may buy or activate is decided by the server (`isTenantAdmin`, `canPurchase`);
 * everyone else sees the same page read-only. Hiding a button is a convenience - the
 * endpoints refuse non-admins themselves.
 */
export default class Subscription extends React.Component<ISubscriptionProps, ISubscriptionState> {
  private _isActive: boolean = false;
  /** The catalog from the last load; kept so new entitlements can be re-merged without refetching it. */
  private _catalog: ISaasCatalog | undefined;

  public constructor(props: ISubscriptionProps) {
    super(props);

    this.state = { isLoading: true, isRefreshing: false, rows: [], subscriptions: [] };
  }

  public componentDidMount(): void {
    this._isActive = true;

    // Read once: the query is cleared below, and the page must not react to that change.
    const query: ISaasReturnQuery = parseSaasReturnQuery(this.props.query);

    if (this.props.query) {
      // React Router ignores a navigate() issued before the route's own effects have run,
      // which componentDidMount precedes - so the query is dropped a microtask later.
      Promise.resolve().then(this._clearQuery, this._clearQuery);
    }

    if (query.serviceKey) {
      this.setState({ highlightKey: query.serviceKey });
    }

    if (query.subscriptionId) {
      // The page loads once the activation has answered, from the entitlements it returns.
      this._activate(query.subscriptionId);
    } else {
      if (query.isPurchaseReturn) {
        this.setState({ banner: { kind: 'returned' } });
      }
      this._load();
    }
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<ISubscriptionProps> {
    const { isLoading, isRefreshing, loadError, entitlements, activatingId, banner } = this.state;

    return (
      <section className={styles.subscription}>
        <div className={styles.header}>
          <h2 className={styles.title}>Subscription</h2>
          {!isLoading && !loadError && (
            <button
              type="button"
              className={styles.secondaryButton}
              disabled={isRefreshing || !!activatingId}
              onClick={this._onRefresh}
            >
              {isRefreshing ? 'Refreshing…' : 'Refresh'}
            </button>
          )}
        </div>
        <p className={styles.intro}>
          The XSProject services your organisation can use. Services are bought through Microsoft
          Marketplace and billed by Microsoft.
        </p>

        {banner && this._renderBanner(banner)}

        {isLoading && (!banner || banner.kind !== 'activating') && (
          <div className={styles.loading} role="status" aria-live="polite">
            <span className={styles.spinner} aria-hidden="true" />
            <span>Loading your subscriptions&hellip;</span>
          </div>
        )}

        {!isLoading && loadError && (
          <div className={styles.errorBox} role="alert">
            <p className={styles.errorTitle}>Your subscriptions could not be loaded.</p>
            {this._renderErrorInfo(loadError)}
            <button type="button" className={styles.primaryButton} onClick={this._onRetryLoad}>
              Try again
            </button>
          </div>
        )}

        {!isLoading && !loadError && entitlements && this._renderContent(entitlements)}
      </section>
    );
  }

  private _renderContent(entitlements: ISaasEntitlements): React.ReactElement {
    const { rows, catalogError, actionError } = this.state;

    return (
      <>
        {!entitlements.tenantKnown && (
          <div className={styles.notice} role="status">
            Your organisation is not linked to XSProject yet, so no services can be enabled. Reload
            the page to sign in again, or contact XSProject support.
          </div>
        )}

        {entitlements.tenantKnown && !entitlements.isBillingEnforced && (
          <div className={styles.notice} role="status">
            Your organisation&apos;s XSProject agreement includes these services, so nothing needs to
            be bought in Microsoft Marketplace.
          </div>
        )}

        {entitlements.isTenantAdmin ? (
          <p className={styles.adminNote}>
            You can enable services and manage their billing.{' '}
            <a href={SAAS_MANAGE_BILLING_URL} target="_blank" rel="noopener noreferrer">
              Manage billing
            </a>{' '}
            opens your subscriptions in the Microsoft 365 admin center.
          </p>
        ) : (
          <div className={styles.notice} role="note">
            You can see your organisation&apos;s services here. Only a Global Administrator or SharePoint
            Administrator can enable services or change billing - ask one of them to open this page.
          </div>
        )}

        {actionError && (
          <div className={styles.errorBox} role="alert">
            {this._renderErrorInfo(actionError)}
          </div>
        )}

        {catalogError && (
          <div className={styles.notice} role="status">
            The service catalog could not be loaded, so plans are not shown. {catalogError.message}
          </div>
        )}

        {rows.length === 0 ? (
          <div className={styles.notice} role="status">No services are available yet.</div>
        ) : (
          <ul className={styles.list}>{rows.map((row: ISaasServiceRow) => this._renderRow(row, entitlements))}</ul>
        )}

        {this._renderSubscriptions(entitlements)}
      </>
    );
  }

  private _renderRow(row: ISaasServiceRow, entitlements: ISaasEntitlements): React.ReactElement {
    const { enablingKey, activatingId, highlightKey } = this.state;
    const tone: SaasTone = getServiceStatusTone(row.status);
    const message: string | undefined = getServiceStatusMessage(row, new Date());
    const source: string | undefined = getSourceLabel(row.source);
    const isHighlighted: boolean = !!highlightKey && highlightKey.toUpperCase() === row.serviceKey.toUpperCase();
    const canEnable: boolean = entitlements.isTenantAdmin && row.canPurchase;

    return (
      <li
        key={row.serviceKey}
        className={`${styles.card} ${isHighlighted ? styles.highlighted : ''}`}
        data-service-key={row.serviceKey}
      >
        <div className={styles.cardHeader}>
          <h3 className={styles.serviceName}>{row.name}</h3>
          <span className={`${styles.badge} ${TONE_CLASSES[tone]}`} data-status={row.status}>
            {getServiceStatusLabel(row.status)}
          </span>
          {row.isFreeTrial && <span className={styles.trial}>Free trial</span>}
        </div>

        {row.description && <p className={styles.description}>{row.description}</p>}

        {(source || row.validTo || row.termEnd) && (
          <ul className={styles.meta}>
            {source && <li>Source: {source}</li>}
            {row.validTo && <li>Valid until {formatSaasDate(row.validTo)}</li>}
            {row.termEnd && (
              <li>
                {row.autoRenew ? 'Renews on' : 'Term ends on'} {formatSaasDate(row.termEnd)}
              </li>
            )}
          </ul>
        )}

        {message && <p className={`${styles.message} ${TONE_CLASSES[tone]}`}>{message}</p>}

        {row.plans.length > 0 && (
          <p className={styles.plans}>
            Plans: {row.plans.map(Subscription._describePlan).join(' · ')}
          </p>
        )}

        {canEnable && (
          <button
            type="button"
            className={styles.primaryButton}
            disabled={!!enablingKey || !!activatingId}
            data-service-key={row.serviceKey}
            onClick={this._onEnableClick}
          >
            {enablingKey === row.serviceKey ? 'Opening Microsoft Marketplace…' : `Enable ${row.name}`}
          </button>
        )}
      </li>
    );
  }

  /** The tenant's Marketplace subscriptions; nothing at all while there are none to show. */
  private _renderSubscriptions(entitlements: ISaasEntitlements): React.ReactElement | undefined {
    const { subscriptions, subscriptionsError } = this.state;

    if (subscriptions.length === 0 && !subscriptionsError) {
      return undefined;
    }

    return (
      <div className={styles.subscriptions}>
        <h3 className={styles.sectionTitle}>Marketplace subscriptions</h3>

        {subscriptionsError && (
          <div className={styles.notice} role="status">
            Your Marketplace subscriptions could not be loaded. {subscriptionsError.message}
            {subscriptionsError.correlationId && <> Reference: {subscriptionsError.correlationId}</>}
          </div>
        )}

        {subscriptions.length > 0 && (
          <ul className={styles.list}>
            {subscriptions.map((subscription: ISaasSubscription) => this._renderSubscription(subscription, entitlements))}
          </ul>
        )}
      </div>
    );
  }

  private _renderSubscription(subscription: ISaasSubscription, entitlements: ISaasEntitlements): React.ReactElement {
    const { activatingId, enablingKey, isRefreshing } = this.state;
    const id: string = subscription.marketplaceSubscriptionId;
    const tone: SaasTone = getSubscriptionStatusTone(subscription.status);
    const canActivate: boolean = canActivateSubscription(subscription, entitlements);
    const isAwaitingActivation: boolean = subscription.status === 'PendingFulfillmentStart';

    return (
      <li key={id} className={styles.card} data-subscription-id={id}>
        <div className={styles.cardHeader}>
          <h4 className={styles.serviceName}>{subscription.subscriptionName || subscription.planName || id}</h4>
          <span className={`${styles.badge} ${TONE_CLASSES[tone]}`} data-status={subscription.status}>
            {getSubscriptionStatusLabel(subscription.status)}
          </span>
          {subscription.isFreeTrial && <span className={styles.trial}>Free trial</span>}
        </div>

        <ul className={styles.meta}>
          <li>Plan: {subscription.planName || subscription.planId}</li>
          {subscription.services.length > 0 && (
            <li>Services: {subscription.services.map((key: string) => this._serviceName(key)).join(', ')}</li>
          )}
          {subscription.termEnd && (
            <li>
              {subscription.autoRenew ? 'Renews on' : 'Term ends on'} {formatSaasDate(subscription.termEnd)}
            </li>
          )}
        </ul>

        {isAwaitingActivation && !entitlements.isTenantAdmin && (
          <p className={`${styles.message} ${TONE_CLASSES.warning}`}>
            Bought but not activated yet. A Global Administrator or SharePoint Administrator can
            activate it on this page.
          </p>
        )}

        {canActivate && (
          <button
            type="button"
            className={styles.primaryButton}
            disabled={!!activatingId || !!enablingKey || isRefreshing}
            data-subscription-id={id}
            onClick={this._onActivateClick}
          >
            {activatingId === id ? 'Activating…' : 'Activate'}
          </button>
        )}
      </li>
    );
  }

  private _renderBanner(banner: ActivationBanner): React.ReactElement {
    switch (banner.kind) {
      case 'activating':
        return (
          <div
            className={`${styles.banner} ${styles.bannerBusy} ${styles.toneNeutral}`}
            role="status"
            aria-live="polite"
            data-banner="activating"
          >
            <span className={styles.spinner} aria-hidden="true" />
            <span>Activating your subscription&hellip;</span>
          </div>
        );
      case 'done':
        return this._renderActivationResult(banner.result);
      case 'failed':
        return (
          <div className={`${styles.banner} ${styles.toneError}`} role="alert" data-banner="failed">
            <p className={styles.bannerTitle}>Your subscription could not be activated.</p>
            {this._renderErrorInfo(banner.error)}
            {banner.canRetry && (
              <button
                type="button"
                className={styles.primaryButton}
                disabled={this.state.isLoading || !!this.state.activatingId}
                onClick={this._onRetryActivation}
              >
                Try again
              </button>
            )}
          </div>
        );
      default:
        return (
          <div className={`${styles.banner} ${styles.toneNeutral}`} role="status" data-banner="returned">
            Welcome back. No completed purchase was reported, so nothing was activated. A service you
            did buy is listed under Marketplace subscriptions, where an administrator can activate it.
          </div>
        );
    }
  }

  private _renderActivationResult(result: ISaasActivationResult): React.ReactElement {
    if (isActivationSuccessful(result)) {
      const names: string[] = result.services.map((key: string) => this._serviceName(key, result.entitlements));
      const text: string =
        names.length === 0
          ? 'Your subscription is active.'
          : `${names.join(', ')} ${names.length === 1 ? 'is' : 'are'} active.`;

      return (
        <div className={`${styles.banner} ${styles.toneSuccess}`} role="status" data-banner="activated">
          {text} Everyone in your organisation can use {names.length === 1 ? 'it' : 'them'} now.
        </div>
      );
    }

    return (
      <div className={`${styles.banner} ${styles.toneWarning}`} role="status" data-banner="notActivated">
        <p className={styles.bannerTitle}>The subscription was not activated.</p>
        <p>{result.message || `Microsoft reports this subscription as ${getSubscriptionStatusLabel(result.state).toLowerCase()}.`}</p>
        {result.correlationId && <p className={styles.reference}>Reference: {result.correlationId}</p>}
      </div>
    );
  }

  private _renderErrorInfo(info: ISaasErrorInfo): React.ReactElement {
    return (
      <>
        <p className={styles.errorMessage}>{info.message}</p>
        {info.correlationId && <p className={styles.reference}>Reference: {info.correlationId}</p>}
      </>
    );
  }

  /**
   * Loads the catalog, the subscriptions and - unless an activation has just returned them -
   * the entitlements (always fresh), in parallel. Only the entitlements are essential:
   * without the catalog the rows come from the server's per-service details, minus the
   * plans, and without the subscriptions only their section is missing.
   */
  private _load(knownEntitlements?: ISaasEntitlements): void {
    const { saasService } = this.props;
    let catalogError: ISaasErrorInfo | undefined;

    this.setState({ isLoading: true, loadError: undefined });
    this._loadSubscriptions();

    const catalogRequest: Promise<void> = saasService.getCatalog().then(
      (catalog: ISaasCatalog): void => {
        this._catalog = catalog;
      },
      (error: unknown): void => {
        catalogError = getSaasErrorInfo(error);
      }
    );
    const entitlementsRequest: Promise<ISaasEntitlements> = knownEntitlements
      ? Promise.resolve(knownEntitlements)
      : saasService.getEntitlements({ refresh: true });

    Promise.all([entitlementsRequest, catalogRequest]).then(
      (results: [ISaasEntitlements, void]): void => {
        if (this._isActive) {
          this.setState({ isLoading: false, catalogError });
          this._showEntitlements(results[0]);
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ isLoading: false, loadError: getSaasErrorInfo(error) });
        }
      }
    );
  }

  private _loadSubscriptions(): void {
    this.props.saasService.getSubscriptions().then(
      (subscriptions: ISaasSubscription[]): void => {
        if (this._isActive) {
          this.setState({ subscriptions, subscriptionsError: undefined });
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ subscriptionsError: getSaasErrorInfo(error) });
        }
      }
    );
  }

  private _showEntitlements(entitlements: ISaasEntitlements): void {
    this.setState({ entitlements, rows: mergeServiceRows(this._catalog, entitlements) });
  }

  /**
   * Activates a Marketplace subscription and shows the outcome.
   *
   * On a return the page has not loaded yet, so it loads afterwards - from the entitlements
   * the activation answered, or afresh when it failed. Otherwise the answer's entitlements
   * replace the rows on screen and the subscriptions are reloaded.
   */
  private _activate(subscriptionId: string): void {
    this.setState({ activatingId: subscriptionId, banner: { kind: 'activating' }, actionError: undefined });

    this.props.saasService.activateSubscription(subscriptionId).then(
      (result: ISaasActivationResult): void => {
        if (!this._isActive) {
          return;
        }

        this.setState({ activatingId: undefined, banner: { kind: 'done', result } });
        if (this.state.entitlements) {
          this._showEntitlements(result.entitlements);
          this._loadSubscriptions();
        } else {
          this._load(result.entitlements);
        }
      },
      (error: unknown): void => {
        if (!this._isActive) {
          return;
        }

        this.setState({
          activatingId: undefined,
          banner: { kind: 'failed', subscriptionId, error: getSaasErrorInfo(error), canRetry: isRetryableSaasError(error) }
        });
        if (!this.state.entitlements) {
          this._load();
        }
      }
    );
  }

  private _clearQuery = (): void => {
    if (this._isActive) {
      this.props.onClearQuery();
    }
  };

  private _onRefresh = (): void => {
    this.setState({ isRefreshing: true, actionError: undefined });
    this._loadSubscriptions();

    this.props.saasService.getEntitlements({ refresh: true }).then(
      (entitlements: ISaasEntitlements): void => {
        if (this._isActive) {
          this.setState({ isRefreshing: false });
          this._showEntitlements(entitlements);
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ isRefreshing: false, actionError: getSaasErrorInfo(error) });
        }
      }
    );
  };

  private _onRetryLoad = (): void => {
    this._load();
  };

  private _onRetryActivation = (): void => {
    const banner: ActivationBanner | undefined = this.state.banner;
    if (banner && banner.kind === 'failed') {
      this._activate(banner.subscriptionId);
    }
  };

  private _onActivateClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const subscriptionId: string | undefined = event.currentTarget.dataset.subscriptionId;
    if (subscriptionId) {
      this._activate(subscriptionId);
    }
  };

  private _onEnableClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const serviceKey: string | undefined = event.currentTarget.dataset.serviceKey;
    if (serviceKey) {
      this._startPurchase(serviceKey);
    }
  };

  /**
   * Records the purchase intent, then leaves for the Marketplace offer in this tab. The
   * button stays in its "opening" state on success - the page is about to be replaced.
   */
  private _startPurchase(serviceKey: string): void {
    const getCurrentUrl: () => string = this.props.getCurrentUrl || ((): string => window.location.href);
    const navigateToUrl: (url: string) => void =
      this.props.navigateToUrl || ((url: string): void => window.location.assign(url));
    const returnUrl: string = buildPurchaseReturnUrl(getCurrentUrl(), serviceKey);

    this.setState({ enablingKey: serviceKey, actionError: undefined });

    this.props.saasService.createPurchaseIntent(serviceKey, returnUrl).then(
      (intent: ISaasPurchaseIntent): void => {
        if (!this._isActive) {
          return;
        }

        if (!isSafeOfferUrl(intent.marketplaceOfferUrl)) {
          this.setState({
            enablingKey: undefined,
            actionError: {
              message: 'The Microsoft Marketplace address for this service is not valid. Please contact XSProject support.'
            }
          });
          return;
        }

        navigateToUrl(intent.marketplaceOfferUrl);
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ enablingKey: undefined, actionError: getSaasErrorInfo(error) });
        }
      }
    );
  }

  /**
   * A service's display name - from `entitlements` when given (an activation's answer, which
   * may arrive before the rows), then from the rows - falling back to its key.
   */
  private _serviceName(serviceKey: string, entitlements?: ISaasEntitlements): string {
    const detail: ISaasServiceStatus | undefined = findServiceStatus(entitlements, serviceKey);
    if (detail) {
      return detail.name;
    }

    const row: ISaasServiceRow | undefined = this.state.rows.filter(
      (candidate: ISaasServiceRow) => candidate.serviceKey.toUpperCase() === serviceKey.toUpperCase()
    )[0];

    return row ? row.name : serviceKey;
  }

  private static _describePlan(plan: ISaasCatalogPlan): string {
    const details: string[] = [plan.billingTerm, plan.pricingModel].filter(
      (value: string | null): value is string => !!value
    );

    return details.length > 0 ? `${plan.name} (${details.join(', ')})` : plan.name;
  }
}
