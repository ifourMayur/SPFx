import * as React from 'react';

import styles from './FeatureGate.module.scss';
import type { IFeatureGateProps } from './IFeatureGateProps';
import type { IFeatureGateState } from './IFeatureGateState';
import {
  ISaasEntitlements,
  ISaasServiceStatus,
  findServiceStatus,
  getSaasErrorInfo,
  getServiceStatusMessage,
  isFeatureEnabled
} from '../../../../models/Saas';

/**
 * Shows a paid feature only while the tenant is entitled to it.
 *
 * Wrap any screen, section or button that belongs to a Marketplace service:
 *
 * ```tsx
 * <FeatureGate saasService={...} featureKey="TENDER.CORE" serviceKey="TENDER"
 *              serviceName="Tender" onEnableService={openSubscription}>
 *   ...the Tender section...
 * </FeatureGate>
 * ```
 *
 * The children render only when `GET saas/entitlements` lists `featureKey`. Otherwise a
 * locked panel says why: an admin gets "Enable {service}" (to the Subscription page),
 * anyone else "Contact your administrator" - or, with `hideWhenLocked`, nothing at all.
 * **This is UX, not security** - the Web API
 * refuses the same feature on its own gated endpoints whatever this shows.
 *
 * It fails **closed**: while loading, and when the check itself fails, the feature stays
 * hidden - but the failure is said out loud, with a Retry, so a network blip does not
 * look like a lost subscription.
 *
 * Every gate asks the shared `ISaasService`, whose cache and in-flight sharing make a
 * screen with several gates cost one request.
 */
export default class FeatureGate extends React.Component<IFeatureGateProps, IFeatureGateState> {
  /** Guards against setting state after unmount. */
  private _isActive: boolean = false;
  /** Ignores an answer that a newer request (a retry) has superseded. */
  private _requestId: number = 0;

  public constructor(props: IFeatureGateProps) {
    super(props);

    this.state = { phase: 'loading' };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._load(false);
  }

  public componentDidUpdate(previous: IFeatureGateProps): void {
    if (previous.saasService !== this.props.saasService) {
      this._load(false);
    }
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactNode {
    const { featureKey, children, hideWhenLocked } = this.props;
    const { phase, entitlements } = this.state;

    if (phase === 'loading') {
      if (hideWhenLocked) {
        return null;
      }
      return (
        <div className={styles.checking} role="status" aria-live="polite" data-feature-gate={featureKey}>
          <span className={styles.spinner} aria-hidden="true" />
          <span>Checking your subscription&hellip;</span>
        </div>
      );
    }

    if (phase === 'error') {
      return this._renderError();
    }

    if (isFeatureEnabled(entitlements, featureKey)) {
      return <>{children}</>;
    }

    return hideWhenLocked ? null : this._renderLocked(entitlements);
  }

  private _renderError(): React.ReactElement {
    const { error } = this.state;

    return (
      <div className={styles.locked} role="alert" data-feature-gate={this.props.featureKey} data-state="error">
        <p className={styles.lockedTitle}>
          We couldn&apos;t check your subscription, so {this.props.serviceName} is hidden for now.
        </p>
        {error && <p className={styles.detail}>{error.message}</p>}
        {error?.correlationId && <p className={styles.reference}>Reference: {error.correlationId}</p>}
        <button type="button" className={styles.button} onClick={this._onRetry}>
          Retry
        </button>
      </div>
    );
  }

  private _renderLocked(entitlements: ISaasEntitlements | undefined): React.ReactElement {
    const { featureKey, serviceKey, onEnableService } = this.props;
    const detail: ISaasServiceStatus | undefined = findServiceStatus(entitlements, serviceKey);
    const name: string = detail?.name || this.props.serviceName;
    const message: string | undefined = detail ? getServiceStatusMessage(detail, new Date()) : undefined;
    const isAdmin: boolean = !!entitlements && entitlements.isTenantAdmin;
    // Nobody can buy a service that is switched off for everyone.
    const canOffer: boolean = isAdmin && !!onEnableService && detail?.status !== 'Unavailable';

    return (
      <div className={styles.locked} data-feature-gate={featureKey} data-state="locked">
        <p className={styles.lockedTitle}>
          <span className={styles.lockIcon} aria-hidden="true">&#128274;</span>
          {name} is not enabled for your organisation.
        </p>
        {message && <p className={styles.detail}>{message}</p>}
        {canOffer ? (
          <button type="button" className={styles.button} onClick={this._onEnable}>
            Enable {name}
          </button>
        ) : (
          !isAdmin && (
            <p className={styles.detail}>
              Contact your administrator to enable {name}. Only a Global Administrator or SharePoint
              Administrator can enable services.
            </p>
          )
        )}
      </div>
    );
  }

  private _load(refresh: boolean): void {
    const requestId: number = ++this._requestId;

    this.setState({ phase: 'loading', error: undefined });

    this.props.saasService.getEntitlements({ refresh }).then(
      (entitlements: ISaasEntitlements): void => {
        if (this._isActive && requestId === this._requestId) {
          this.setState({ phase: 'ready', entitlements });
        }
      },
      (error: unknown): void => {
        // Already logged with its code and reference by the service.
        if (this._isActive && requestId === this._requestId) {
          this.setState({ phase: 'error', entitlements: undefined, error: getSaasErrorInfo(error) });
        }
      }
    );
  }

  // Assigned as properties so `this` is bound without a per-render closure.
  private _onRetry = (): void => {
    this._load(true);
  };

  private _onEnable = (): void => {
    if (this.props.onEnableService) {
      this.props.onEnableService(this.props.serviceKey);
    }
  };
}
