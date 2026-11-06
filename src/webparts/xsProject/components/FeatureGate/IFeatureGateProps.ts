import * as React from 'react';

import { ISaasService } from '../../../../services/SaasService';

export interface IFeatureGateProps {
  /** Where the tenant's entitlements come from. Cached and shared, so many gates cost one call. */
  saasService: ISaasService;
  /** The feature this gate guards, e.g. `TENDER.CORE` (see `SAAS_FEATURE_KEYS`). */
  featureKey: string;
  /** The service that provides the feature, e.g. `TENDER` - what an admin would enable. */
  serviceKey: string;
  /** Name shown before the entitlements arrive, or when the server does not name the service. */
  serviceName: string;
  /**
   * Opens the Subscription page on `serviceKey`. Offered to admins only; without it the
   * locked panel explains but offers no button.
   */
  onEnableService?: (serviceKey: string) => void;
  /**
   * Render nothing at all - no "checking" state, no locked panel - unless the feature is
   * enabled. For a section that should simply not exist for a tenant without the service
   * (its heading goes inside the gate). A failed check still shows its error and Retry, so
   * a network blip is not mistaken for a missing subscription.
   */
  hideWhenLocked?: boolean;
  /** The paid feature, rendered only while the server reports it enabled. */
  children?: React.ReactNode;
}
