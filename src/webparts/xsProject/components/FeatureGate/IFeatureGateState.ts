import { ISaasEntitlements, ISaasErrorInfo } from '../../../../models/Saas';

export interface IFeatureGateState {
  /** `loading` until the first answer, then `ready` or `error`. A retry goes back to `loading`. */
  phase: 'loading' | 'ready' | 'error';
  entitlements?: ISaasEntitlements;
  /** Why the check failed, when `phase` is `error`. */
  error?: ISaasErrorInfo;
}
