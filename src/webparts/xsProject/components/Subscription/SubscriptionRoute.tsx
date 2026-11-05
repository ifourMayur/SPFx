import * as React from 'react';
import { useLocation, useNavigate } from 'react-router-dom';

import Subscription from './Subscription';
import { ROUTE_PATHS } from '../../../../models/Navigation';
import { ISaasService } from '../../../../services/SaasService';

export interface ISubscriptionRouteProps {
  saasService: ISaasService;
}

/**
 * Route adapter for `Subscription`, in the same shape as `ProjectAddEditRoute`.
 *
 * React Router 6 exposes the location and navigation through hooks only, so this small
 * function component hands the class what it needs as plain props: the route's query
 * (`?saas=returned&service=TENDER&subscriptionId={guid}` - the Marketplace return, or
 * `?service=TENDER` - the service a locked `FeatureGate` asked for) and a way to drop that
 * query once it has been read. Replacing the history entry, rather than pushing one, keeps
 * Back from replaying the return and activating again.
 */
export default function SubscriptionRoute(props: ISubscriptionRouteProps): React.ReactElement {
  const location = useLocation();
  const navigate = useNavigate();

  const clearQuery = React.useCallback((): void => {
    navigate(ROUTE_PATHS.subscription, { replace: true });
  }, [navigate]);

  return <Subscription saasService={props.saasService} query={location.search} onClearQuery={clearQuery} />;
}
