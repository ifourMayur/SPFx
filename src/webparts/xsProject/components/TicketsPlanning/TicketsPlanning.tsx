import * as React from 'react';

import styles from './TicketsPlanning.module.scss';
import PlanningList from './PlanningList';
import TicketList from './TicketList';
import FeatureGate from '../FeatureGate/FeatureGate';
import { SAAS_FEATURE_KEYS, SAAS_SERVICE_KEYS } from '../../../../models/Saas';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';
import { ISaasService } from '../../../../services/SaasService';

export interface ITicketsPlanningProps {
  integrationService: IProjectIntegrationService;
  saasService: ISaasService;
  onOpenSubscription?: (serviceKey: string) => void;
}

/**
 * The user's ED Controls tickets and KYP planning across their projects - the Tickets and
 * Planning tabs of the reference's "My tasks" page (`Views/Task/Index.cshtml`). Each section is
 * a paid service, so it renders through its own `FeatureGate`, and loads only once unlocked.
 */
export default class TicketsPlanning extends React.Component<ITicketsPlanningProps> {
  public render(): React.ReactElement<ITicketsPlanningProps> {
    const { integrationService, saasService, onOpenSubscription } = this.props;

    return (
      <section className={styles.ticketsPlanning}>
        <h2 className={styles.title}>Tickets &amp; Planning</h2>

        <div className={styles.section}>
          <div className={styles.sectionLabel}>Tickets</div>
          <hr className={styles.sectionRule} />
          <FeatureGate
            saasService={saasService}
            featureKey={SAAS_FEATURE_KEYS.edControl}
            serviceKey={SAAS_SERVICE_KEYS.edControl}
            serviceName="ED Control"
            onEnableService={onOpenSubscription}
          >
            <TicketList integrationService={integrationService} />
          </FeatureGate>
        </div>

        <div className={styles.section}>
          <div className={styles.sectionLabel}>Planning</div>
          <hr className={styles.sectionRule} />
          <FeatureGate
            saasService={saasService}
            featureKey={SAAS_FEATURE_KEYS.kyp}
            serviceKey={SAAS_SERVICE_KEYS.kyp}
            serviceName="Know Your Property"
            onEnableService={onOpenSubscription}
          >
            <PlanningList integrationService={integrationService} />
          </FeatureGate>
        </div>
      </section>
    );
  }
}
