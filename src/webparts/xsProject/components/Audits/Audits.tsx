import * as React from 'react';

import styles from './Audits.module.scss';
import AuditManager from './AuditManager';
import FeatureGate from '../FeatureGate/FeatureGate';
import { SAAS_FEATURE_KEYS, SAAS_SERVICE_KEYS } from '../../../../models/Saas';
import { IEdControlAuditService } from '../../../../services/EdControlAuditService';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';
import { ISaasService } from '../../../../services/SaasService';

export interface IAuditsProps {
  projectId: number;
  /** The name from the listing, shown until the project's own is read. */
  fallbackProjectName?: string;
  integrationService: IProjectIntegrationService;
  auditService: IEdControlAuditService;
  lookupService: ILookupService;
  saasService: ISaasService;
  onOpenSubscription?: (serviceKey: string) => void;
  isGrandChildFolderEnabled: boolean;
  /** From `canManageAudits`: Admin, Manager and Assistant Manager. */
  canManage: boolean;
  onBack: () => void;
}

interface IAuditsState {
  projectName?: string;
}

/**
 * The Audits page of one project - the reference's `Building/Audits`. ED Control is a paid
 * service, so everything below the title renders through a `FeatureGate` (`EDCONTROL.CORE`);
 * the Web API enforces the same on `api/EDControlAudit` and `api/EDControls`.
 */
export default class Audits extends React.Component<IAuditsProps, IAuditsState> {
  public constructor(props: IAuditsProps) {
    super(props);

    this.state = {};
  }

  public render(): React.ReactElement<IAuditsProps> {
    const { canManage, onBack } = this.props;
    const projectName: string | undefined = this.state.projectName || this.props.fallbackProjectName;

    return (
      <section className={styles.audits}>
        <div className={styles.header}>
          <div>
            <h2 className={styles.title}>Audits</h2>
            {projectName && <div className={styles.projectName}>{projectName}</div>}
          </div>
          <button type="button" className={styles.secondaryButton} onClick={onBack}>Back</button>
        </div>

        {!canManage ? (
          <div className={styles.accessDenied} role="alert">
            Your account does not have permission to manage audits.
          </div>
        ) : (
          <FeatureGate
            saasService={this.props.saasService}
            featureKey={SAAS_FEATURE_KEYS.edControl}
            serviceKey={SAAS_SERVICE_KEYS.edControl}
            serviceName="ED Control"
            onEnableService={this.props.onOpenSubscription}
          >
            <AuditManager
              projectId={this.props.projectId}
              integrationService={this.props.integrationService}
              auditService={this.props.auditService}
              lookupService={this.props.lookupService}
              isGrandChildFolderEnabled={this.props.isGrandChildFolderEnabled}
              onProjectName={this._onProjectName}
            />
          </FeatureGate>
        )}
      </section>
    );
  }

  private _onProjectName = (projectName: string): void => {
    this.setState({ projectName });
  };
}
