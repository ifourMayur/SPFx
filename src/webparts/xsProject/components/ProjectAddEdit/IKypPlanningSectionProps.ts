import { ILookupOption } from '../../../../models/Building';
import { IKypForm } from '../../../../models/ProjectIntegration';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

export interface IKypPlanningSectionProps {
  /** The project being edited; `0` while adding. */
  projectId: number;
  integrationService: IProjectIntegrationService;
  /** The form's value; `undefined` until this section has loaded the project's settings. */
  value: IKypForm | undefined;
  planningProjectError?: string;
  isDisabled: boolean;
  /** The stored value, once loaded - also what the form's Reset returns to. */
  onLoaded: (value: IKypForm) => void;
  onChange: (value: IKypForm) => void;
}

export interface IKypPlanningSectionState {
  /** `notConfigured`: neither the account nor the project has a KYP token. */
  status: 'loading' | 'ready' | 'notConfigured' | 'failed';
  loadError?: string;
  projects: ILookupOption[];
  isLoadingProjects: boolean;
  projectsError?: string;
}
