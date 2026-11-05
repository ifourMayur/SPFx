import * as React from 'react';
import { Navigate, useNavigate, useParams } from 'react-router-dom';

import ProjectAddEdit from './ProjectAddEdit';
import { IBuildingForm, IProjectFormFeatures } from '../../../../models/Building';
import { ROUTE_PATHS, projectAuditsPath } from '../../../../models/Navigation';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';
import { ISaasService } from '../../../../services/SaasService';

export interface IProjectAddEditRouteProps {
  /** Reference data for the form's dropdowns. */
  lookupService: ILookupService;
  /** The ED Controls and KYP sections' lookups. */
  integrationService: IProjectIntegrationService;
  /** Offers "Manage audits" on a saved, ED Controls-linked project (`canManageAudits`). */
  canManageAudits: boolean;
  /** Entitlements behind the form's gated Tender and Minutes of Meeting sections. */
  saasService: ISaasService;
  /** Opens the Subscription page on a service - offered to admins by a locked section. */
  onOpenSubscription?: (serviceKey: string) => void;
  /** `add` serves `/projects/add`; `edit` serves `/projects/edit/:id`. */
  mode: 'add' | 'edit';
  /**
   * Names of the projects seen on the listing, keyed by id, used to title the form for the
   * project being edited.
   *
   * A name map rather than the listing's rows: the listing is server-paged and keeps no
   * rows that outlive it, and the name is the only thing this route ever needed from them.
   */
  projectNames: { [projectId: number]: string };
  /** Forms saved earlier in this session, keyed by project id. */
  savedForms: { [projectId: number]: IBuildingForm };
  features: IProjectFormFeatures;
  clientId: string;
  clientName: string;
  canSave: boolean;
  /** Persists the form. Navigation back to the listing happens here, once it resolves. */
  onSave: (form: IBuildingForm) => Promise<void>;
}

/**
 * Route adapter for `ProjectAddEdit`.
 *
 * React Router 6 exposes the current route only through hooks, and every component in this
 * web part is a class (the convention documented in `src/index.ts`). This is the one small
 * function component that bridges the two: it reads the `:id` parameter and the navigate
 * function, and hands them to the class as plain props.
 *
 * Keeping the bridge here means `ProjectAddEdit` itself knows nothing about routing and
 * could be rendered from a dialog or another shell unchanged.
 */
export default function ProjectAddEditRoute(props: IProjectAddEditRouteProps): React.ReactElement {
  const { mode, projectNames, savedForms, onSave } = props;
  const parameters = useParams<{ id: string }>();
  const navigate = useNavigate();

  const projectId: number | undefined = mode === 'edit' ? Number(parameters.id) : undefined;

  // Both callbacks are declared before the guard below: hooks must run in the same order
  // on every render, so none may sit after a conditional return.
  const goToListing = React.useCallback((): void => {
    navigate(ROUTE_PATHS.projectList);
  }, [navigate]);

  const goToAudits = React.useCallback((): void => {
    if (projectId !== undefined) {
      navigate(projectAuditsPath(projectId));
    }
  }, [navigate, projectId]);

  const saveThenReturn = React.useCallback(
    async (form: IBuildingForm): Promise<void> => {
      await onSave(form);
      // Only reached when the save succeeded; a rejection is reported on the form instead.
      navigate(ROUTE_PATHS.projectList);
    },
    [navigate, onSave]
  );

  // A malformed id (`/projects/edit/abc`) is not an editable project. Redirecting beats
  // silently opening a blank add form under an edit URL.
  if (projectId !== undefined && (!isFinite(projectId) || projectId <= 0)) {
    return <Navigate to={ROUTE_PATHS.projectList} replace />;
  }

  return (
    <ProjectAddEdit
      lookupService={props.lookupService}
      integrationService={props.integrationService}
      onOpenAudits={props.canManageAudits && projectId !== undefined ? goToAudits : undefined}
      saasService={props.saasService}
      onOpenSubscription={props.onOpenSubscription}
      // Remounts on a switch between add and edit, or between two projects, so the form
      // rebuilds its initial state instead of showing the previous project's values.
      key={`${mode}-${projectId || 0}`}
      projectId={projectId}
      initialForm={projectId === undefined ? undefined : savedForms[projectId]}
      fallbackProjectName={projectId === undefined ? undefined : projectNames[projectId]}
      features={props.features}
      clientId={props.clientId}
      clientName={props.clientName}
      canSave={props.canSave}
      onSave={saveThenReturn}
      onCancel={goToListing}
    />
  );
}
