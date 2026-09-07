import * as React from 'react';
import { Navigate, useNavigate, useParams } from 'react-router-dom';

import ProjectAddEdit from './ProjectAddEdit';
import { IBuildingForm, IProjectFormFeatures } from '../../../../models/Building';
import { ROUTE_PATHS } from '../../../../models/Navigation';
import { IProjectListRow } from '../../../../models/ProjectListRow';
import { ILookupService } from '../../../../services/LookupService';

export interface IProjectAddEditRouteProps {
  /** Reference data for the form's dropdowns. */
  lookupService: ILookupService;
  /** `add` serves `/projects/add`; `edit` serves `/projects/edit/:id`. */
  mode: 'add' | 'edit';
  /** Listing rows, used to resolve the name of the project being edited. */
  rows: IProjectListRow[];
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
  const { mode, rows, savedForms, onSave } = props;
  const parameters = useParams<{ id: string }>();
  const navigate = useNavigate();

  const projectId: number | undefined = mode === 'edit' ? Number(parameters.id) : undefined;

  // Both callbacks are declared before the guard below: hooks must run in the same order
  // on every render, so none may sit after a conditional return.
  const goToListing = React.useCallback((): void => {
    navigate(ROUTE_PATHS.projectList);
  }, [navigate]);

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

  const row: IProjectListRow | undefined = projectId === undefined
    ? undefined
    : rows.filter((candidate: IProjectListRow) => candidate.id === projectId)[0];

  return (
    <ProjectAddEdit
      lookupService={props.lookupService}
      // Remounts on a switch between add and edit, or between two projects, so the form
      // rebuilds its initial state instead of showing the previous project's values.
      key={`${mode}-${projectId || 0}`}
      projectId={projectId}
      initialForm={projectId === undefined ? undefined : savedForms[projectId]}
      fallbackProjectName={row?.name}
      features={props.features}
      clientId={props.clientId}
      clientName={props.clientName}
      canSave={props.canSave}
      onSave={saveThenReturn}
      onCancel={goToListing}
    />
  );
}
