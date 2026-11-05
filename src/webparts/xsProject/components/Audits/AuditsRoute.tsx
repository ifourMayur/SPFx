import * as React from 'react';
import { Navigate, useNavigate, useParams } from 'react-router-dom';

import Audits, { IAuditsProps } from './Audits';
import { ROUTE_PATHS, projectEditPath } from '../../../../models/Navigation';

export type IAuditsRouteProps = Omit<IAuditsProps, 'projectId' | 'fallbackProjectName' | 'onBack'> & {
  /** Names of the projects seen on the listing, keyed by id. */
  projectNames: { [projectId: number]: string };
};

/**
 * Route adapter for `Audits` (`/projects/audits/:id`): reads the id, and goes back to the
 * project's form. Same bridge as `ProjectAddEditRoute` - the page itself knows no routing.
 */
export default function AuditsRoute(props: IAuditsRouteProps): React.ReactElement {
  const { projectNames, ...rest } = props;
  const parameters = useParams<{ id: string }>();
  const navigate = useNavigate();
  const projectId: number = Number(parameters.id);

  const goBack = React.useCallback((): void => {
    navigate(projectEditPath(projectId));
  }, [navigate, projectId]);

  if (!isFinite(projectId) || projectId <= 0) {
    return <Navigate to={ROUTE_PATHS.projectList} replace />;
  }

  return (
    <Audits
      {...rest}
      key={projectId}
      projectId={projectId}
      fallbackProjectName={projectNames[projectId]}
      onBack={goBack}
    />
  );
}
