import * as React from 'react';
import { Navigate, Route, Routes, useLocation, useNavigate } from 'react-router-dom';

import styles from './XsProject.module.scss';
import Menu from '../Menu/Menu';
import Project from '../Project/Project';
import ProjectList from '../ProjectList/ProjectList';
import ProjectAddEditRoute from '../ProjectAddEdit/ProjectAddEditRoute';
import Document from '../Document/Document';
import Search from '../Search/Search';
import Suppliers from '../Suppliers/Suppliers';
import { IBuildingForm, IProjectFormFeatures } from '../../../../models/Building';
import { DEFAULT_ROUTE_PATH, ROUTE_PATHS, projectEditPath } from '../../../../models/Navigation';
import { IProjectListRow } from '../../../../models/ProjectListRow';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectService } from '../../../../services/ProjectService';

export interface IAppShellProps {
  projectService: IProjectService;
  lookupService: ILookupService;
  rows: IProjectListRow[];
  savedForms: { [projectId: number]: IBuildingForm };
  features: IProjectFormFeatures;
  clientId: string;
  clientName: string;
  /** Whether this user may add or edit projects, from `canAddEditProject`. */
  canAddEdit: boolean;
  /** Whether this user may open the project form, from `canOpenProjectForm`. */
  canOpenForm: boolean;
  /** Save confirmation, shown by `ProjectList` as a toast. */
  notification?: string;
  /** Title of the SharePoint site chosen in `SharePointSites`, named in the site bar. */
  siteTitle: string;
  /** Returns the user to the site picker. */
  onChangeSite: () => void;
  onSaveProject: (form: IBuildingForm) => Promise<void>;
  onToggleFavorite: (projectId: number) => void;
  onDismissNotification: () => void;
}

/**
 * Routed shell: the shared `Menu` beside whichever page the current route selects.
 *
 * This is the only component in the web part written as a function, and it exists to be
 * exactly that - React Router 6 exposes the location and the navigate function through
 * hooks alone, so one function component keeps every other component a class, as the
 * convention in `src/index.ts` documents.
 *
 * `Menu` is rendered once here, outside `<Routes>`, which is what makes it genuinely
 * shared: no page carries navigation markup, and adding a page cannot forget the menu.
 */
export default function AppShell(props: IAppShellProps): React.ReactElement {
  const location = useLocation();
  const navigate = useNavigate();

  // Wrapped rather than passed straight to `Menu`: `navigate` is overloaded (a path or a
  // history delta), and the menu's contract is a plain path.
  const goTo = React.useCallback(
    (path: string): void => {
      navigate(path);
    },
    [navigate]
  );

  const goToAdd = React.useCallback((): void => {
    navigate(ROUTE_PATHS.projectAdd);
  }, [navigate]);

  const goToEdit = React.useCallback(
    (projectId: number): void => {
      navigate(projectEditPath(projectId));
    },
    [navigate]
  );

  const listing = (
    <ProjectList
      rows={props.rows}
      canAddEdit={props.canAddEdit}
      notification={props.notification}
      onAddProject={goToAdd}
      onEditProject={goToEdit}
      onToggleFavorite={props.onToggleFavorite}
      onDismissNotification={props.onDismissNotification}
    />
  );

  const form = (mode: 'add' | 'edit'): React.ReactElement => (
    <ProjectAddEditRoute
      lookupService={props.lookupService}
      mode={mode}
      rows={props.rows}
      savedForms={props.savedForms}
      features={props.features}
      clientId={props.clientId}
      clientName={props.clientName}
      canSave={props.canOpenForm}
      onSave={props.onSaveProject}
    />
  );

  return (
    <div className={styles.appLayout}>
      <Menu activePath={location.pathname} onNavigate={goTo} />

      <div className={styles.content}>
        {/* Which site the session is working in is not obvious from any page, and it is
            chosen before the app renders - so the shell says so, and is where the choice
            is revisited. */}
        <div className={styles.siteBar}>
          <span className={styles.siteName}>
            Site: <strong>{props.siteTitle}</strong>
          </span>
          <button type="button" className={styles.changeSite} onClick={props.onChangeSite}>
            Change site
          </button>
        </div>

        <Routes>
          <Route path={ROUTE_PATHS.projectList} element={listing} />
          <Route path={ROUTE_PATHS.projectAdd} element={form('add')} />
          <Route path={ROUTE_PATHS.projectEdit} element={form('edit')} />
          <Route path={ROUTE_PATHS.projectDashboard} element={<Project projectService={props.projectService} />} />
          <Route path={ROUTE_PATHS.projectDocument} element={<Document />} />
          <Route path={ROUTE_PATHS.search} element={<Search />} />
          <Route path={ROUTE_PATHS.suppliers} element={<Suppliers />} />
          {/* `/` and anything unrecognised land on the view the web part opened on
              before it had routes. `replace` keeps the bad path out of history. */}
          <Route path="*" element={<Navigate to={DEFAULT_ROUTE_PATH} replace />} />
        </Routes>
      </div>
    </div>
  );
}
