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
import Subscription from '../Subscription/Subscription';
import { IBuildingForm, IProjectFormFeatures } from '../../../../models/Building';
import {
  SUBSCRIPTION_RETURN_PARAMETER,
  SUBSCRIPTION_RETURN_VALUE,
  isMarketplaceReturn,
  readMarketplaceToken
} from '../../../../models/FeatureSubscription';
import { DEFAULT_ROUTE_PATH, ROUTE_PATHS, projectEditPath } from '../../../../models/Navigation';
import { IProjectListingRow } from '../../../../models/ProjectListing';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectListingService } from '../../../../services/ProjectListingService';
import { IProjectService } from '../../../../services/ProjectService';
import { ISubscriptionService } from '../../../../services/SubscriptionService';

export interface IAppShellProps {
  projectService: IProjectService;
  /** The paged read behind the overview table, handed straight to `ProjectList`. */
  projectListingService: IProjectListingService;
  lookupService: ILookupService;
  /** Which paid features the tenant has bought; behind the Subscription page. */
  subscriptionService: ISubscriptionService;
  /**
   * Names of the projects seen on the listing this session, keyed by id.
   *
   * Only the edit form reads them, for its title. The listing owns no rows of its own that
   * outlive it - it is server-paged and refetches on mount - so this is all that has to
   * survive the trip to the form and back.
   */
  projectNames: { [projectId: number]: string };
  savedForms: { [projectId: number]: IBuildingForm };
  features: IProjectFormFeatures;
  clientId: string;
  clientName: string;
  /** Whether this user may add or edit projects, from `canAddEditProject`. */
  canAddEdit: boolean;
  /** Whether this user may open the project form, from `canOpenProjectForm`. */
  canOpenForm: boolean;
  /** Whether this user may switch paid features on or off, from `canManageFeatureSubscriptions`. */
  canManageFeatures: boolean;
  /**
   * Whether this host can buy a paid feature: false inside Teams, which Marketplace cannot
   * send the buyer back to. Handed on to the Subscription page.
   */
  isPurchaseAvailable: boolean;
  /** `ILoginResponse.domainGroup`, which decides how the listing's cells are coloured. */
  domainGroup?: number;
  /** Save confirmation, shown by `ProjectList` as a toast. */
  notification?: string;
  /** Title of the SharePoint site chosen in `SharePointSites`, named in the site bar. */
  siteTitle: string;
  /** Absolute URL of that same site, named at the top of the listing. */
  siteUrl: string;
  /** Returns the user to the site picker. */
  onChangeSite: () => void;
  onSaveProject: (form: IBuildingForm) => Promise<void>;
  /** Records the names of a freshly loaded page, for the edit form's title. */
  onProjectsLoaded: (rows: IProjectListingRow[]) => void;
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

  // The purchase token has been handed to the API: replaced rather than pushed, so the spent
  // token leaves the URL and history alike, while `?subscription=returned` stays.
  const onTokenConsumed = React.useCallback((): void => {
    navigate(`${ROUTE_PATHS.subscription}?${SUBSCRIPTION_RETURN_PARAMETER}=${SUBSCRIPTION_RETURN_VALUE}`, { replace: true });
  }, [navigate]);

  const goToEdit = React.useCallback(
    (projectId: number): void => {
      navigate(projectEditPath(projectId));
    },
    [navigate]
  );

  const listing = (
    <ProjectList
      listingService={props.projectListingService}
      lookupService={props.lookupService}
      canAddEdit={props.canAddEdit}
      domainGroup={props.domainGroup}
      notification={props.notification}
      siteUrl={props.siteUrl}
      onAddProject={goToAdd}
      onEditProject={goToEdit}
      onProjectsLoaded={props.onProjectsLoaded}
      onDismissNotification={props.onDismissNotification}
    />
  );

  const form = (mode: 'add' | 'edit'): React.ReactElement => (
    <ProjectAddEditRoute
      lookupService={props.lookupService}
      mode={mode}
      projectNames={props.projectNames}
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
          {/* The query string sits inside the fragment,
              `#/subscription?subscription=returned&marketplaceToken=…`, which is where the
              Marketplace landing page sends the buyer back to. */}
          <Route
            path={ROUTE_PATHS.subscription}
            element={
              <Subscription
                subscriptionService={props.subscriptionService}
                canManage={props.canManageFeatures}
                isPurchaseAvailable={props.isPurchaseAvailable}
                marketplaceToken={readMarketplaceToken(location.search)}
                isReturningFromMarketplace={isMarketplaceReturn(location.search)}
                onTokenConsumed={onTokenConsumed}
              />
            }
          />
          {/* `/` and anything unrecognised land on the view the web part opened on
              before it had routes. `replace` keeps the bad path out of history. */}
          <Route path="*" element={<Navigate to={DEFAULT_ROUTE_PATH} replace />} />
        </Routes>
      </div>
    </div>
  );
}
