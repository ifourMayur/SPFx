import * as React from 'react';
import { HashRouter } from 'react-router-dom';

import styles from './XsProject.module.scss';
import type { IXsProjectProps } from './IXsProjectProps';
import type { IXsProjectState } from './IXsProjectState';
import { escape } from '@microsoft/sp-lodash-subset';
import AppShell from './AppShell';
import Login from '../Login/Login';
import welcomeDark from '../../assets/welcome-dark.png';
import welcomeLight from '../../assets/welcome-light.png';
import { ILoginResponse } from '../../../../models/Auth';
import { IBuildingForm, IProjectFormFeatures, toDomainGroup } from '../../../../models/Building';
import { canAddEditProject, canOpenProjectForm } from '../../../../models/Permissions';
import { IProjectListRow, SAMPLE_PROJECT_ROWS, upsertProjectRow } from '../../../../models/ProjectListRow';

/** `BuildingResource.msgAddProject` / `msgUpdateProject`, including their double spaces. */
const SAVE_MESSAGES = {
  added: 'Project  Added successfully!',
  updated: 'Project  Updated successfully!'
};

/**
 * Root component rendered by `XsProjectWebPart`.
 *
 * Owns the authenticated/not-authenticated view state: the router, the shared `Menu` and
 * the pages it navigates to are only ever rendered once `Login` reports a successful
 * sign-in, using the `onLoginSucceeded`/`onLoginFailed` callbacks `Login` already exposes
 * rather than a second authentication mechanism.
 *
 * It also owns the project rows and the forms saved during the session, because routing
 * unmounts the page that would otherwise hold them - see `IXsProjectState`.
 */
export default class XsProject extends React.Component<IXsProjectProps, IXsProjectState> {
  public constructor(props: IXsProjectProps) {
    super(props);

    this.state = {
      isAuthenticated: false,
      rows: SAMPLE_PROJECT_ROWS,
      savedForms: {}
    };
  }

  public render(): React.ReactElement<IXsProjectProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName,
      projectService,
      lookupService,
      loginService
    } = this.props;
    const { isAuthenticated, login, rows, savedForms, notification } = this.state;
    const userRoleId: number | undefined = login?.userRoleId;

    return (
      <section className={`${styles.xsProject}`}>
        <Login
          loginService={loginService}
          onLoginSucceeded={this._onLoginSucceeded}
          onLoginFailed={this._onLoginFailed}
        />

        {isAuthenticated && (
          // The app's routes live in the URL fragment: a SPFx web part is a guest on a
          // SharePoint page whose path SharePoint itself owns and navigates.
          <HashRouter>
            <AppShell
              projectService={projectService}
              lookupService={lookupService}
              rows={rows}
              savedForms={savedForms}
              features={this._getFeatures()}
              clientId={login ? String(login.id) : ''}
              clientName={login?.userName || userDisplayName}
              canAddEdit={canAddEditProject(userRoleId)}
              canOpenForm={canOpenProjectForm(userRoleId)}
              notification={notification}
              onSaveProject={this._onSaveProject}
              onToggleFavorite={this._onToggleFavorite}
              onDismissNotification={this._onDismissNotification}
            />
          </HashRouter>
        )}

        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? welcomeDark : welcomeLight} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }

  /**
   * The reference view's server-side `@if` gates, resolved from the sign-in response.
   *
   * `domainGroup` decides whether the postcode is required and whether Spatial Breakdown
   * and the ISO 19650 toggle appear. The fourth folder level corresponds to
   * `ViewBag.IsGrandChildFolderEnabled`, which the Web API does not expose to this client
   * yet, so it stays on to keep the whole cascade reviewable.
   */
  private _getFeatures(): IProjectFormFeatures {
    return {
      domainGroup: toDomainGroup(this.state.login?.domainGroup),
      isGrandChildFolderEnabled: true
    };
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` itself uses for its own click handler.
  private _onLoginSucceeded = (response: ILoginResponse): void => {
    this.setState({ isAuthenticated: true, login: response });
  };

  private _onLoginFailed = (): void => {
    this.setState({ isAuthenticated: false, login: undefined });
  };

  /**
   * Records a saved project.
   *
   * Async because this is where a `BuildingService.save()` call belongs: the form already
   * awaits this promise, shows its saving state while it is pending, and reports a
   * rejection as a form-level error. Nothing is sent anywhere today.
   */
  private _onSaveProject = async (form: IBuildingForm): Promise<void> => {
    const isUpdate: boolean = form.id > 0;

    this.setState((state: IXsProjectState) => {
      const result = upsertProjectRow(state.rows, form);

      return {
        rows: result.rows,
        // Stored under the id the row actually got, so a project created with `id: 0`
        // can be reopened for editing under its new id.
        savedForms: { ...state.savedForms, [result.id]: { ...form, id: result.id } },
        notification: isUpdate ? SAVE_MESSAGES.updated : SAVE_MESSAGES.added
      };
    });
  };

  private _onToggleFavorite = (projectId: number): void => {
    this.setState((state: IXsProjectState) => ({
      rows: state.rows.map((row: IProjectListRow) =>
        row.id === projectId ? { ...row, isFavorite: !row.isFavorite } : row
      )
    }));
  };

  private _onDismissNotification = (): void => {
    this.setState({ notification: undefined });
  };
}
