import * as React from 'react';
import { HashRouter } from 'react-router-dom';

import styles from './XsProject.module.scss';
import type { IXsProjectProps } from './IXsProjectProps';
import type { IXsProjectState } from './IXsProjectState';
import { escape } from '@microsoft/sp-lodash-subset';
import AppShell from './AppShell';
import Login from '../Login/Login';
import SharePointSites from '../SharePointSites/SharePointSites';
import welcomeDark from '../../assets/welcome-dark.png';
import welcomeLight from '../../assets/welcome-light.png';
import { ILoginResponse } from '../../../../models/Auth';
import { IBuildingForm, IProjectFormFeatures, toDomainGroup } from '../../../../models/Building';
import { IProjectSaveContext, IProjectSaveResult, toHostName } from '../../../../models/BuildingSave';
import { IFolderProvisionFailure, IFolderProvisionResult } from '../../../../models/SharePointFolder';
import { canAddEditProject, canOpenProjectForm } from '../../../../models/Permissions';
import { IProjectListRow, SAMPLE_PROJECT_ROWS, upsertProjectRow } from '../../../../models/ProjectListRow';
import { ISharePointSite } from '../../../../models/SharePointSite';

/** `BuildingResource.msgAddProject` / `msgUpdateProject`, including their double spaces. */
const SAVE_MESSAGES = {
  added: 'Project  Added successfully!',
  updated: 'Project  Updated successfully!'
};

/**
 * Root component rendered by `XsProjectWebPart`.
 *
 * Owns the view state that gates the app, in two steps. The router, the shared `Menu` and
 * the pages it navigates to are only ever rendered once `Login` reports a successful
 * sign-in - using the `onLoginSucceeded`/`onLoginFailed` callbacks `Login` already exposes
 * rather than a second authentication mechanism - and then only once `SharePointSites`
 * reports the site the user chose to work in.
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
      loginService,
      sharePointSiteService
    } = this.props;
    const { isAuthenticated, login, selectedSite, isChangingSite, rows, savedForms, notification } =
      this.state;
    const userRoleId: number | undefined = login?.userRoleId;
    // The app is shown only when a site has been chosen and the user is not in the middle
    // of changing it.
    const isSiteChosen: boolean = !!selectedSite && !isChangingSite;

    return (
      <section className={`${styles.xsProject}`}>
        <Login
          loginService={loginService}
          onLoginSucceeded={this._onLoginSucceeded}
          onLoginFailed={this._onLoginFailed}
        />

        {isAuthenticated && !isSiteChosen && (
          <SharePointSites
            siteService={sharePointSiteService}
            selectedUrl={selectedSite?.url}
            onSiteSelected={this._onSiteSelected}
          />
        )}

        {isAuthenticated && isSiteChosen && selectedSite && (
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
              siteTitle={selectedSite.title}
              onChangeSite={this._onChangeSite}
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

  private _onSiteSelected = (site: ISharePointSite): void => {
    this.setState({ selectedSite: site, isChangingSite: false });
  };

  /**
   * Sends the user back to the picker.
   *
   * Only the site changes: the sign-in, the project rows and the forms saved this session
   * all survive, because changing site is a change of where the user is working - not a
   * sign-out.
   */
  private _onChangeSite = (): void => {
    this.setState({ isChangingSite: true });
  };

  /**
   * Saves a project through `POST api/Building`, then records the result locally.
   *
   * The API call is awaited rather than fired and forgotten: `ProjectAddEdit` shows its
   * saving state while this promise is pending, navigates back to the listing when it
   * resolves, and renders a rejection as a form-level error with the user's input intact.
   * So everything below the call only runs once the API has accepted the save - which is
   * also why nothing is caught here.
   *
   * The listing rows are still local sample data, so the saved project is folded into them
   * here, under the id the API assigned rather than a locally allocated one.
   */
  private _onSaveProject = async (form: IBuildingForm): Promise<void> => {
    const result: IProjectSaveResult = await this.props.buildingService.saveProject(
      form,
      this._getSaveContext()
    );

    const saved: IBuildingForm = { ...form, id: result.projectId };

    this.setState((state: IXsProjectState) => ({
      // `upsertProjectRow` adds an id it has not seen rather than dropping it, which is
      // what puts a newly created project at the top of the list.
      rows: upsertProjectRow(state.rows, saved).rows,
      // Stored under the id the API gave it, so a project created with `id: 0` can be
      // reopened for editing under its new id.
      savedForms: { ...state.savedForms, [result.projectId]: saved },
      notification: XsProject._toSaveNotification(result, this.state.selectedSite)
    }));
  };

  /**
   * The message shown on the listing after a save.
   *
   * Names where the folders went, because that is the only way to check them: the Web API
   * is not told their ids yet, so nothing else in the app can show them. A site with no
   * folders to report - an update, or a template with an empty tree - just gets the
   * reference application's own save message.
   */
  private static _toSaveNotification(
    result: IProjectSaveResult,
    site: ISharePointSite | undefined
  ): string {
    const saved: string = result.isCreated ? SAVE_MESSAGES.added : SAVE_MESSAGES.updated;
    const provision: IFolderProvisionResult | undefined = result.folderProvision;

    if (!provision) {
      return saved;
    }

    const provisioned: number = provision.created.length + provision.skipped.length;
    const failed: string[] = provision.failed.map((failure: IFolderProvisionFailure) => failure.path);
    const where: string = site ? ` in ${site.title}` : '';
    const parts: string[] = [saved];

    if (provisioned > 0) {
      parts.push(`${provisioned} folder(s) ready${where} - check ${provision.rootUrl}`);
    }

    if (failed.length > 0) {
      // Named rather than counted: which folder failed is what tells the user whether one
      // branch was refused or the whole library was.
      parts.push(`${failed.length} folder(s) could not be created: ${failed.join(', ')}`);
    }

    return parts.join(' ');
  }

  /**
   * The values `POST api/Building` needs that the form does not hold.
   *
   * The reference reads all of these out of the MVC session (step 3 of the process
   * document); a SPFx bundle has no session, so they come from the sign-in response and
   * the hosting page instead.
   */
  private _getSaveContext(): IProjectSaveContext {
    const login: ILoginResponse | undefined = this.state.login;

    return {
      clientId: login?.id || 0,
      clientName: login?.userName || this.props.userDisplayName,
      accountId: login?.accountId || 0,
      // `Request.Host.Value` in the reference: the host of the application the user is
      // signed in to, falling back to the SharePoint page hosting this web part.
      domainName: toHostName(login?.domainUrl || '') || window.location.host,
      companyName: login?.masterCompanyName || login?.companyName || '',
      // The reference resolves a file path on the server; a browser has no equivalent.
      companyLogo: '',
      // Where a new project's folders are created. Always set by the time a save can run:
      // the app is not rendered until a site has been chosen.
      site: this.state.selectedSite
    };
  }

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
