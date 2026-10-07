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
import { IProjectSaveContext, IProjectSaveResult, toProjectSaveContext } from '../../../../models/BuildingSave';
import { IFolderProvisionFailure, IFolderProvisionResult } from '../../../../models/SharePointFolder';
import { canAddEditProject, canOpenProjectForm } from '../../../../models/Permissions';
import { IProjectListingRow } from '../../../../models/ProjectListing';
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
 * It also owns the forms saved during the session, and the names of the projects the
 * listing has shown, because routing unmounts the page that would otherwise hold them -
 * see `IXsProjectState`. The listing's rows are not among them: it is server-paged and
 * refetches when the router brings the user back to it.
 */
export default class XsProject extends React.Component<IXsProjectProps, IXsProjectState> {
  public constructor(props: IXsProjectProps) {
    super(props);

    this.state = {
      isAuthenticated: false,
      projectNames: {},
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
      projectListingService,
      lookupService,
      loginService,
      sharePointSiteService
    } = this.props;
    const {
      isAuthenticated,
      login,
      selectedSite,
      isChangingSite,
      projectNames,
      savedForms,
      notification
    } = this.state;
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
              projectListingService={projectListingService}
              lookupService={lookupService}
              projectNames={projectNames}
              savedForms={savedForms}
              features={this._getFeatures()}
              clientId={login ? String(login.id) : ''}
              clientName={login?.userName || userDisplayName}
              canAddEdit={canAddEditProject(userRoleId)}
              canOpenForm={canOpenProjectForm(userRoleId)}
              domainGroup={login?.domainGroup}
              notification={notification}
              siteTitle={selectedSite.title}
              siteUrl={selectedSite.url}
              onChangeSite={this._onChangeSite}
              onSaveProject={this._onSaveProject}
              onProjectsLoaded={this._onProjectsLoaded}
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
   * Only the site changes: the sign-in and the forms saved this session both survive,
   * because changing site is a change of where the user is working - not a sign-out.
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
   * Nothing is folded into the listing here: it is server-paged, and the router unmounts
   * it on the way to the form, so it refetches - and shows the saved project - on the way
   * back. Only the name is kept, under the id the API assigned rather than a locally
   * allocated one, so the form can title itself if the project is reopened.
   */
  private _onSaveProject = async (form: IBuildingForm): Promise<void> => {
    const result: IProjectSaveResult = await this.props.buildingService.saveProject(
      form,
      this._getSaveContext()
    );

    const saved: IBuildingForm = { ...form, id: result.projectId };

    this.setState((state: IXsProjectState) => ({
      projectNames: { ...state.projectNames, [result.projectId]: saved.buildingName.trim() },
      // Stored under the id the API gave it, so a project created with `id: 0` can be
      // reopened for editing under its new id.
      savedForms: { ...state.savedForms, [result.projectId]: saved },
      notification: XsProject._toSaveNotification(result)
    }));
  };

  /**
   * The toast shown on the listing after a save.
   *
   * Just the outcome: the reference's own save message, and nothing about what the Web API
   * answered. The SharePoint ids bound onto the project's folder tree used to be printed
   * here because nothing else in the app could show them, but that is a development
   * detail - `BuildingService` logs the whole result - not something to put in front of
   * the user on every save.
   *
   * Failed folders are the one exception, and stay named. The folder steps deliberately
   * never fail the save (the project exists by then, so reporting failure would invite a
   * duplicate), which makes this the only signal the user gets that a folder they will
   * look for is not there.
   */
  private static _toSaveNotification(result: IProjectSaveResult): string {
    const saved: string = result.isCreated ? SAVE_MESSAGES.added : SAVE_MESSAGES.updated;
    const provision: IFolderProvisionResult | undefined = result.folderProvision;

    if (!provision || provision.failed.length === 0) {
      return saved;
    }

    // Named rather than counted: which folder failed is what tells the user whether one
    // branch was refused or the whole library was.
    const failed: string[] = provision.failed.map((failure: IFolderProvisionFailure) => failure.path);

    return `${saved} ${failed.length} folder(s) could not be created: ${failed.join(', ')}`;
  }

  /**
   * The values `POST api/Building` needs that the form does not hold; the mapping itself
   * is `toProjectSaveContext`, and this only supplies what the hosting page knows.
   */
  private _getSaveContext(): IProjectSaveContext {
    return toProjectSaveContext(this.state.login, this.state.selectedSite, {
      userDisplayName: this.props.userDisplayName,
      hostName: window.location.host
    });
  }

  /**
   * Remembers the names of the page the listing just loaded.
   *
   * The edit form is reached from a row, and the router unmounts the listing on the way
   * there - so without this the form would open with no title for a project the user is
   * looking straight at.
   */
  private _onProjectsLoaded = (rows: IProjectListingRow[]): void => {
    this.setState((state: IXsProjectState) => {
      const projectNames: { [projectId: number]: string } = { ...state.projectNames };

      rows.forEach((row: IProjectListingRow): void => {
        projectNames[row.buildingId] = row.buildingName;
      });

      return { projectNames };
    });
  };

  /** Clears the save toast, whether dismissed by the user or by its own timer. */
  private _onDismissNotification = (): void => {
    this.setState({ notification: undefined });
  };
}
