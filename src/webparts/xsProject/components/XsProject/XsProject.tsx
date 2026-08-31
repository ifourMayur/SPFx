import * as React from 'react';
import styles from './XsProject.module.scss';
import type { IXsProjectProps } from './IXsProjectProps';
import type { IXsProjectState } from './IXsProjectState';
import { escape } from '@microsoft/sp-lodash-subset';
import Login from '../Login/Login';
import Menu from '../Menu/Menu';
import Project from '../Project/Project';
import Document from '../Document/Document';
import Search from '../Search/Search';
import Suppliers from '../Suppliers/Suppliers';
import welcomeDark from '../../assets/welcome-dark.png';
import welcomeLight from '../../assets/welcome-light.png';
import { AppView } from '../../../../models/Navigation';

/**
 * Root component rendered by `XsProjectWebPart`.
 *
 * Owns the authenticated/not-authenticated view state: the shared `Menu` and the pages it
 * navigates to are only ever rendered once `Login` reports a successful sign-in, using the
 * `onLoginSucceeded`/`onLoginFailed` callbacks `Login` already exposes rather than a second
 * authentication mechanism.
 */
export default class XsProject extends React.Component<IXsProjectProps, IXsProjectState> {
  public constructor(props: IXsProjectProps) {
    super(props);

    this.state = { isAuthenticated: false, activeView: 'projectDashboard' };
  }

  public render(): React.ReactElement<IXsProjectProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName,
      projectService,
      loginService
    } = this.props;
    const { isAuthenticated, activeView } = this.state;

    return (
      <section className={`${styles.xsProject}`}>
        <Login
          loginService={loginService}
          onLoginSucceeded={this._onLoginSucceeded}
          onLoginFailed={this._onLoginFailed}
        />

        {isAuthenticated && (
          <div className={styles.appLayout}>
            <Menu activeView={activeView} onNavigate={this._onNavigate} />
            <div className={styles.content}>
              {activeView === 'projectDashboard' && <Project projectService={projectService} />}
              {activeView === 'projectDocument' && <Document />}
              {activeView === 'search' && <Search />}
              {activeView === 'suppliers' && <Suppliers />}
            </div>
          </div>
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

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` itself uses for its own click handler.
  private _onLoginSucceeded = (): void => {
    this.setState({ isAuthenticated: true });
  };

  private _onLoginFailed = (): void => {
    this.setState({ isAuthenticated: false });
  };

  private _onNavigate = (view: AppView): void => {
    this.setState({ activeView: view });
  };
}
