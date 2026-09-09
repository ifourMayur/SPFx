import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';

import styles from './SharePointSites.module.scss';
import type { ISharePointSitesProps } from './ISharePointSitesProps';
import type { ISharePointSitesState } from './ISharePointSitesState';
import { ISharePointSite } from '../../../../models/SharePointSite';

/** Source name used for SPFx log entries emitted by this component. */
const LOG_SOURCE: string = 'SharePointSites';

/**
 * Message shown when the site query fails.
 *
 * Deliberately fixed rather than built from the error: a SharePoint failure body can carry
 * server detail, and the correlation id belongs in the log, not on screen. The error
 * itself goes to `Log.error`, where the SPFx console shows it during development.
 */
const FAILURE_MESSAGE: string =
  'Could not load the SharePoint sites for your account. Check that you are signed in to this tenant, then try again.';

/**
 * Chooses the SharePoint site the session works in.
 *
 * The step between sign-in and the app: `XsProject` renders this once `Login` reports a
 * successful sign-in and shows the menu and its pages only after a site has been chosen,
 * so the picker gates the app exactly as sign-in does.
 *
 * Owns nothing but view state - `ISharePointSiteService` runs the query, so no SharePoint
 * REST call, URL or transport appears in this class. Loading, failure and empty-tenant
 * outcomes are all distinct on screen: an empty dropdown would otherwise be the answer to
 * three different questions.
 */
export default class SharePointSites extends React.Component<ISharePointSitesProps, ISharePointSitesState> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;

  public constructor(props: ISharePointSitesProps) {
    super(props);

    this.state = {
      isLoading: false,
      sites: [],
      selectedUrl: props.selectedUrl || '',
      hasFailed: false
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._runLoad();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<ISharePointSitesProps> {
    const { isLoading, sites, selectedUrl, hasFailed } = this.state;

    return (
      <section className={styles.sharePointSites}>
        <h3 className={styles.title}>Choose a SharePoint site</h3>

        {isLoading && (
          <div className={`${styles.status} ${styles.pending}`} role="status" aria-live="polite">
            <span className={styles.spinner} aria-hidden="true" />
            <span>Loading SharePoint sites&hellip;</span>
          </div>
        )}

        {!isLoading && hasFailed && (
          <div className={`${styles.status} ${styles.failure}`} role="alert">
            <div>{FAILURE_MESSAGE}</div>
            <button type="button" className={styles.action} onClick={this._onRetryClick}>
              Try again
            </button>
          </div>
        )}

        {!isLoading && !hasFailed && sites.length === 0 && (
          <div className={`${styles.status} ${styles.empty}`} role="status">
            <div>
              No SharePoint sites were found for your account. A site created in the last few minutes may
              not be indexed yet.
            </div>
            <button type="button" className={styles.action} onClick={this._onRetryClick}>
              Try again
            </button>
          </div>
        )}

        {!isLoading && !hasFailed && sites.length > 0 && (
          <>
            <div className={styles.field}>
              <label className={styles.label} htmlFor="sharepoint-site">
                SharePoint site
              </label>
              <select
                id="sharepoint-site"
                className={styles.select}
                value={selectedUrl}
                onChange={this._onSiteChange}
              >
                <option value="">Select a site</option>
                {sites.map((site: ISharePointSite) => (
                  <option key={site.url} value={site.url}>
                    {site.title}
                  </option>
                ))}
              </select>
            </div>

            <button
              id="sharepoint-site-continue"
              type="button"
              className={styles.action}
              // Nothing to continue to without a site, and the app below expects one.
              disabled={!selectedUrl}
              onClick={this._onContinueClick}
            >
              Continue
            </button>
          </>
        )}
      </section>
    );
  }

  /**
   * Runs the site query and reduces the outcome to view state.
   *
   * A failure is recorded as `hasFailed` rather than as an empty list, so the retry the
   * user is offered is genuinely offered - not hidden behind a dropdown that merely looks
   * empty.
   */
  private async _load(): Promise<void> {
    this.setState({ isLoading: true, hasFailed: false });

    let sites: ISharePointSite[] | undefined;

    try {
      sites = await this.props.siteService.getSites();
    } catch (error) {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    }

    if (!this._isActive) {
      return;
    }

    this.setState((state: ISharePointSitesState) => ({
      isLoading: false,
      hasFailed: sites === undefined,
      sites: sites || [],
      // A site that is no longer offered cannot stay selected, or Continue would hand the
      // app a site the user can no longer see.
      selectedUrl: (sites || []).some((site: ISharePointSite) => site.url === state.selectedUrl)
        ? state.selectedUrl
        : ''
    }));
  }

  /**
   * Starts the query from a synchronous caller.
   *
   * `_load` already reports every query failure through the component state; this catch
   * only covers an unexpected render-time error, so it can never surface as an unhandled
   * rejection.
   */
  private _runLoad(): void {
    this._load().catch((error: unknown): void => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` uses for its own click handler.
  private _onRetryClick = (): void => {
    this._runLoad();
  };

  private _onSiteChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ selectedUrl: event.target.value });
  };

  private _onContinueClick = (): void => {
    const chosen: ISharePointSite | undefined = this.state.sites.filter(
      (site: ISharePointSite) => site.url === this.state.selectedUrl
    )[0];

    if (chosen) {
      this.props.onSiteSelected(chosen);
    }
  };
}
