import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';

import styles from './Project.module.scss';
import type { IProjectProps } from './IProjectProps';
import { IProjectListState } from '../../../../models/Project';
import { getErrorMessage } from '../../../../models/ApiError';

/** Source name used for SPFx log entries emitted by this component. */
const LOG_SOURCE: string = 'Project';

/**
 * Project view opened from the shared `Menu`.
 *
 * Lists projects through `IProjectService`, the same abstraction `XsProjectWebPart`
 * injects into every other component - no direct HTTP/SPFx access here.
 */
export default class Project extends React.Component<IProjectProps, IProjectListState> {
  /** Guards against completing a request after the component has been unmounted. */
  private _isActive: boolean = false;

  public constructor(props: IProjectProps) {
    super(props);

    this.state = { projects: [], isLoading: false };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._loadProjects();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IProjectProps> {
    const { projects, isLoading, errorMessage } = this.state;

    return (
      <section className={styles.project}>
        <h3 className={styles.title}>Projects</h3>

        {isLoading && (
          <div className={styles.status} role="status" aria-live="polite">Loading projects&hellip;</div>
        )}

        {!isLoading && errorMessage && (
          <div className={`${styles.status} ${styles.failure}`} role="alert">{errorMessage}</div>
        )}

        {!isLoading && !errorMessage && projects.length === 0 && (
          <div className={styles.status}>No projects found.</div>
        )}

        {!isLoading && !errorMessage && projects.length > 0 && (
          <ul className={styles.list}>
            {projects.map(project => (
              <li key={project.id} className={styles.listItem}>
                <span className={styles.name}>{project.name}</span>
                <span className={styles.badge}>{project.status}</span>
              </li>
            ))}
          </ul>
        )}
      </section>
    );
  }

  private _loadProjects(): void {
    this.setState({ isLoading: true, errorMessage: undefined });

    this.props.projectService.getProjects()
      .then((projects) => {
        if (this._isActive) {
          this.setState({ isLoading: false, projects });
        }
      })
      .catch((error: unknown) => {
        Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));

        if (this._isActive) {
          this.setState({ isLoading: false, projects: [], errorMessage: getErrorMessage(error, 'Could not load projects.') });
        }
      });
  }
}
