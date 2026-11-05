import * as React from 'react';

import styles from './ProjectAddEdit.module.scss';
import type { IKypPlanningSectionProps, IKypPlanningSectionState } from './IKypPlanningSectionProps';
import { ILookupOption } from '../../../../models/Building';
import { getErrorMessage } from '../../../../models/ApiError';
import { IKypForm, IKypSettings, toKypForm } from '../../../../models/ProjectIntegration';

/**
 * The project form's Planning (KYP) section, ported from the reference's
 * `Views/Building/AddEdit.cshtml` (the `IsPlanningSource` block).
 *
 * - The "KYP Project" toggle links the project to a KYP project (required while on).
 * - "Use alternative token" stores a project-specific KYP token instead of using the
 *   account's; Authorize lists the projects that token can see. A stored token is never sent
 *   back to the browser: leaving the field empty keeps it, unticking the box removes it on save.
 *
 * The section loads the project's stored settings itself (`GET KYP/GetProjectSettings`) and
 * hands them to the form through `onLoaded`; until then - and when no KYP token exists at
 * all - the form holds no KYP value, so the save keeps what the project already stores. KYP
 * itself is only ever called by the Web API (`POST KYP/GetProjects`).
 */
export default class KypPlanningSection extends React.Component<IKypPlanningSectionProps, IKypPlanningSectionState> {
  private _isActive: boolean = false;
  /** Bumped per projects request, so an older answer cannot overwrite a newer one. */
  private _projectsRequest: number = 0;

  public constructor(props: IKypPlanningSectionProps) {
    super(props);

    this.state = { status: 'loading', projects: [], isLoadingProjects: false };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._loadSettings();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IKypPlanningSectionProps> {
    const { status, loadError } = this.state;
    const { value } = this.props;

    return (
      <div className={styles.integrationSection}>
        {status === 'loading' && (
          <div className={styles.loadingState} role="status" aria-live="polite">
            <span className={styles.spinner} aria-hidden="true" />
            <span>Loading KYP&hellip;</span>
          </div>
        )}

        {status === 'failed' && (
          <div className={styles.formError} role="alert">
            {loadError}{' '}
            <button type="button" className={styles.secondaryButton} onClick={this._onRetry}>
              Try again
            </button>
          </div>
        )}

        {status === 'notConfigured' && (
          <div className={styles.notice} role="status">
            KYP is not set up for your organisation. Ask your administrator to add the KYP token in
            Settings.
          </div>
        )}

        {status === 'ready' && value && this._renderForm(value)}
      </div>
    );
  }

  private _renderForm(value: IKypForm): React.ReactElement {
    const { isDisabled, planningProjectError } = this.props;
    const { projects, isLoadingProjects, projectsError } = this.state;

    return (
      <>
        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.toggleField}>
              <span className={styles.toggleLabel}>KYP Project</span>
              <span className={styles.toggle}>
                <input
                  type="checkbox"
                  checked={value.isPlanningSource}
                  disabled={isDisabled}
                  onChange={this._onToggle}
                />
                <span className={styles.toggleTrack} aria-hidden="true" />
              </span>
            </label>

            {value.isPlanningSource && (
              <label className={styles.checkboxField}>
                <input
                  type="checkbox"
                  checked={value.isUseProjectToken}
                  disabled={isDisabled}
                  onChange={this._onUseProjectTokenChange}
                />
                <span>Use alternative token</span>
              </label>
            )}
          </div>

          {value.isPlanningSource && value.isUseProjectToken && (
            <div className={styles.field}>
              <label className={styles.label} htmlFor="project-kyp-token">KYP Project Auth Token</label>
              <div className={styles.inlineControls}>
                <input
                  id="project-kyp-token"
                  className={styles.input}
                  type="password"
                  value={value.projectToken}
                  placeholder={value.hasStoredProjectToken ? 'Saved - leave empty to keep it' : ''}
                  autoComplete="off"
                  disabled={isDisabled}
                  onChange={this._onTokenChange}
                />
                <button
                  type="button"
                  className={styles.primaryButton}
                  disabled={isDisabled || isLoadingProjects}
                  onClick={this._onAuthorize}
                >
                  Authorize
                </button>
              </div>
            </div>
          )}
        </div>

        {value.isPlanningSource && (
          <div className={styles.row}>
            <div className={styles.field}>
              <label className={styles.label} htmlFor="project-kyp-planningProjectId">
                Project <span className={styles.required} aria-hidden="true">*</span>
              </label>
              <select
                id="project-kyp-planningProjectId"
                className={`${styles.input} ${planningProjectError ? styles.inputInvalid : ''}`}
                value={value.planningProjectId}
                disabled={isDisabled || isLoadingProjects}
                aria-invalid={!!planningProjectError}
                onChange={this._onProjectChange}
              >
                <option value="">{isLoadingProjects ? 'Loading…' : 'Select project'}</option>
                {KypPlanningSection._withSelected(projects, value.planningProjectId).map(KypPlanningSection._renderOption)}
              </select>
              {planningProjectError && <span className={styles.fieldError} role="alert">{planningProjectError}</span>}
              {projectsError && <span className={styles.fieldError} role="alert">{projectsError}</span>}
              {!isLoadingProjects && !projectsError && projects.length === 0 && (
                <span className={styles.emptyHint}>KYP returned no projects for this token.</span>
              )}
            </div>
          </div>
        )}
      </>
    );
  }

  private _loadSettings(): void {
    const { integrationService, projectId } = this.props;

    this.setState({ status: 'loading', loadError: undefined });
    integrationService.getKypSettings(projectId).then(
      (settings: IKypSettings): void => {
        if (!this._isActive) {
          return;
        }
        // The reference shows the section when a token exists at all - the account's or the project's own.
        if (!settings.isConfigured && !settings.hasProjectToken) {
          this.setState({ status: 'notConfigured' });
          return;
        }

        const value: IKypForm = toKypForm(settings);
        this.setState({ status: 'ready' });
        this.props.onLoaded(value);
        if (value.isPlanningSource) {
          this._loadProjects(value);
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ status: 'failed', loadError: getErrorMessage(error, 'The KYP settings could not be loaded.') });
        }
      }
    );
  }

  private _loadProjects(value: IKypForm): void {
    const request: number = ++this._projectsRequest;

    this.setState({ isLoadingProjects: true, projectsError: undefined });
    this.props.integrationService
      .getKypProjects({
        projectId: Math.max(0, this.props.projectId),
        useProjectToken: value.isUseProjectToken,
        authToken: value.isUseProjectToken ? value.projectToken.trim() : ''
      })
      .then(
        (projects: ILookupOption[]): void => {
          if (this._isActive && request === this._projectsRequest) {
            this.setState({ projects, isLoadingProjects: false });
          }
        },
        (error: unknown): void => {
          if (this._isActive && request === this._projectsRequest) {
            this.setState({
              projects: [],
              isLoadingProjects: false,
              projectsError: getErrorMessage(error, 'The KYP projects could not be loaded.')
            });
          }
        }
      );
  }

  private _update(changes: Partial<IKypForm>): IKypForm | undefined {
    if (!this.props.value) {
      return undefined;
    }
    const next: IKypForm = { ...this.props.value, ...changes };
    this.props.onChange(next);
    return next;
  }

  private _onRetry = (): void => {
    this._loadSettings();
  };

  private _onToggle = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const next: IKypForm | undefined = this._update({ isPlanningSource: event.target.checked });
    if (next && next.isPlanningSource) {
      this._loadProjects(next);
    }
  };

  private _onUseProjectTokenChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const next: IKypForm | undefined = this._update({ isUseProjectToken: event.target.checked });
    // Back to the account's token: its projects are the ones on offer again. With the box
    // ticked, the list changes on Authorize, once a token has been typed.
    if (next && !next.isUseProjectToken) {
      this._loadProjects(next);
    }
  };

  private _onTokenChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this._update({ projectToken: event.target.value });
  };

  private _onAuthorize = (): void => {
    if (this.props.value) {
      this._loadProjects(this.props.value);
    }
  };

  private _onProjectChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this._update({ planningProjectId: event.target.value });
  };

  private static _withSelected(options: ILookupOption[], selectedId: string): ILookupOption[] {
    if (!selectedId || options.some((option: ILookupOption) => option.id === selectedId)) {
      return options;
    }
    return options.concat([{ id: selectedId, name: selectedId }]);
  }

  private static _renderOption(option: ILookupOption): React.ReactElement {
    return <option key={option.id} value={option.id}>{option.name}</option>;
  }
}
