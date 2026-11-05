import * as React from 'react';

import styles from './ProjectAddEdit.module.scss';
import type { IEdControlsSectionProps, IEdControlsSectionState } from './IEdControlsSectionProps';
import FolderCascade from '../FolderCascade/FolderCascade';
import MultiSelect from '../MultiSelect/MultiSelect';
import { IFolderSelection, ILookupOption } from '../../../../models/Building';
import { getErrorMessage } from '../../../../models/ApiError';
import {
  FOLDER_LEVELS,
  FolderLevel,
  IFolderLevelOptions,
  clearDeeperLevelOptions,
  createEmptyFolderLevelOptions
} from '../../../../models/FolderTree';
import {
  IEdControlsForm,
  IEdControlsSettings,
  ITagRow,
  ITagRowErrors,
  createTagRow,
  toEdControlsForm,
  toTagOptions,
  validateTagRows
} from '../../../../models/ProjectIntegration';

/**
 * The project form's ED Controls / Audits section, ported from the reference's
 * `Views/Building/AddEdit.cshtml` (the `IsTicketSource` block).
 *
 * - The "ED controls" toggle links the project to an ED Controls project (required while on).
 * - In edit mode, tag rows file audits carrying the chosen tags into a folder path and
 *   supplier; "Periodic" reveals the document name. "Add tag row" needs every row complete,
 *   as the reference's `validateAllRowsOnEdit` does. Unlike the reference, a row can be removed.
 *
 * The section loads the project's stored settings itself (`GET EDControls/GetProjectSettings`)
 * and hands them to the form through `onLoaded`; until then - and when ED Controls is not
 * configured for the account - the form holds no ED Controls value, so the save keeps what
 * the project already stores. Projects and tags come from the Web API, which calls ED
 * Controls with the account's credentials.
 */
export default class EdControlsSection extends React.Component<IEdControlsSectionProps, IEdControlsSectionState> {
  private _isActive: boolean = false;

  public constructor(props: IEdControlsSectionProps) {
    super(props);

    this.state = {
      status: 'loading',
      hasSavedLink: false,
      projects: [],
      isLoadingProjects: false,
      tags: [],
      isLoadingTags: false,
      rowFolderOptions: {},
      rowLoadingLevels: {},
      addRowErrors: {}
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._loadSettings();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IEdControlsSectionProps> {
    const { status, loadError } = this.state;
    const { value } = this.props;

    return (
      <div className={styles.integrationSection}>
        {status === 'loading' && (
          <div className={styles.loadingState} role="status" aria-live="polite">
            <span className={styles.spinner} aria-hidden="true" />
            <span>Loading ED Controls&hellip;</span>
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
            ED Controls is not set up for your organisation. Ask your administrator to add the ED
            Controls credentials in Settings.
          </div>
        )}

        {status === 'ready' && value && this._renderForm(value)}
      </div>
    );
  }

  private _renderForm(value: IEdControlsForm): React.ReactElement {
    const { isDisabled, ticketProjectError } = this.props;
    const { projects, isLoadingProjects, projectsError } = this.state;
    const isEditMode: boolean = this.props.projectId > 0;

    return (
      <>
        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.toggleField}>
              <span className={styles.toggleLabel}>ED controls</span>
              <span className={styles.toggle}>
                <input
                  type="checkbox"
                  checked={value.isTicketSource}
                  disabled={isDisabled}
                  onChange={this._onToggle}
                />
                <span className={styles.toggleTrack} aria-hidden="true" />
              </span>
            </label>
          </div>

          {value.isTicketSource && (
            <div className={styles.field}>
              <label className={styles.label} htmlFor="project-edControls-ticketProjectId">
                Project <span className={styles.required} aria-hidden="true">*</span>
              </label>
              <select
                id="project-edControls-ticketProjectId"
                className={`${styles.input} ${ticketProjectError ? styles.inputInvalid : ''}`}
                value={value.ticketProjectId}
                disabled={isDisabled || isLoadingProjects}
                aria-invalid={!!ticketProjectError}
                onChange={this._onProjectChange}
              >
                <option value="">{isLoadingProjects ? 'Loading…' : 'Select project'}</option>
                {EdControlsSection._withSelected(projects, value.ticketProjectId).map(EdControlsSection._renderOption)}
              </select>
              {ticketProjectError && <span className={styles.fieldError} role="alert">{ticketProjectError}</span>}
              {projectsError && <span className={styles.fieldError} role="alert">{projectsError}</span>}
              {!isLoadingProjects && !projectsError && projects.length === 0 && (
                <span className={styles.emptyHint}>ED Controls returned no projects for this account.</span>
              )}
            </div>
          )}
        </div>

        {value.isTicketSource && isEditMode && this._renderTagRows(value)}
        {isEditMode && this.state.hasSavedLink && this.props.onOpenAudits && (
          <div className={styles.actions}>
            <button type="button" className={styles.secondaryButton} disabled={isDisabled} onClick={this.props.onOpenAudits}>
              Manage audits
            </button>
            <span className={styles.emptyHint}>Opens the audits of the saved ED Controls project; unsaved changes here are not kept.</span>
          </div>
        )}
        {value.isTicketSource && !isEditMode && (
          <span className={styles.emptyHint}>Tag rows can be added once the project has been saved.</span>
        )}
      </>
    );
  }

  private _renderTagRows(value: IEdControlsForm): React.ReactElement {
    const { isDisabled } = this.props;
    const { tagsError, isLoadingTags } = this.state;

    return (
      <>
        {tagsError && <div className={styles.formError} role="alert">{tagsError}</div>}
        {value.tagRows.map((row: ITagRow) => this._renderTagRow(value, row))}
        <div className={styles.actions}>
          <button
            type="button"
            className={styles.secondaryButton}
            disabled={isDisabled || isLoadingTags}
            onClick={this._onAddRow}
          >
            + Add tag row
          </button>
        </div>
      </>
    );
  }

  private _renderTagRow(value: IEdControlsForm, row: ITagRow): React.ReactElement {
    const { isDisabled, suppliers, isGrandChildFolderEnabled } = this.props;
    const { tags, rowFolderOptions, rowLoadingLevels } = this.state;
    const errors: ITagRowErrors = {
      ...(this.state.addRowErrors[row.key] || {}),
      ...((this.props.tagRowErrors || {})[row.key] || {})
    };
    const idPrefix: string = `project-edControls-${row.key}`;

    return (
      <div className={styles.tagRow} key={row.key}>
        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.checkboxField}>
              <input
                type="checkbox"
                checked={row.periodic}
                disabled={isDisabled}
                onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                  this._updateRow(row.key, { periodic: event.target.checked })}
              />
              <span>Periodic</span>
            </label>
            {row.periodic && (
              <>
                <label className={styles.label} htmlFor={`${idPrefix}-documentName`}>Name</label>
                <input
                  id={`${idPrefix}-documentName`}
                  className={styles.input}
                  type="text"
                  value={row.documentName}
                  placeholder="Name"
                  autoComplete="off"
                  disabled={isDisabled}
                  onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
                    this._updateRow(row.key, { documentName: event.target.value })}
                />
              </>
            )}
          </div>

          <div className={styles.field}>
            <MultiSelect
              inputId={`${idPrefix}-tags`}
              label="Tag"
              options={toTagOptions(tags, row.tags)}
              selectedIds={row.tags}
              isDisabled={isDisabled}
              placeholder="Select tags"
              emptyText={value.ticketProjectId ? 'ED Controls returned no tags for this project.' : 'Select an ED Controls project first.'}
              onChange={(selected: string[]) => this._updateRow(row.key, { tags: selected })}
            />
            {errors.tags && <span className={styles.fieldError} role="alert">{errors.tags}</span>}
          </div>
        </div>

        <FolderCascade
          idPrefix={idPrefix}
          options={rowFolderOptions[row.key] || createEmptyFolderLevelOptions()}
          suppliers={suppliers}
          value={row.folder}
          isGrandChildFolderEnabled={isGrandChildFolderEnabled}
          loadingLevels={rowLoadingLevels[row.key] || []}
          isDisabled={isDisabled}
          onChange={(selection: IFolderSelection, changedLevel: FolderLevel) =>
            this._onRowFolderChange(row.key, selection, changedLevel)}
          onSupplierChange={(supplierId: string) =>
            this._updateRow(row.key, { folder: { ...row.folder, supplierId } })}
        />
        {errors.folderId && <span className={styles.fieldError} role="alert">{errors.folderId}</span>}

        {value.tagRows.length > 1 && (
          <div className={styles.tagRowActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              disabled={isDisabled}
              onClick={() => this._onRemoveRow(row.key)}
            >
              Remove row
            </button>
          </div>
        )}
      </div>
    );
  }

  private _loadSettings(): void {
    const { integrationService, projectId } = this.props;

    this.setState({ status: 'loading', loadError: undefined });
    integrationService.getEdControlsSettings(projectId).then(
      (settings: IEdControlsSettings): void => {
        if (!this._isActive) {
          return;
        }
        if (!settings.isConfigured) {
          this.setState({ status: 'notConfigured' });
          return;
        }

        const value: IEdControlsForm = toEdControlsForm(settings, projectId > 0);
        this.setState({ status: 'ready', hasSavedLink: settings.isTicketSource && !!settings.ticketProjectId });
        this.props.onLoaded(value);
        if (value.isTicketSource) {
          this._loadProjects();
          this._loadTags(value.ticketProjectId);
        }
        if (projectId > 0) {
          value.tagRows.forEach((row: ITagRow) => this._loadRowFolders(row));
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({
            status: 'failed',
            loadError: getErrorMessage(error, 'The ED Controls settings could not be loaded.')
          });
        }
      }
    );
  }

  private _loadProjects(): void {
    this.setState({ isLoadingProjects: true, projectsError: undefined });
    this.props.integrationService.getEdControlsProjects().then(
      (projects: ILookupOption[]): void => {
        if (this._isActive) {
          this.setState({ projects, isLoadingProjects: false });
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({
            isLoadingProjects: false,
            projectsError: getErrorMessage(error, 'The ED Controls projects could not be loaded.')
          });
        }
      }
    );
  }

  private _loadTags(ticketProjectId: string): void {
    if (!ticketProjectId || this.props.projectId <= 0) {
      this.setState({ tags: [], tagsError: undefined });
      return;
    }

    this.setState({ isLoadingTags: true, tagsError: undefined });
    this.props.integrationService.getEdControlsTags(ticketProjectId).then(
      (tags: string[]): void => {
        // A later project change may have overtaken this answer.
        if (this._isActive && this.props.value?.ticketProjectId === ticketProjectId) {
          this.setState({ tags, isLoadingTags: false });
        }
      },
      (error: unknown): void => {
        if (this._isActive && this.props.value?.ticketProjectId === ticketProjectId) {
          this.setState({
            tags: [],
            isLoadingTags: false,
            tagsError: getErrorMessage(error, 'The ED Controls tags could not be loaded.')
          });
        }
      }
    );
  }

  /** Every level a stored row has a selection above - they load in parallel, each needs only its ancestors. */
  private _loadRowFolders(row: ITagRow): void {
    FOLDER_LEVELS.forEach((level: FolderLevel, index: number): void => {
      if (index === 0 || row.folder[FOLDER_LEVELS[index - 1]]) {
        this._loadFolderLevel(row.key, level, row.folder);
      }
    });
  }

  private _loadFolderLevel(rowKey: string, level: FolderLevel, selection: IFolderSelection): void {
    this.setState((state: IEdControlsSectionState) => ({
      rowLoadingLevels: { ...state.rowLoadingLevels, [rowKey]: (state.rowLoadingLevels[rowKey] || []).concat([level]) }
    }));

    const settle = (options: ILookupOption[]): void => {
      if (!this._isActive) {
        return;
      }
      this.setState((state: IEdControlsSectionState) => {
        const current: IFolderLevelOptions = state.rowFolderOptions[rowKey] || createEmptyFolderLevelOptions();
        return {
          rowFolderOptions: { ...state.rowFolderOptions, [rowKey]: { ...current, [level]: options } },
          rowLoadingLevels: {
            ...state.rowLoadingLevels,
            [rowKey]: (state.rowLoadingLevels[rowKey] || []).filter((pending: FolderLevel) => pending !== level)
          }
        };
      });
    };

    // A failed level simply has no options, like the MOM cascade.
    this.props.lookupService.getFolderLevel(this.props.projectId, level, selection).then(settle, () => settle([]));
  }

  private _update(changes: Partial<IEdControlsForm>): void {
    if (this.props.value) {
      this.props.onChange({ ...this.props.value, ...changes });
    }
  }

  private _updateRow(rowKey: string, changes: Partial<ITagRow>): void {
    const value: IEdControlsForm | undefined = this.props.value;
    if (!value) {
      return;
    }

    this._update({
      tagRows: value.tagRows.map((row: ITagRow) => (row.key === rowKey ? { ...row, ...changes } : row))
    });
    if (this.state.addRowErrors[rowKey]) {
      this.setState((state: IEdControlsSectionState) => {
        const addRowErrors: { [key: string]: ITagRowErrors } = { ...state.addRowErrors };
        delete addRowErrors[rowKey];
        return { addRowErrors };
      });
    }
  }

  private _onRetry = (): void => {
    this._loadSettings();
  };

  private _onToggle = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const isTicketSource: boolean = event.target.checked;

    this._update({ isTicketSource });
    if (isTicketSource && this.state.projects.length === 0 && !this.state.isLoadingProjects) {
      this._loadProjects();
    }
  };

  private _onProjectChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const ticketProjectId: string = event.target.value;

    this._update({ ticketProjectId });
    // `_loadTags` reads the new value from props, so it runs once the form has re-rendered.
    this.setState({ tags: [] }, (): void => this._loadTags(ticketProjectId));
  };

  private _onRowFolderChange(rowKey: string, selection: IFolderSelection, changedLevel: FolderLevel): void {
    this._updateRow(rowKey, { folder: selection });
    this.setState((state: IEdControlsSectionState) => ({
      rowFolderOptions: {
        ...state.rowFolderOptions,
        [rowKey]: clearDeeperLevelOptions(state.rowFolderOptions[rowKey] || createEmptyFolderLevelOptions(), changedLevel)
      }
    }));

    const childLevel: FolderLevel | undefined = FOLDER_LEVELS[FOLDER_LEVELS.indexOf(changedLevel) + 1];
    if (childLevel && selection[changedLevel]) {
      this._loadFolderLevel(rowKey, childLevel, selection);
    }
  }

  private _onAddRow = (): void => {
    const value: IEdControlsForm | undefined = this.props.value;
    if (!value) {
      return;
    }

    const addRowErrors: { [key: string]: ITagRowErrors } = validateTagRows(value.tagRows, true);
    this.setState({ addRowErrors });
    if (Object.keys(addRowErrors).length > 0) {
      return;
    }

    const row: ITagRow = createTagRow();
    this._update({ tagRows: value.tagRows.concat([row]) });
    this._loadFolderLevel(row.key, 'folderId', row.folder);
  };

  private _onRemoveRow(rowKey: string): void {
    const value: IEdControlsForm | undefined = this.props.value;
    if (!value || value.tagRows.length <= 1) {
      return;
    }

    this._update({ tagRows: value.tagRows.filter((row: ITagRow) => row.key !== rowKey) });
    this.setState((state: IEdControlsSectionState) => {
      const rowFolderOptions: { [key: string]: IFolderLevelOptions } = { ...state.rowFolderOptions };
      const addRowErrors: { [key: string]: ITagRowErrors } = { ...state.addRowErrors };
      delete rowFolderOptions[rowKey];
      delete addRowErrors[rowKey];
      return { rowFolderOptions, addRowErrors };
    });
  }

  /** The options, plus the stored value when the list does not have it (kept selectable rather than silently lost). */
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
