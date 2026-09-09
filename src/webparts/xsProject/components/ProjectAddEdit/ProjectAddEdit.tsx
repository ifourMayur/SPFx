import * as React from 'react';

import styles from './ProjectAddEdit.module.scss';
import type { IProjectAddEditProps } from './IProjectAddEditProps';
import type { IProjectAddEditState } from './IProjectAddEditState';
import FolderCascade from '../FolderCascade/FolderCascade';
import MultiSelect from '../MultiSelect/MultiSelect';
import {
  IBuildingForm,
  IBuildingFormErrors,
  IFolderSelection,
  ILookupOption,
  createEmptyBuildingForm
} from '../../../../models/Building';
import { hasErrors, validateBuildingForm } from '../../../../models/BuildingValidation';
import {
  FOLDER_LEVELS,
  FolderLevel,
  IFolderLevelOptions,
  clearDeeperLevelOptions,
  createEmptyFolderLevelOptions
} from '../../../../models/FolderTree';
import { getErrorMessage } from '../../../../models/ApiError';

/**
 * Keys of `IBuildingForm` whose value is a `string`, so one handler can serve them all.
 *
 * The `-?` strips the optional modifier while mapping: without it the optional properties
 * would carry `undefined` into the result, and the union would no longer be assignable to
 * `keyof IBuildingForm`.
 */
type StringFormField = {
  [K in keyof IBuildingForm]-?: IBuildingForm[K] extends string ? K : never;
}[keyof IBuildingForm];

/** Keys of `IBuildingForm` whose value is a `boolean` - the toggles and checkboxes. */
type BooleanFormField = {
  [K in keyof IBuildingForm]-?: IBuildingForm[K] extends boolean ? K : never;
}[keyof IBuildingForm];

/** Keys of `IBuildingForm` holding a multi-select's chosen ids. */
type MultiSelectFormField = 'userIds' | 'spatialBreakdownIds';

/**
 * Image types the reference's file input accepts.
 *
 * Used both for the `accept` hint and to check the chosen file, because `accept` only
 * filters the file picker's default view - a user can still pick anything.
 */
const ACCEPTED_IMAGE_TYPES: string[] = ['image/jpg', 'image/png', 'image/jpeg'];

/** Shown when the chosen file is not an image, or cannot be read. */
const IMAGE_MESSAGES = {
  wrongType: 'The project image must be a JPG or PNG file.',
  unreadable: 'That image could not be read. Please choose it again.'
};

/**
 * Add / edit form for a project, ported from the reference application's
 * `Views/Building/AddEdit.cshtml`.
 *
 * One component serves both modes, decided by `projectId`: absent means add (Reset is
 * offered, Project Template is editable), present means edit (Reset is hidden, Project
 * Template is locked - `@if (Model.Id > 0)` in the reference).
 *
 * The ED Controls / Audits and KYP Planning sections are deliberately not ported: they
 * are third-party integration features, out of scope for the core project form. The
 * SharePoint Site URL field is out of scope for the same reason - the reference only shows
 * it when server-side SharePoint credentials are configured, which the browser cannot see.
 *
 * Every dropdown is bound to the same Web API endpoint the reference binds it to, through
 * `ILookupService` - see that interface for the endpoint-by-dropdown table. They are
 * fetched once on mount, in parallel, and the form is not shown until they arrive.
 *
 * Saving goes through `onSave`, the single seam the caller uses to reach the API - today
 * `IBuildingService.saveProject`, which posts to `api/Building`. This component stays
 * unaware of that: it awaits the promise, shows its saving state while it is pending, and
 * renders a rejection as the form-level error with the user's input left intact.
 */
export default class ProjectAddEdit extends React.Component<IProjectAddEditProps, IProjectAddEditState> {
  /** Guards against completing a save after the component has been unmounted. */
  private _isActive: boolean = false;

  /** Form to restore when Reset is pressed. */
  private readonly _initialForm: IBuildingForm;

  public constructor(props: IProjectAddEditProps) {
    super(props);

    this._initialForm = ProjectAddEdit._buildInitialForm(props);
    this.state = {
      form: this._initialForm,
      errors: {},
      isLoading: true,
      lookups: {
        projectTemplates: [],
        users: [],
        spatialBreakdowns: [],
        tenderTemplates: [],
        suppliers: []
      },
      folderOptions: createEmptyFolderLevelOptions(),
      loadingFolderLevels: [],
      isSaving: false
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._loadLookups();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IProjectAddEditProps> {
    if (!this.props.canSave) {
      return (
        <section className={styles.projectAddEdit}>
          <div className={styles.accessDenied} role="alert">
            <h2 className={styles.title}>Project</h2>
            <p>Your account does not have permission to add or edit projects.</p>
            <button type="button" className={styles.secondaryButton} onClick={this.props.onCancel}>
              Back
            </button>
          </div>
        </section>
      );
    }

    const { isLoading, isSaving, loadError, saveError } = this.state;
    // Drives the three edit-mode differences the reference has: the "details not in this
    // session" notice, the locked Project Template, and the hidden Reset button.
    const isEditMode: boolean = ProjectAddEdit._isEditMode(this.props);

    return (
      <section className={styles.projectAddEdit}>
        <div className={styles.header}>
          <button type="button" className={styles.secondaryButton} onClick={this.props.onCancel}>
            Back
          </button>
        </div>

        <h2 className={styles.title}>Project</h2>
        <hr className={styles.titleRule} />

        {/* The form is withheld until its dropdowns exist: showing empty selects the user
            could submit would be worse than making them wait. */}
        {isLoading && (
          <div className={styles.panel}>
            <div className={styles.loadingState} role="status" aria-live="polite">
              <span className={styles.spinner} aria-hidden="true" />
              <span>Loading&hellip;</span>
            </div>
          </div>
        )}

        {!isLoading && loadError && (
          <div className={styles.panel}>
            <div className={styles.formError} role="alert">{loadError}</div>
            <div className={styles.actions}>
              <button type="button" className={styles.primaryButton} onClick={this._onRetryLoad}>
                Try again
              </button>
              <button type="button" className={styles.secondaryButton} onClick={this.props.onCancel}>
                Back
              </button>
            </div>
          </div>
        )}

        {isEditMode && !this.props.initialForm && (
          <div className={styles.notice} role="status">
            This project&apos;s saved details are not available in this session, so only its name
            has been filled in. Editing and saving will update the project name.
          </div>
        )}

        {/* noValidate: the browser's own bubbles would pre-empt the reference's messages. */}
        {!isLoading && !loadError && (
        <form className={styles.panel} onSubmit={this._onSubmit} noValidate>
          {this._renderImageUpload()}
          {this._renderCoreFields()}
          {this._renderSettingsAndTender()}
          {this._renderMinutesOfMeeting()}

          {saveError && <div className={styles.formError} role="alert">{saveError}</div>}

          <div className={styles.actions}>
            <button type="submit" className={styles.primaryButton} disabled={isSaving}>
              {isSaving ? 'Saving…' : 'Save'}
            </button>
            {/* The reference offers Reset in add mode only (`@if (Model.Id == 0)`). */}
            {!isEditMode && (
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={isSaving}
                onClick={this._onReset}
              >
                Reset
              </button>
            )}
          </div>
        </form>
        )}
      </section>
    );
  }

  /**
   * Loads every dropdown's options.
   *
   * All five run in parallel: they are independent endpoints, and doing them in sequence
   * would make the form wait for the sum of five round trips instead of the slowest one.
   * A single failure fails the whole load, because a form with some dropdowns silently
   * empty invites saving a half-filled project.
   *
   * The folder cascade is loaded separately, by `_loadFolderLevel`, since its endpoints
   * need a project id.
   */
  private _loadLookups(): void {
    const { lookupService } = this.props;

    this.setState({ isLoading: true, loadError: undefined });

    Promise.all([
      lookupService.getProjectTemplates(),
      lookupService.getUsers(),
      lookupService.getSpatialBreakdowns(),
      lookupService.getTenderTemplates(),
      lookupService.getSuppliers()
    ]).then(
      (results: ILookupOption[][]): void => {
        if (!this._isActive) {
          return;
        }

        this.setState(
          {
            isLoading: false,
            lookups: {
              projectTemplates: results[0],
              users: results[1],
              spatialBreakdowns: results[2],
              tenderTemplates: results[3],
              suppliers: results[4]
            }
          },
          (): void => {
            // Folders only exist once the project does, so this is a no-op while adding.
            if (this.state.form.isAddMomPdf) {
              this._loadFolderLevel('folderId');
            }
          }
        );
      },
      (error: unknown): void => {
        if (!this._isActive) {
          return;
        }

        this.setState({
          isLoading: false,
          loadError: getErrorMessage(error, 'The form data could not be loaded. Please try again.')
        });
      }
    );
  }

  /**
   * Fetches one level of the folder cascade for the project being edited.
   *
   * Every folder endpoint is scoped to a project, so in add mode (`id === 0`) there is
   * nothing to ask for and the level is left empty - the same gap the reference has, where
   * `projectId: $('#Id').val()` is `0` while adding.
   *
   * A failure here leaves the level empty rather than failing the whole form: the cascade
   * is optional, and MOM has no validation.
   */
  private _loadFolderLevel(level: FolderLevel): void {
    const { lookupService } = this.props;
    const { form } = this.state;
    const projectId: number = form.id;

    if (projectId <= 0) {
      return;
    }

    // Each level's endpoint takes the ids of every level above it, so the deeper the
    // level the longer the parameter list. The ids are strings in the form (a `<select>`
    // value always is) and integers on the wire.
    const selection: IFolderSelection = form.momFolder;
    const folderId: number = Number(selection.folderId);
    const subFolderId: number = Number(selection.subFolderId);
    let request: Promise<ILookupOption[]>;

    switch (level) {
      case 'folderId':
        request = lookupService.getProjectFolders(projectId);
        break;
      case 'subFolderId':
        request = lookupService.getProjectSubFolders(projectId, folderId);
        break;
      case 'subSubFolderId':
        request = lookupService.getProjectSubSubFolders(projectId, folderId, subFolderId);
        break;
      default:
        request = lookupService.getProjectSubSubSubFolders(
          projectId,
          folderId,
          subFolderId,
          Number(selection.subSubFolderId)
        );
        break;
    }

    this.setState((state: IProjectAddEditState) => ({
      loadingFolderLevels: state.loadingFolderLevels.concat([level])
    }));

    // Both outcomes go to the same handler: a failed level simply has no options, which
    // clears its loading state and leaves the dropdown empty rather than failing the form.
    request.then(
      (options: ILookupOption[]): void => this._onFolderLevelSettled(level, options),
      (): void => this._onFolderLevelSettled(level, [])
    );
  }

  private _onFolderLevelSettled(level: FolderLevel, options: ILookupOption[]): void {
    if (!this._isActive) {
      return;
    }

    this.setState((state: IProjectAddEditState) => {
      const folderOptions: IFolderLevelOptions = { ...state.folderOptions };
      folderOptions[level] = options;

      return {
        folderOptions,
        loadingFolderLevels: state.loadingFolderLevels.filter(
          (pending: FolderLevel) => pending !== level
        )
      };
    });
  }

  private _onRetryLoad = (): void => {
    this._loadLookups();
  };

  private _renderImageUpload(): React.ReactElement {
    const { imageDataUrl } = this.state.form;

    return (
      <div className={styles.imageRow}>
        {/* A label styled as a button, so choosing a file needs no click-forwarding script
            (the reference clicks the hidden input from the button's onclick). */}
        <label className={styles.imageButton}>
          {imageDataUrl
            ? <img className={styles.imagePreview} src={imageDataUrl} alt="Project image preview" />
            : <span className={styles.imagePlaceholder} aria-hidden="true">&#128247;</span>}
          <span className={styles.imageCaption}>Upload Image</span>
          <input
            type="file"
            className={styles.visuallyHidden}
            accept={ACCEPTED_IMAGE_TYPES.join(',')}
            disabled={this.state.isSaving}
            onChange={this._onImageChange}
          />
        </label>
      </div>
    );
  }

  private _renderCoreFields(): React.ReactElement {
    const { features } = this.props;
    const { form } = this.state;
    const isEditMode: boolean = ProjectAddEdit._isEditMode(this.props);
    // McD hides Spatial Breakdown, and is also the tenant grouping that does not require
    // a postcode - both are `domainGroup != McD` gates in the reference.
    const isMcD: boolean = features.domainGroup === 'mcD';

    return (
      <>
        <div className={styles.row}>
          {this._renderTextField('buildingName', 'Project Name', 'Enter Project Name', { isRequired: true })}
          {this._renderTextField('buldingId', 'Project ID', 'Enter Project ID')}
        </div>

        <div className={styles.row}>
          {this._renderTextField('postcode', 'Postcode', 'Enter Postcode', { isRequired: !isMcD })}
          {this._renderTextField('address', 'Address', 'Enter Address', { isRequired: true })}
        </div>

        <div className={styles.row}>
          {this._renderTextField('plaats', 'City', 'Enter City')}
          {this._renderTextField('contactpersoon', 'Contact person', 'Enter Contact person')}
        </div>

        <div className={styles.row}>
          {this._renderTextField('functie', 'Function', 'Enter Function')}
          {this._renderTextField('email', 'Email', 'Enter E-mail', { type: 'email' })}
        </div>

        <div className={styles.row}>
          {this._renderTextField('telefoon', 'Phone Number', 'Enter Phone', { type: 'tel' })}

          <div className={styles.field}>
            <label className={styles.label} htmlFor="project-clientName">
              Client <span className={styles.required} aria-hidden="true">*</span>
            </label>
            {/* Read-only, as in the reference: the client is always the signed-in user.
                `clientId` travels in `IBuildingForm`, so unlike the Razor view there is no
                hidden field to carry it - nothing here is submitted as an HTML form. */}
            <input
              id="project-clientName"
              className={styles.input}
              type="text"
              value={form.clientName}
              disabled
            />
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.label} htmlFor="project-projectTemplateId">Project Template</label>
            <select
              id="project-projectTemplateId"
              className={styles.input}
              value={form.projectTemplateId}
              // Locked once the project exists: its folder structure is already built from
              // the template it was created with.
              disabled={isEditMode || this.state.isSaving}
              data-field="projectTemplateId"
              onChange={this._onFieldChange}
            >
              <option value="">Select Project Template</option>
              {this.state.lookups.projectTemplates.map(ProjectAddEdit._renderOption)}
            </select>
            {/* Says so explicitly: an empty dropdown and a broken one look identical. */}
            {this.state.lookups.projectTemplates.length === 0 && (
              <span className={styles.emptyHint}>
                The Web API returned no project templates for your account.
              </span>
            )}
          </div>

          {this._renderMultiSelect(
            'userIds',
            'User List',
            this.state.lookups.users,
            'The Web API returned no users for your account.'
          )}
        </div>

        {!isMcD && (
          <div className={styles.row}>
            {this._renderMultiSelect(
              'spatialBreakdownIds',
              'Spatial Breakdown',
              this.state.lookups.spatialBreakdowns,
              'The Web API returned no spatial breakdowns.'
            )}
          </div>
        )}
      </>
    );
  }

  private _renderSettingsAndTender(): React.ReactElement {
    const { features } = this.props;
    const { form, errors } = this.state;

    return (
      <div className={styles.row}>
        <div className={styles.field}>
          <div className={styles.sectionLabel}>Project Settings</div>
          <hr className={styles.sectionRule} />

          <div className={styles.toggleRow}>
            {this._renderToggle('isMaintenanceFolderAvailable', 'Enable Maintenance Folder')}
            {this._renderToggle('isEnableBimFolder', 'Enable BIM folder')}
            {/* TaskXs is the only domain group offered the ISO 19650 naming format. */}
            {features.domainGroup === 'taskXs' && this._renderToggle('isEnableIsoFormat', 'ISO 19650 Code')}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.sectionLabel}>Tender</div>
          <hr className={styles.sectionRule} />

          <div className={styles.toggleRow}>
            {this._renderToggle('isEnableTender', 'Tender')}
          </div>

          {/* Revealed by the toggle, exactly as `getTenderTemplate()` shows/hides it. */}
          {form.isEnableTender && (
            <div className={styles.tenderField}>
              <label className={styles.label} htmlFor="project-tenderTemplateId">
                Tender Templates <span className={styles.required} aria-hidden="true">*</span>
              </label>
              <select
                id="project-tenderTemplateId"
                className={`${styles.input} ${errors.tenderTemplateId ? styles.inputInvalid : ''}`}
                value={form.tenderTemplateId}
                disabled={this.state.isSaving}
                aria-invalid={!!errors.tenderTemplateId}
                data-field="tenderTemplateId"
                onChange={this._onFieldChange}
              >
                <option value="">Select tender template</option>
                {this.state.lookups.tenderTemplates.map(ProjectAddEdit._renderOption)}
              </select>
              {errors.tenderTemplateId && (
                <span className={styles.fieldError} role="alert">{errors.tenderTemplateId}</span>
              )}
              {this.state.lookups.tenderTemplates.length === 0 && (
                <span className={styles.emptyHint}>
                  The Web API returned no tender templates.
                </span>
              )}
            </div>
          )}
        </div>
      </div>
    );
  }

  private _renderMinutesOfMeeting(): React.ReactElement {
    const { form } = this.state;

    return (
      <div className={styles.momSection}>
        <div className={styles.sectionLabel}>Minutes of Meeting</div>
        <hr className={styles.sectionRule} />

        <div className={styles.toggleRow}>
          {this._renderToggle('isAddMomPdf', 'Save minutes in this folder')}
        </div>

        {/* Hidden until the toggle is on, matching `#formMomBinding`'s initial state. */}
        {form.isAddMomPdf && (
          <FolderCascade
            idPrefix="project-mom"
            options={this.state.folderOptions}
            suppliers={this.state.lookups.suppliers}
            value={form.momFolder}
            isGrandChildFolderEnabled={this.props.features.isGrandChildFolderEnabled}
            loadingLevels={this.state.loadingFolderLevels}
            isDisabled={this.state.isSaving}
            onChange={this._onMomFolderChange}
            onSupplierChange={this._onMomSupplierChange}
          />
        )}
      </div>
    );
  }

  private _renderTextField(
    field: StringFormField,
    label: string,
    placeholder: string,
    options?: { isRequired?: boolean; type?: string }
  ): React.ReactElement {
    const { errors, form, isSaving } = this.state;
    const message: string | undefined = errors[field as keyof IBuildingFormErrors];
    const inputId: string = `project-${field}`;

    return (
      <div className={styles.field}>
        <label className={styles.label} htmlFor={inputId}>
          {label}
          {options?.isRequired && <span className={styles.required} aria-hidden="true"> *</span>}
        </label>
        <input
          id={inputId}
          className={`${styles.input} ${message ? styles.inputInvalid : ''}`}
          type={options?.type || 'text'}
          value={form[field]}
          placeholder={placeholder}
          autoComplete="off"
          disabled={isSaving}
          aria-invalid={!!message}
          data-field={field}
          onChange={this._onFieldChange}
        />
        {message && <span className={styles.fieldError} role="alert">{message}</span>}
      </div>
    );
  }

  /**
   * A `ListBoxFor` field from the reference, rendered as a chip-based multi-select.
   *
   * The reference upgrades these to select2, so a bare `<select multiple>` would be a
   * downgrade - see `MultiSelect` for why.
   */
  private _renderMultiSelect(
    field: MultiSelectFormField,
    label: string,
    options: ILookupOption[],
    emptyText: string
  ): React.ReactElement {
    return (
      <div className={styles.field}>
        <MultiSelect
          inputId={`project-${field}`}
          label={label}
          options={options}
          selectedIds={this.state.form[field]}
          isDisabled={this.state.isSaving}
          placeholder={`Select ${label.toLowerCase()}`}
          emptyText={emptyText}
          onChange={field === 'userIds' ? this._onUserIdsChange : this._onSpatialBreakdownIdsChange}
        />
      </div>
    );
  }

  private _renderToggle(field: BooleanFormField, label: string): React.ReactElement {
    return (
      <label className={styles.toggleField} key={field}>
        <span className={styles.toggleLabel}>{label}</span>
        <span className={styles.toggle}>
          <input
            type="checkbox"
            checked={this.state.form[field]}
            disabled={this.state.isSaving}
            data-field={field}
            onChange={this._onToggleChange}
          />
          <span className={styles.toggleTrack} aria-hidden="true" />
        </span>
      </label>
    );
  }

  private static _renderOption(option: ILookupOption): React.ReactElement {
    return <option key={option.id} value={option.id}>{option.name}</option>;
  }

  /** True when the route carried a project id, i.e. an existing project is being edited. */
  private static _isEditMode(props: IProjectAddEditProps): boolean {
    return props.projectId !== undefined && props.projectId > 0;
  }

  /**
   * The form the screen opens with.
   *
   * Add mode starts blank with the signed-in user as the client. Edit mode prefers the
   * caller's `initialForm`, and otherwise falls back to a blank form carrying just the
   * project's id and name - see `initialForm` on the props for why that case exists.
   */
  private static _buildInitialForm(props: IProjectAddEditProps): IBuildingForm {
    const base: IBuildingForm = {
      ...createEmptyBuildingForm(),
      clientId: props.clientId,
      clientName: props.clientName
    };

    if (!ProjectAddEdit._isEditMode(props)) {
      return base;
    }

    if (props.initialForm) {
      return { ...props.initialForm };
    }

    return {
      ...base,
      id: props.projectId as number,
      buildingName: props.fallbackProjectName || ''
    };
  }

  // Handlers are assigned as properties so the `this` pointer is bound without a
  // per-render closure, matching the pattern `Login` / `Menu` / `ProjectList` use. A
  // single handler per input kind, keyed off `data-field`, replaces one handler per field.
  private _onFieldChange = (
    event: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>
  ): void => {
    this._setFormValue(event.target.dataset.field as StringFormField, event.target.value);
  };

  private _onToggleChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const field: BooleanFormField = event.target.dataset.field as BooleanFormField;
    const isChecked: boolean = event.target.checked;

    this._setFormValue(field, isChecked);

    // Switching MOM on reveals the cascade, so its top level is fetched then rather than
    // up front - exactly when `getMOMProjectFolderList` runs in the reference.
    if (field === 'isAddMomPdf' && isChecked && this.state.folderOptions.folderId.length === 0) {
      this._loadFolderLevel('folderId');
    }
  };

  private _onUserIdsChange = (userIds: string[]): void => {
    this._setFormValue('userIds', userIds);
  };

  private _onSpatialBreakdownIdsChange = (spatialBreakdownIds: string[]): void => {
    this._setFormValue('spatialBreakdownIds', spatialBreakdownIds);
  };

  /**
   * Stores the new selection, discards the options below the level that changed, then
   * fetches the next level down - the same sequence as `onChangeFolderDD` and its
   * siblings in the reference.
   */
  private _onMomFolderChange = (momFolder: IFolderSelection, changedLevel: FolderLevel): void => {
    this.setState(
      (state: IProjectAddEditState) => ({
        form: { ...state.form, momFolder },
        folderOptions: clearDeeperLevelOptions(state.folderOptions, changedLevel)
      }),
      (): void => {
        const childLevel: FolderLevel | undefined = ProjectAddEdit._childLevel(changedLevel);
        if (childLevel && momFolder[changedLevel]) {
          this._loadFolderLevel(childLevel);
        }
      }
    );
  };

  private _onMomSupplierChange = (supplierId: string): void => {
    this._setFormValue('momFolder', { ...this.state.form.momFolder, supplierId });
  };

  /** The level immediately below `level`, or undefined for the deepest one. */
  private static _childLevel(level: FolderLevel): FolderLevel | undefined {
    return FOLDER_LEVELS[FOLDER_LEVELS.indexOf(level) + 1];
  }

  /**
   * Loads the chosen file as a `data:` URL.
   *
   * A data URL rather than the `URL.createObjectURL` blob the reference's `loadImage`
   * uses, because the same string has to do two jobs: show the live preview, and carry the
   * Base64 bytes `POST api/Building` stores in `ImageBytes`. A blob URL is a handle to
   * memory in this tab, so it can do the first job but never the second.
   *
   * The reference silently ignores a file whose content type it does not recognise
   * (finding 14 of the process document); this says so and leaves the previous image
   * in place.
   */
  private _onImageChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const input: HTMLInputElement = event.target;
    const file: File | undefined = input.files ? input.files[0] : undefined;
    if (!file) {
      return;
    }

    // Cleared either way, so choosing the same file again always re-fires `change`.
    input.value = '';

    if (ACCEPTED_IMAGE_TYPES.indexOf((file.type || '').toLowerCase()) < 0) {
      this.setState({ saveError: IMAGE_MESSAGES.wrongType });
      return;
    }

    const reader: FileReader = new FileReader();

    reader.onload = (): void => {
      if (!this._isActive) {
        return;
      }

      this.setState((state: IProjectAddEditState) => ({
        form: { ...state.form, imageDataUrl: String(reader.result) },
        saveError: undefined
      }));
    };

    reader.onerror = (): void => {
      if (this._isActive) {
        this.setState({ saveError: IMAGE_MESSAGES.unreadable });
      }
    };

    reader.readAsDataURL(file);
  };

  /**
   * Restores the form to what it opened with.
   *
   * The reference calls `location.reload()`, which would discard the whole web part - and
   * the hosting SharePoint page with it - so the initial state is restored instead.
   */
  private _onReset = (): void => {
    this.setState({ form: { ...this._initialForm }, errors: {}, saveError: undefined });
  };

  private _onSubmit = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();

    const errors: IBuildingFormErrors = validateBuildingForm(this.state.form, this.props.features);
    if (hasErrors(errors)) {
      // Every message is shown at once; the reference stops at the first failure.
      this.setState({ errors, saveError: undefined });
      return;
    }

    this.setState({ errors: {}, saveError: undefined, isSaving: true });

    this.props.onSave(this.state.form).then(
      (): void => {
        // The caller navigates away on success, so this only matters if it chose not to.
        if (this._isActive) {
          this.setState({ isSaving: false });
        }
      },
      (error: unknown): void => {
        if (!this._isActive) {
          return;
        }

        this.setState({
          isSaving: false,
          saveError: error instanceof Error && error.message
            ? error.message
            : 'The project could not be saved. Please try again.'
        });
      }
    );
  };

  private _setFormValue<K extends keyof IBuildingForm>(field: K, value: IBuildingForm[K]): void {
    this.setState((state: IProjectAddEditState) => ({
      form: { ...state.form, [field]: value } as IBuildingForm
    }));
  }
}
