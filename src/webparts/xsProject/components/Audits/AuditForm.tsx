import * as React from 'react';

import styles from './Audits.module.scss';
import type { IAuditFormProps, IAuditFormState } from './IAuditFormProps';
import FolderCascade from '../FolderCascade/FolderCascade';
import { IFolderSelection, ILookupOption } from '../../../../models/Building';
import { getErrorMessage } from '../../../../models/ApiError';
import {
  IAuditForm,
  IAuditFormErrors,
  MAX_AUDIT_NAME_LENGTH,
  validateAuditForm
} from '../../../../models/EdControlsAudit';
import {
  FOLDER_LEVELS,
  FolderLevel,
  clearDeeperLevelOptions,
  createEmptyFolderLevelOptions
} from '../../../../models/FolderTree';

/**
 * The add-audit form, also used to edit a duplicate draft - the reference's `#edAuditModal`.
 *
 * Name, ED Controls template, supplier (its contact list follows it), due date (dd/mm/yyyy)
 * and the folder the completed report is filed into. Filling it in creates nothing: the
 * parent decides what Save does (create the audit, or update the draft).
 */
export default class AuditForm extends React.Component<IAuditFormProps, IAuditFormState> {
  private _isActive: boolean = false;

  public constructor(props: IAuditFormProps) {
    super(props);

    this.state = {
      form: { ...props.initial, folder: { ...props.initial.folder } },
      errors: {},
      templates: [],
      isLoadingTemplates: false,
      contacts: [],
      isLoadingContacts: false,
      folderOptions: createEmptyFolderLevelOptions(),
      loadingFolderLevels: []
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._loadTemplates();
    if (this.state.form.supplierId) {
      this._loadContacts(this.state.form.supplierId);
    }
    // Every level a draft has a selection above, in parallel - each needs only its ancestors.
    FOLDER_LEVELS.forEach((level: FolderLevel, index: number): void => {
      if (index === 0 || this.state.form.folder[FOLDER_LEVELS[index - 1]]) {
        this._loadFolderLevel(level, this.state.form.folder);
      }
    });
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IAuditFormProps> {
    const { title, isSaving, error, suppliers, isGrandChildFolderEnabled } = this.props;
    const { form, errors, templates, isLoadingTemplates, templatesError, contacts, isLoadingContacts } = this.state;

    return (
      <form className={styles.panel} onSubmit={this._onSubmit} noValidate aria-label={title}>
        <h3 className={styles.subTitle}>{title}</h3>

        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.label} htmlFor="audit-name">
              Audit name <span className={styles.required} aria-hidden="true">*</span>
            </label>
            <input
              id="audit-name"
              className={`${styles.input} ${errors.auditName ? styles.inputInvalid : ''}`}
              type="text"
              maxLength={MAX_AUDIT_NAME_LENGTH}
              value={form.auditName}
              disabled={isSaving}
              onChange={(event: React.ChangeEvent<HTMLInputElement>) => this._set({ auditName: event.target.value })}
            />
            {errors.auditName && <span className={styles.fieldError} role="alert">{errors.auditName}</span>}
          </div>

          <div className={styles.field}>
            <label className={styles.label} htmlFor="audit-templateId">
              Ed Controls template <span className={styles.required} aria-hidden="true">*</span>
            </label>
            <select
              id="audit-templateId"
              className={`${styles.input} ${errors.templateId ? styles.inputInvalid : ''}`}
              value={form.templateId}
              disabled={isSaving || isLoadingTemplates}
              onChange={(event: React.ChangeEvent<HTMLSelectElement>) => this._set({ templateId: event.target.value })}
            >
              <option value="">{isLoadingTemplates ? 'Loading…' : 'Select template'}</option>
              {AuditForm._withSelected(templates, form.templateId, form.templateName).map(AuditForm._renderOption)}
            </select>
            {errors.templateId && <span className={styles.fieldError} role="alert">{errors.templateId}</span>}
            {templatesError && <span className={styles.fieldError} role="alert">{templatesError}</span>}
            {!isLoadingTemplates && !templatesError && templates.length === 0 && (
              <span className={styles.emptyHint}>ED Controls returned no templates for this project.</span>
            )}
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.label} htmlFor="audit-supplierId">
              Supplier <span className={styles.required} aria-hidden="true">*</span>
            </label>
            <select
              id="audit-supplierId"
              className={`${styles.input} ${errors.supplierId ? styles.inputInvalid : ''}`}
              value={form.supplierId}
              disabled={isSaving}
              onChange={this._onSupplierChange}
            >
              <option value="">Select supplier</option>
              {suppliers.map(AuditForm._renderOption)}
            </select>
            {errors.supplierId && <span className={styles.fieldError} role="alert">{errors.supplierId}</span>}
          </div>

          <div className={styles.field}>
            <label className={styles.label} htmlFor="audit-contactId">Contact</label>
            <select
              id="audit-contactId"
              className={styles.input}
              value={form.contactId}
              disabled={isSaving || !form.supplierId || isLoadingContacts}
              onChange={(event: React.ChangeEvent<HTMLSelectElement>) => this._set({ contactId: event.target.value })}
            >
              <option value="">
                {!form.supplierId ? 'Select a supplier first' : isLoadingContacts ? 'Loading…' : 'Select contact'}
              </option>
              {AuditForm._withSelected(contacts, form.contactId, form.contactName).map(AuditForm._renderOption)}
            </select>
          </div>
        </div>

        <div className={styles.row}>
          <div className={styles.field}>
            <label className={styles.label} htmlFor="audit-dueDate">
              Due date <span className={styles.required} aria-hidden="true">*</span>
            </label>
            <input
              id="audit-dueDate"
              className={`${styles.input} ${errors.dueDate ? styles.inputInvalid : ''}`}
              type="text"
              placeholder="dd/mm/yyyy"
              autoComplete="off"
              value={form.dueDate}
              disabled={isSaving}
              onChange={(event: React.ChangeEvent<HTMLInputElement>) => this._set({ dueDate: event.target.value })}
            />
            {errors.dueDate && <span className={styles.fieldError} role="alert">{errors.dueDate}</span>}
          </div>
        </div>

        <FolderCascade
          idPrefix="audit"
          options={this.state.folderOptions}
          suppliers={[]}
          hideSupplier
          value={form.folder}
          isGrandChildFolderEnabled={isGrandChildFolderEnabled}
          loadingLevels={this.state.loadingFolderLevels}
          isDisabled={isSaving}
          onChange={this._onFolderChange}
          onSupplierChange={() => undefined}
        />
        {errors.folderId && <span className={styles.fieldError} role="alert">{errors.folderId}</span>}

        {error && <div className={styles.formError} role="alert">{error}</div>}

        <div className={styles.actions}>
          <button type="submit" className={styles.primaryButton} disabled={isSaving}>
            {isSaving ? 'Saving…' : 'Save'}
          </button>
          <button type="button" className={styles.secondaryButton} disabled={isSaving} onClick={this.props.onCancel}>
            Cancel
          </button>
        </div>
      </form>
    );
  }

  private _set(changes: Partial<IAuditForm>): void {
    this.setState((state: IAuditFormState) => ({ form: { ...state.form, ...changes } }));
  }

  private _loadTemplates(): void {
    this.setState({ isLoadingTemplates: true, templatesError: undefined });
    this.props.auditService.getTemplates(this.props.ticketProjectId).then(
      (templates: ILookupOption[]): void => {
        if (this._isActive) {
          this.setState({ templates, isLoadingTemplates: false });
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({
            isLoadingTemplates: false,
            templatesError: getErrorMessage(error, 'The ED Controls templates could not be loaded.')
          });
        }
      }
    );
  }

  private _loadContacts(supplierId: string): void {
    this.setState({ contacts: [], isLoadingContacts: true });
    this.props.auditService.getSupplierContacts(Number(supplierId)).then(
      (contacts: ILookupOption[]): void => {
        if (this._isActive && this.state.form.supplierId === supplierId) {
          this.setState({ contacts, isLoadingContacts: false });
        }
      },
      (): void => {
        // The contact is optional: no list simply leaves it unchosen.
        if (this._isActive && this.state.form.supplierId === supplierId) {
          this.setState({ contacts: [], isLoadingContacts: false });
        }
      }
    );
  }

  private _loadFolderLevel(level: FolderLevel, selection: IFolderSelection): void {
    this.setState((state: IAuditFormState) => ({ loadingFolderLevels: state.loadingFolderLevels.concat([level]) }));

    const settle = (options: ILookupOption[]): void => {
      if (this._isActive) {
        this.setState((state: IAuditFormState) => ({
          folderOptions: { ...state.folderOptions, [level]: options },
          loadingFolderLevels: state.loadingFolderLevels.filter((pending: FolderLevel) => pending !== level)
        }));
      }
    };

    this.props.lookupService.getFolderLevel(this.props.projectId, level, selection).then(settle, () => settle([]));
  }

  private _onSupplierChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const supplierId: string = event.target.value;

    this._set({ supplierId, contactId: '', contactName: '' });
    if (supplierId) {
      this._loadContacts(supplierId);
    } else {
      this.setState({ contacts: [] });
    }
  };

  private _onFolderChange = (folder: IFolderSelection, changedLevel: FolderLevel): void => {
    this.setState((state: IAuditFormState) => ({
      form: { ...state.form, folder },
      folderOptions: clearDeeperLevelOptions(state.folderOptions, changedLevel)
    }));

    const childLevel: FolderLevel | undefined = FOLDER_LEVELS[FOLDER_LEVELS.indexOf(changedLevel) + 1];
    if (childLevel && folder[changedLevel]) {
      this._loadFolderLevel(childLevel, folder);
    }
  };

  private _onSubmit = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();

    const errors: IAuditFormErrors = validateAuditForm(this.state.form);
    this.setState({ errors });
    if (Object.keys(errors).length > 0) {
      return;
    }

    this.props.onSubmit(this._withDisplayNames(this.state.form));
  };

  /** The names a draft row shows, read off the dropdowns as the reference does; a name already held is kept. */
  private _withDisplayNames(form: IAuditForm): IAuditForm {
    const { templates, contacts, folderOptions } = this.state;
    const name = (options: ILookupOption[], id: string, fallback: string): string => {
      const match: ILookupOption | undefined = options.filter((option: ILookupOption) => option.id === id)[0];
      return match ? match.name : fallback;
    };
    const segments: string[] = [];
    FOLDER_LEVELS.forEach((level: FolderLevel): void => {
      const id: string = form.folder[level];
      const label: string = id ? name(folderOptions[level], id, '') : '';
      if (label) {
        segments.push(label);
      }
    });

    return {
      ...form,
      auditName: form.auditName.trim(),
      templateName: name(templates, form.templateId, form.templateName),
      supplierName: name(this.props.suppliers, form.supplierId, form.supplierName),
      contactName: form.contactId ? name(contacts, form.contactId, form.contactName) : '',
      folderPath: segments.length > 0 ? segments.join(' / ') : form.folderPath
    };
  }

  /** The options, plus the held value when the list does not have it (kept visible rather than silently lost). */
  private static _withSelected(options: ILookupOption[], id: string, name: string): ILookupOption[] {
    if (!id || options.some((option: ILookupOption) => option.id === id)) {
      return options;
    }
    return options.concat([{ id, name: name || id }]);
  }

  private static _renderOption(option: ILookupOption): React.ReactElement {
    return <option key={option.id} value={option.id}>{option.name}</option>;
  }
}
