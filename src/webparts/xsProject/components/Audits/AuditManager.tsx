import * as React from 'react';

import styles from './Audits.module.scss';
import type { IAuditManagerProps, IAuditManagerState } from './IAuditManagerProps';
import AuditForm from './AuditForm';
import { ILookupOption } from '../../../../models/Building';
import { getErrorMessage } from '../../../../models/ApiError';
import {
  AUDIT_MESSAGES,
  AUDIT_STATUS,
  IAudit,
  IAuditDraft,
  IAuditForm,
  IAuditSaveResult,
  MAX_DUPLICATES,
  createAuditForm,
  createDuplicateDrafts,
  formatDueDate,
  parseDuplicateCount,
  toAuditSaveRequest
} from '../../../../models/EdControlsAudit';
import { IEdControlsSettings } from '../../../../models/ProjectIntegration';

/**
 * A project's configured ED Controls audits, ported from the reference's
 * `_AddEdControlAudit.cshtml` (its standalone Audits page).
 *
 * - The list is always read back from the API, so it reflects the database.
 * - **Add Audits** creates the audit in ED Controls (from a template, in the project's saved
 *   ED Controls project) and a placeholder document, on Save only.
 * - **Duplicate** asks how many copies (1-20), then shows draft rows named "<name> - n",
 *   each editable (its name inline, everything else through the form); nothing is created
 *   until **Save duplicates**, and rows that fail stay on screen with their reason.
 * - **Remove** permanently deletes the audit's document from the project after a
 *   confirmation; ED Controls itself keeps the audit (it offers no delete).
 */
export default class AuditManager extends React.Component<IAuditManagerProps, IAuditManagerState> {
  private _isActive: boolean = false;

  public constructor(props: IAuditManagerProps) {
    super(props);

    this.state = {
      status: 'loading',
      isConfigured: false,
      ticketProjectId: '',
      audits: [],
      isLoadingAudits: false,
      suppliers: [],
      view: 'list',
      isSaving: false,
      countText: '1',
      drafts: [],
      draftNameErrors: {}
    };
  }

  public componentDidMount(): void {
    this._isActive = true;
    this._load();
  }

  public componentWillUnmount(): void {
    this._isActive = false;
  }

  public render(): React.ReactElement<IAuditManagerProps> {
    const { status, loadError, view } = this.state;

    if (status === 'loading') {
      return (
        <div className={styles.loadingState} role="status" aria-live="polite">
          <span className={styles.spinner} aria-hidden="true" />
          <span>Loading audits&hellip;</span>
        </div>
      );
    }

    if (status === 'failed') {
      return (
        <div className={styles.formError} role="alert">
          {loadError}{' '}
          <button type="button" className={styles.secondaryButton} onClick={this._onRetry}>Try again</button>
        </div>
      );
    }

    if (!this.state.isConfigured) {
      return (
        <div className={styles.notice} role="status">
          ED Controls is not set up for your organisation. Ask your administrator to add the ED Controls
          credentials in Settings.
        </div>
      );
    }

    switch (view) {
      case 'add':
        return this._renderForm('Add Audits', createAuditForm());
      case 'editDraft':
        return this._renderForm('Edit duplicate', this._editingDraft()?.form || createAuditForm());
      case 'count':
        return this._renderCount();
      case 'drafts':
        return this._renderDrafts();
      default:
        return this._renderList();
    }
  }

  private _renderList(): React.ReactElement {
    const { ticketProjectId, audits, isLoadingAudits, auditsError, notice, actionError } = this.state;

    return (
      <>
        <div className={styles.toolbar}>
          <button
            type="button"
            className={styles.primaryButton}
            disabled={!ticketProjectId}
            onClick={this._onAdd}
          >
            Add Audits
          </button>
        </div>
        {!ticketProjectId && (
          <div className={styles.notice} role="status">
            Select an ED Controls project on the project form and save the project before adding audits.
          </div>
        )}
        {notice && <div className={styles.notice} role="status">{notice}</div>}
        {actionError && <div className={styles.formError} role="alert">{actionError}</div>}
        {auditsError && <div className={styles.formError} role="alert">{auditsError}</div>}

        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Audit name</th>
                <th>Ed Controls template</th>
                <th>Supplier</th>
                <th>Contact</th>
                <th>Due date</th>
                <th>Location</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {isLoadingAudits && (
                <tr><td colSpan={8}>Loading&hellip;</td></tr>
              )}
              {!isLoadingAudits && audits.length === 0 && (
                <tr><td colSpan={8}>No audits configured yet.</td></tr>
              )}
              {!isLoadingAudits && audits.map((audit: IAudit) => this._renderAudit(audit))}
            </tbody>
          </table>
        </div>
      </>
    );
  }

  private _renderAudit(audit: IAudit): React.ReactElement {
    const { isSaving, removingId } = this.state;
    const isCompleted: boolean = audit.status === AUDIT_STATUS.completed;

    return (
      <tr key={audit.id}>
        <td>{audit.auditName}</td>
        <td>{audit.edControlsTemplateName}</td>
        <td>{audit.supplierName}</td>
        <td>{audit.contactName}</td>
        <td>{formatDueDate(audit.dueDate)}</td>
        <td>{audit.folderPath}</td>
        <td>
          <span className={isCompleted ? styles.badgeCompleted : styles.badgeWaiting}>
            {isCompleted ? 'Completed' : 'Waiting'}
          </span>
        </td>
        <td>
          {removingId === audit.id ? (
            <div className={styles.confirm} role="alertdialog" aria-label="Remove audit">
              <p>
                Are you sure you want to remove this audit? Its document will be permanently deleted
                from the project, including any report already imported into it.
              </p>
              <div className={styles.actions}>
                <button type="button" className={styles.primaryButton} disabled={isSaving} onClick={this._onConfirmRemove}>
                  {isSaving ? 'Removing…' : 'Remove'}
                </button>
                <button type="button" className={styles.secondaryButton} disabled={isSaving} onClick={this._onCancelRemove}>
                  Cancel
                </button>
              </div>
            </div>
          ) : (
            <>
              <button type="button" className={styles.linkButton} disabled={isSaving} onClick={() => this._onDuplicate(audit)}>
                Duplicate
              </button>
              <span className={styles.separator} aria-hidden="true">|</span>
              <button type="button" className={styles.linkButton} disabled={isSaving} onClick={() => this._onRemove(audit.id)}>
                Delete
              </button>
            </>
          )}
        </td>
      </tr>
    );
  }

  private _renderForm(title: string, initial: IAuditForm): React.ReactElement {
    return (
      <AuditForm
        key={this.state.editingDraftKey || 'add'}
        title={title}
        projectId={this.props.projectId}
        ticketProjectId={this.state.ticketProjectId}
        initial={initial}
        suppliers={this.state.suppliers}
        isGrandChildFolderEnabled={this.props.isGrandChildFolderEnabled}
        auditService={this.props.auditService}
        lookupService={this.props.lookupService}
        isSaving={this.state.isSaving}
        error={this.state.actionError}
        onSubmit={this.state.view === 'editDraft' ? this._onDraftFormSubmit : this._onAddSubmit}
        onCancel={this.state.view === 'editDraft' ? this._onBackToDrafts : this._onBackToList}
      />
    );
  }

  private _renderCount(): React.ReactElement {
    const { countText, countError } = this.state;

    return (
      <form className={styles.panel} onSubmit={this._onConfirmCount} noValidate aria-label="Duplicate">
        <h3 className={styles.subTitle}>Duplicate</h3>
        <div className={styles.field}>
          <label className={styles.label} htmlFor="audit-duplicateCount">
            How many duplicates do you need? <span className={styles.required} aria-hidden="true">*</span>
          </label>
          <input
            id="audit-duplicateCount"
            className={`${styles.input} ${countError ? styles.inputInvalid : ''}`}
            type="number"
            min={1}
            max={MAX_DUPLICATES}
            step={1}
            value={countText}
            onChange={(event: React.ChangeEvent<HTMLInputElement>) => this.setState({ countText: event.target.value })}
          />
          {countError && <span className={styles.fieldError} role="alert">{countError}</span>}
        </div>
        <div className={styles.actions}>
          <button type="submit" className={styles.primaryButton}>Confirm</button>
          <button type="button" className={styles.secondaryButton} onClick={this._onBackToList}>Cancel</button>
        </div>
      </form>
    );
  }

  private _renderDrafts(): React.ReactElement {
    const { drafts, draftNameErrors, isSaving, actionError } = this.state;

    return (
      <div className={styles.panel}>
        <h3 className={styles.subTitle}>Save duplicates</h3>
        {actionError && <div className={styles.formError} role="alert">{actionError}</div>}
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Audit name</th>
                <th>Ed Controls template</th>
                <th>Supplier</th>
                <th>Contact</th>
                <th>Due date</th>
                <th>Location</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {drafts.map((draft: IAuditDraft) => (
                <tr key={draft.key}>
                  <td>
                    <input
                      className={`${styles.input} ${draftNameErrors[draft.key] ? styles.inputInvalid : ''}`}
                      type="text"
                      aria-label="Audit name"
                      value={draft.form.auditName}
                      disabled={isSaving}
                      onChange={(event: React.ChangeEvent<HTMLInputElement>) => this._onDraftNameChange(draft.key, event.target.value)}
                    />
                    {draftNameErrors[draft.key] && (
                      <span className={styles.fieldError} role="alert">{draftNameErrors[draft.key]}</span>
                    )}
                    {draft.error && <span className={styles.fieldError} role="alert">{draft.error}</span>}
                  </td>
                  <td>{draft.form.templateName}</td>
                  <td>{draft.form.supplierName}</td>
                  <td>{draft.form.contactName}</td>
                  <td>{draft.form.dueDate}</td>
                  <td>{draft.form.folderPath}</td>
                  <td>{draft.error ? <span className={styles.badgeWaiting} aria-label="Failed">&#10007;</span> : null}</td>
                  <td>
                    <button type="button" className={styles.linkButton} disabled={isSaving} onClick={() => this._onEditDraft(draft.key)}>
                      Edit
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className={styles.actions}>
          <button type="button" className={styles.primaryButton} disabled={isSaving || drafts.length === 0} onClick={this._onSaveDrafts}>
            {isSaving ? 'Saving…' : 'Save'}
          </button>
          <button type="button" className={styles.secondaryButton} disabled={isSaving} onClick={this._onBackToList}>
            Cancel
          </button>
        </div>
      </div>
    );
  }

  private _load(): void {
    const { integrationService, lookupService, projectId } = this.props;

    this.setState({ status: 'loading', loadError: undefined });
    Promise.all([integrationService.getEdControlsSettings(projectId), lookupService.getSuppliers()]).then(
      ([settings, suppliers]: [IEdControlsSettings, ILookupOption[]]): void => {
        if (!this._isActive) {
          return;
        }
        if (settings.projectName && this.props.onProjectName) {
          this.props.onProjectName(settings.projectName);
        }
        this.setState({
          status: 'ready',
          isConfigured: settings.isConfigured,
          // Audits are created in the project's saved ED Controls project, as on the reference's standalone page.
          ticketProjectId: settings.isTicketSource ? settings.ticketProjectId : '',
          suppliers
        });
        if (settings.isConfigured) {
          this._loadAudits();
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({ status: 'failed', loadError: getErrorMessage(error, 'The audits could not be loaded.') });
        }
      }
    );
  }

  private _loadAudits(): void {
    this.setState({ isLoadingAudits: true, auditsError: undefined });
    this.props.auditService.getAudits(this.props.projectId).then(
      (audits: IAudit[]): void => {
        if (this._isActive) {
          this.setState({ audits, isLoadingAudits: false });
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({
            audits: [],
            isLoadingAudits: false,
            auditsError: getErrorMessage(error, 'The audits could not be loaded.')
          });
        }
      }
    );
  }

  private _editingDraft(): IAuditDraft | undefined {
    return this.state.drafts.filter((draft: IAuditDraft) => draft.key === this.state.editingDraftKey)[0];
  }

  private _onRetry = (): void => {
    this._load();
  };

  private _onAdd = (): void => {
    this.setState({ view: 'add', notice: undefined, actionError: undefined, removingId: undefined });
  };

  private _onBackToList = (): void => {
    this.setState({ view: 'list', actionError: undefined, drafts: [], draftNameErrors: {}, editingDraftKey: undefined });
  };

  private _onBackToDrafts = (): void => {
    this.setState({ view: 'drafts', actionError: undefined, editingDraftKey: undefined });
  };

  private _onAddSubmit = (form: IAuditForm): void => {
    this.setState({ isSaving: true, actionError: undefined });
    this.props.auditService
      .createAudit(toAuditSaveRequest(form, this.props.projectId, this.state.ticketProjectId))
      .then(
        (message: string): void => {
          if (this._isActive) {
            this.setState({ isSaving: false, view: 'list', notice: message || 'Audit added successfully.' });
            this._loadAudits();
          }
        },
        (error: unknown): void => {
          if (this._isActive) {
            this.setState({ isSaving: false, actionError: getErrorMessage(error, 'The audit could not be added.') });
          }
        }
      );
  };

  private _onRemove(id: number): void {
    this.setState({ removingId: id, notice: undefined, actionError: undefined });
  }

  private _onCancelRemove = (): void => {
    this.setState({ removingId: undefined });
  };

  private _onConfirmRemove = (): void => {
    const id: number | undefined = this.state.removingId;
    if (id === undefined) {
      return;
    }

    this.setState({ isSaving: true, actionError: undefined });
    this.props.auditService.deleteAudit(id).then(
      (): void => {
        if (this._isActive) {
          this.setState({ isSaving: false, removingId: undefined, notice: 'Audit removed successfully.' });
          this._loadAudits();
        }
      },
      (error: unknown): void => {
        if (this._isActive) {
          this.setState({
            isSaving: false,
            removingId: undefined,
            actionError: getErrorMessage(error, 'The audit could not be removed.')
          });
        }
      }
    );
  };

  private _onDuplicate(audit: IAudit): void {
    this.setState({
      view: 'count',
      duplicateSource: audit,
      countText: '1',
      countError: undefined,
      notice: undefined,
      actionError: undefined,
      removingId: undefined
    });
  }

  private _onConfirmCount = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();

    const count: number | undefined = parseDuplicateCount(this.state.countText);
    const source: IAudit | undefined = this.state.duplicateSource;
    if (count === undefined || !source) {
      this.setState({ countError: AUDIT_MESSAGES.countInvalid });
      return;
    }

    this.setState({ view: 'drafts', drafts: createDuplicateDrafts(source, count), draftNameErrors: {}, countError: undefined });
  };

  private _onDraftNameChange(key: string, auditName: string): void {
    this.setState((state: IAuditManagerState) => ({
      drafts: state.drafts.map((draft: IAuditDraft) =>
        draft.key === key ? { ...draft, form: { ...draft.form, auditName } } : draft
      )
    }));
  }

  private _onEditDraft(key: string): void {
    this.setState({ view: 'editDraft', editingDraftKey: key, actionError: undefined });
  }

  private _onDraftFormSubmit = (form: IAuditForm): void => {
    const key: string | undefined = this.state.editingDraftKey;

    this.setState((state: IAuditManagerState) => ({
      view: 'drafts',
      editingDraftKey: undefined,
      drafts: state.drafts.map((draft: IAuditDraft) => (draft.key === key ? { ...draft, form, error: undefined } : draft))
    }));
  };

  private _onSaveDrafts = (): void => {
    const { drafts } = this.state;
    const draftNameErrors: { [key: string]: string } = {};
    drafts.forEach((draft: IAuditDraft): void => {
      if (!draft.form.auditName.trim()) {
        draftNameErrors[draft.key] = AUDIT_MESSAGES.nameRequired;
      }
    });
    this.setState({ draftNameErrors });
    if (Object.keys(draftNameErrors).length > 0 || drafts.length === 0) {
      return;
    }

    const { projectId } = this.props;
    const ticketProjectId: string = this.state.ticketProjectId;
    this.setState({ isSaving: true, actionError: undefined });
    this.props.auditService
      .createAudits(drafts.map((draft: IAuditDraft) => toAuditSaveRequest(draft.form, projectId, ticketProjectId)))
      .then(
        (outcome: { message: string; results: IAuditSaveResult[] }): void => {
          if (!this._isActive) {
            return;
          }
          // Rows that failed stay as drafts, carrying their reason; the rest are now audits.
          const failed: IAuditDraft[] = [];
          outcome.results.forEach((result: IAuditSaveResult): void => {
            const draft: IAuditDraft | undefined = drafts[result.index];
            if (!result.success && draft) {
              failed.push({ ...draft, error: result.message || 'This audit could not be added.' });
            }
          });

          this._loadAudits();
          if (failed.length === 0) {
            this.setState({ isSaving: false, view: 'list', drafts: [], notice: outcome.message });
          } else {
            this.setState({ isSaving: false, drafts: failed, actionError: outcome.message });
          }
        },
        (error: unknown): void => {
          if (this._isActive) {
            this.setState({ isSaving: false, actionError: getErrorMessage(error, 'The duplicates could not be saved.') });
          }
        }
      );
  };
}
