import { ILookupOption } from '../../../../models/Building';
import { IAudit, IAuditDraft } from '../../../../models/EdControlsAudit';
import { IEdControlAuditService } from '../../../../services/EdControlAuditService';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

export interface IAuditManagerProps {
  projectId: number;
  integrationService: IProjectIntegrationService;
  auditService: IEdControlAuditService;
  lookupService: ILookupService;
  isGrandChildFolderEnabled: boolean;
  /** Reports the project's name once read, for the page title. */
  onProjectName?: (name: string) => void;
}

/**
 * What the page shows: the list, the add form, the duplicate count, the duplicate drafts, or
 * the form editing one draft - the reference's list and its three modals.
 */
export type AuditView = 'list' | 'add' | 'count' | 'drafts' | 'editDraft';

export interface IAuditManagerState {
  status: 'loading' | 'ready' | 'failed';
  loadError?: string;
  isConfigured: boolean;
  /** The project's saved ED Controls project; empty means audits cannot be added yet. */
  ticketProjectId: string;
  audits: IAudit[];
  isLoadingAudits: boolean;
  auditsError?: string;
  suppliers: ILookupOption[];
  view: AuditView;
  isSaving: boolean;
  /** A save failure shown on the current view. */
  actionError?: string;
  /** The last success message ("Audit added successfully."). */
  notice?: string;
  /** The audit whose removal is being confirmed. */
  removingId?: number;
  /** The audit being duplicated, and the typed count. */
  duplicateSource?: IAudit;
  countText: string;
  countError?: string;
  drafts: IAuditDraft[];
  /** Per draft: a missing name, from Save duplicates. */
  draftNameErrors: { [key: string]: string };
  editingDraftKey?: string;
}
