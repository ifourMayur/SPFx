import { ILookupOption } from '../../../../models/Building';
import { IAuditForm, IAuditFormErrors } from '../../../../models/EdControlsAudit';
import { FolderLevel, IFolderLevelOptions } from '../../../../models/FolderTree';
import { IEdControlAuditService } from '../../../../services/EdControlAuditService';
import { ILookupService } from '../../../../services/LookupService';

export interface IAuditFormProps {
  /** "Add Audits" or "Edit duplicate". */
  title: string;
  projectId: number;
  /** The project's ED Controls project; its templates are offered. */
  ticketProjectId: string;
  /** The values to start from: empty to add, a draft's to edit it. */
  initial: IAuditForm;
  suppliers: ILookupOption[];
  isGrandChildFolderEnabled: boolean;
  auditService: IEdControlAuditService;
  lookupService: ILookupService;
  isSaving: boolean;
  /** A save failure to show above the buttons. */
  error?: string;
  /** Called with the validated form, its display names filled in from the dropdowns. */
  onSubmit: (form: IAuditForm) => void;
  onCancel: () => void;
}

export interface IAuditFormState {
  form: IAuditForm;
  errors: IAuditFormErrors;
  templates: ILookupOption[];
  isLoadingTemplates: boolean;
  templatesError?: string;
  contacts: ILookupOption[];
  isLoadingContacts: boolean;
  folderOptions: IFolderLevelOptions;
  loadingFolderLevels: FolderLevel[];
}
