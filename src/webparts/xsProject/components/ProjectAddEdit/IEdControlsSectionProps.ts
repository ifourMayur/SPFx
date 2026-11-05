import { ILookupOption } from '../../../../models/Building';
import { FolderLevel, IFolderLevelOptions } from '../../../../models/FolderTree';
import { IEdControlsForm, ITagRowErrors } from '../../../../models/ProjectIntegration';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

export interface IEdControlsSectionProps {
  /** The project being edited; `0` while adding (no tag rows then: folders exist once the project does). */
  projectId: number;
  integrationService: IProjectIntegrationService;
  /** The tag rows' folder cascades, through the same endpoints as the MOM cascade. */
  lookupService: ILookupService;
  /** Supplier options, already loaded by the form. */
  suppliers: ILookupOption[];
  isGrandChildFolderEnabled: boolean;
  /** The form's value; `undefined` until this section has loaded the project's settings. */
  value: IEdControlsForm | undefined;
  ticketProjectError?: string;
  tagRowErrors?: { [key: string]: ITagRowErrors };
  isDisabled: boolean;
  /** The stored value, once loaded - also what the form's Reset returns to. */
  onLoaded: (value: IEdControlsForm) => void;
  onChange: (value: IEdControlsForm) => void;
  /** Opens the project's Audits page; offered while the project's saved link is to an ED Controls project. */
  onOpenAudits?: () => void;
}

export interface IEdControlsSectionState {
  /** `notConfigured`: the account has no ED Controls credentials, so nothing can be linked. */
  status: 'loading' | 'ready' | 'notConfigured' | 'failed';
  loadError?: string;
  /** The project is saved with an ED Controls project linked - what audits are created in. */
  hasSavedLink: boolean;
  projects: ILookupOption[];
  isLoadingProjects: boolean;
  projectsError?: string;
  /** Tags of the selected ED Controls project. */
  tags: string[];
  isLoadingTags: boolean;
  tagsError?: string;
  /** Folder options per tag row, keyed by `ITagRow.key`. */
  rowFolderOptions: { [key: string]: IFolderLevelOptions };
  rowLoadingLevels: { [key: string]: FolderLevel[] };
  /** Messages from "Add tag row", which needs every existing row complete. */
  addRowErrors: { [key: string]: ITagRowErrors };
}
