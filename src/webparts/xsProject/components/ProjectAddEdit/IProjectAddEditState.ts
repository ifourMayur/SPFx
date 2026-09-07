import { IBuildingForm, IBuildingFormErrors, ILookupOption } from '../../../../models/Building';
import { FolderLevel, IFolderLevelOptions } from '../../../../models/FolderTree';

/** Options for every dropdown on the form, as loaded from the Web API. */
export interface IProjectFormLookups {
  projectTemplates: ILookupOption[];
  users: ILookupOption[];
  spatialBreakdowns: ILookupOption[];
  tenderTemplates: ILookupOption[];
  suppliers: ILookupOption[];
}

export interface IProjectAddEditState {
  /** The single source of truth for every input; the form is fully controlled. */
  form: IBuildingForm;
  /** Per-field validation messages. Empty means the form may be saved. */
  errors: IBuildingFormErrors;
  /** True while the dropdown data is being fetched, before the form is usable. */
  isLoading: boolean;
  /** Message shown when the dropdown data could not be loaded, with a retry offered. */
  loadError?: string;
  /** Options for each dropdown. Empty lists until the load completes. */
  lookups: IProjectFormLookups;
  /** Options loaded for each folder cascade level. */
  folderOptions: IFolderLevelOptions;
  /** Cascade levels currently being fetched, so each dropdown can show its own state. */
  loadingFolderLevels: FolderLevel[];
  /** True while `onSave` is in flight; disables the actions and shows the saving state. */
  isSaving: boolean;
  /** Form-level failure - a rejected `onSave` - shown above the actions. */
  saveError?: string;
}
