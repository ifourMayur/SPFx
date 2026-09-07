import { IFolderSelection, ILookupOption } from '../../../../models/Building';
import { FolderLevel, IFolderLevelOptions } from '../../../../models/FolderTree';

export interface IFolderCascadeProps {
  /**
   * Distinguishes this cascade's inputs from any other on the page, since a screen can
   * hold more than one. Used to build the `id`/`htmlFor` pairs that tie each label to its
   * dropdown.
   */
  idPrefix: string;
  /**
   * Options currently loaded for each level.
   *
   * Supplied by the parent rather than derived here: each level comes from its own Web API
   * call (`Building/GetProjectFolders` and its three `…SubFolders` siblings), so only the
   * parent knows what has arrived.
   */
  options: IFolderLevelOptions;
  /** Options for the supplier dropdown that sits below the folder path. */
  suppliers: ILookupOption[];
  /** Current selection. This component holds no state of its own. */
  value: IFolderSelection;
  /**
   * Called with the next selection - deeper levels already cleared - and which level the
   * user changed, so the parent knows which child level to fetch.
   */
  onChange: (selection: IFolderSelection, changedLevel: FolderLevel) => void;
  /** Called when the supplier changes. Separate because it is not part of the cascade. */
  onSupplierChange: (supplierId: string) => void;
  /**
   * Offers the fourth folder level, standing in for `ViewBag.IsGrandChildFolderEnabled`.
   * Defaults to `false`, matching a tenant that has not enabled it.
   */
  isGrandChildFolderEnabled?: boolean;
  /** Levels currently being fetched, shown as a loading placeholder. */
  loadingLevels?: FolderLevel[];
  /** Disables every dropdown, for example while the form is saving. */
  isDisabled?: boolean;
}
