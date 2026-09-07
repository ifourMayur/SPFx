/**
 * Form and view-state shapes for the project (building) add/edit screen.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any
 * SPFx package, so these stay reusable and unit-testable in isolation.
 *
 * Field names keep the reference application's Dutch-origin spellings (`plaats`,
 * `contactpersoon`, `functie`, `telefoon`) so the mapping back to the Web API's
 * `BuildingModel` / `BuildingAddModel` stays obvious - the same precedent `ProjectList`
 * sets with its Dutch phase columns. Visible labels are the reference's own English
 * resource strings ("City", "Contact person", "Function", "Phone Number").
 */

/**
 * Tenant grouping the reference keeps in session as `domainGroup`, matching
 * `DomainGroupEnum`. Drives the two server-side `@if` gates that survive into this form:
 * McD skips the postcode requirement and hides Spatial Breakdown, TaskXs alone shows the
 * ISO 19650 toggle.
 *
 * Available client-side from `ILoginResponse.domainGroup`, so no extra call is needed.
 */
export type DomainGroup = 'mcD' | 'taskXs' | 'xsWallet';

/** Maps the Web API's numeric `domainGroup` onto {@link DomainGroup}. */
export function toDomainGroup(domainGroup: number | undefined): DomainGroup {
  switch (domainGroup) {
    case 1:
      return 'mcD';
    case 2:
      return 'taskXs';
    case 3:
      return 'xsWallet';
    default:
      // An unknown tenant grouping gets the strictest, most featureful treatment rather
      // than silently losing a required field: `taskXs` enforces the postcode and shows
      // every section.
      return 'taskXs';
  }
}

/** A single `<option>`: the `LookupItem { id, name }` shape the Web API returns. */
export interface ILookupOption {
  id: string;
  name: string;
}

/**
 * One `Folder -> SubFolder -> SubSubFolder -> SubSubSubFolder` path plus its supplier.
 *
 * The deepest level is only ever offered when `isGrandChildFolderEnabled`; the supplier
 * sits alongside the path rather than inside it, matching the reference's own layout.
 */
export interface IFolderSelection {
  folderId: string;
  subFolderId: string;
  subSubFolderId: string;
  subSubSubFolderId: string;
  supplierId: string;
}

/**
 * Everything the project form edits.
 *
 * `id` is `0` in add mode and the project's identifier in edit mode, mirroring how the
 * Web API's single `POST api/Building` endpoint distinguishes an insert from an update.
 *
 * ED Controls (tag rows, ticket source) and KYP planning are deliberately absent: they are
 * third-party integration sections, excluded from this screen.
 */
export interface IBuildingForm {
  id: number;
  /** Object URL or data URL for the uploaded image preview; absent means "no image chosen". */
  imageDataUrl?: string;
  buildingName: string;
  /** The reference's `BuldingID` - a human-facing project code, not the primary key. */
  buldingId: string;
  postcode: string;
  address: string;
  /** City. */
  plaats: string;
  /** Contact person. */
  contactpersoon: string;
  /** Function/role of the contact person. */
  functie: string;
  email: string;
  /** Phone number. */
  telefoon: string;
  /** Read-only in the UI: the signed-in user is the client. */
  clientId: string;
  clientName: string;
  projectTemplateId: string;
  userIds: string[];
  spatialBreakdownIds: string[];
  isMaintenanceFolderAvailable: boolean;
  isEnableBimFolder: boolean;
  isEnableIsoFormat: boolean;
  isEnableTender: boolean;
  tenderTemplateId: string;
  /** "Save minutes in this folder" - reveals {@link momFolder}. */
  isAddMomPdf: boolean;
  momFolder: IFolderSelection;
}

/**
 * Stands in for the reference view's server-side `@if` gates and `ViewBag` flags.
 *
 * Modelled as an explicit prop rather than hidden in the markup, so a ported view makes
 * its multi-tenant behavior visible and testable.
 */
export interface IProjectFormFeatures {
  domainGroup: DomainGroup;
  /** `ViewBag.IsGrandChildFolderEnabled`: offers the fourth folder level. */
  isGrandChildFolderEnabled: boolean;
}

/**
 * Validation messages, keyed by the field they belong to.
 *
 * Only the fields the reference actually validates appear here - see
 * `validateBuildingForm` for why MOM and the folder cascade are absent.
 */
export interface IBuildingFormErrors {
  buildingName?: string;
  address?: string;
  postcode?: string;
  email?: string;
  telefoon?: string;
  tenderTemplateId?: string;
}

/** An empty folder path, used for a new form and to clear the MOM cascade. */
export function createEmptyFolderSelection(): IFolderSelection {
  return {
    folderId: '',
    subFolderId: '',
    subSubFolderId: '',
    subSubSubFolderId: '',
    supplierId: ''
  };
}

/**
 * A blank form in add mode.
 *
 * Every collection starts as an empty array and every text field as an empty string, so
 * the form is always fully controlled - React warns about an input flipping between
 * controlled and uncontrolled, and the Web API's `Add` dereferences `UserId` and
 * `spatialBreakdownId` without a null check.
 *
 * Toggle defaults follow the reference's own: everything off, so no section is revealed
 * until the user asks for it.
 */
export function createEmptyBuildingForm(): IBuildingForm {
  return {
    id: 0,
    buildingName: '',
    buldingId: '',
    postcode: '',
    address: '',
    plaats: '',
    contactpersoon: '',
    functie: '',
    email: '',
    telefoon: '',
    clientId: '',
    clientName: '',
    projectTemplateId: '',
    userIds: [],
    spatialBreakdownIds: [],
    isMaintenanceFolderAvailable: false,
    isEnableBimFolder: false,
    isEnableIsoFormat: false,
    isEnableTender: false,
    tenderTemplateId: '',
    isAddMomPdf: false,
    momFolder: createEmptyFolderSelection()
  };
}
