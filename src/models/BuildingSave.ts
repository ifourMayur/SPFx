/**
 * Wire shapes for `POST api/Building` - the one endpoint that both inserts and updates a
 * project - plus the pure mapping from the form onto its request body.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package, so the mapping stays unit-testable without a browser or an HTTP client.
 *
 * This is step 8 of `D:\Projects\TaskXS\BMDeskV2\docs\BuildingController-AddEdit-POST-Process.md`
 * - the reference application's `DoActionForPost<BuildingModel>(buildingModel, "Building")`.
 * **Only that single call is modelled here**; the steps around it belong to their own
 * modules. What the save does with them is in `BuildingService`, and the folder tree it
 * provisions is planned in `SharePointFolder.ts`. The reference's "rebind" second POST, its
 * document-storage rows and its tender folders have no counterpart yet.
 *
 * ## Field naming
 *
 * The API deserializes the body into `BuildingAddModel` with the ASP.NET Core *web*
 * defaults, so property matching is case-insensitive and `[JsonPropertyName]` wins where
 * present. The names below are therefore the camelCase forms the server itself emits when
 * it serializes the same model back, which is what makes {@link IBuildingSaveResponse}
 * readable with the same spellings. Two of them are renamed by the server's attributes and
 * are easy to misread:
 *
 * | Wire name | `BuildingAddModel` property | Meaning |
 * | --- | --- | --- |
 * | `projectName` | `BuildingName` | the project's name |
 * | `projectId` | `BuldingID` | a **human-facing project code**, not the primary key |
 *
 * The primary key is `id`, and `id === 0` is what tells the API to insert rather than
 * update (`BuildingController.Post` routes on `model.Id`).
 */

import { IBuildingForm, IFolderSelection } from './Building';
import { IProjectTemplateFolders } from './ProjectTemplateFolder';
import { IFolderProvisionResult } from './SharePointFolder';
import { ISharePointSite } from './SharePointSite';

/** `DocumentStorageTypeEnum` on the server. */
export const DOCUMENT_STORAGE_TYPE = {
  xsCloud: 0,
  sharePoint: 1
};

/**
 * Values the form itself does not hold, which the reference reads out of the MVC session
 * in step 3 of the process document.
 *
 * A SPFx bundle has no such session, so these come from the sign-in response
 * (`ILoginResponse`) and the hosting page instead, and are passed in explicitly rather
 * than reached for from inside the mapping - which is what keeps this module pure.
 */
export interface IProjectSaveContext {
  /**
   * Signed-in user, who becomes the project's client. `ILoginResponse.id`.
   *
   * Used only as the fallback when the form's own read-only `clientId` field is empty, so
   * a project is never inserted against client `0`.
   */
  clientId: number;
  clientName: string;
  /** `ILoginResponse.accountId`. The API's subscription-limit check reads it. */
  accountId: number;
  /**
   * Host name the API associates the project with - `Request.Host.Value` in the reference.
   *
   * It drives multi-tenant resolution and the links in the "you were added to a project"
   * e-mail the API sends to each co-manager, so it is a host name, not a full URL.
   */
  domainName: string;
  /** `Session["MasterCompanyName"]`; appears in that same e-mail. */
  companyName: string;
  /**
   * Company logo for that e-mail. The reference resolves a server-side file path
   * (`ConfigHelper.GetCompanyLogo`), which a browser cannot do, so this is normally empty.
   */
  companyLogo: string;
  /**
   * The SharePoint site the user chose in `SharePointSites`, and where the project's
   * folders are created.
   *
   * Optional because the save itself does not need it: a project is stored by the Web API
   * whether or not its folders can be provisioned. Absent means no folders are created.
   */
  site?: ISharePointSite;
}

/**
 * The part of `GET api/Building?id=` this client reads back before an update.
 *
 * `BuildingService.Get` returns the whole `BuildingModel`, so many more keys arrive than
 * are declared here; only the fields that have to survive an update are typed.
 */
export interface IBuildingDetails {
  /** `0` or absent means the project could not be read - the save is then refused. */
  id: number;
  /**
   * Not preserved - `Update` never assigns it, so it cannot be lost. Read because it is
   * the id the template's folder tree is then fetched under, and the form's Project
   * Template dropdown is locked in edit mode and may not know it.
   */
  projectTemplateId?: number;
  documentStorageType?: number;
  sharepointFolderId?: string;
  /** Base64, which is how System.Text.Json serializes the server's `byte[]`. */
  imageBytes?: string;
  /** ED Controls tag rows, as a JSON string. */
  tagData?: string;
  /** `TicketSourceTypeEnum` - which ticket system the project is bound to. */
  ticketSourceId?: number;
  ticketProjectId?: string;
  /** KYP planning source. Note the server's own spelling of the property. */
  planningSourcetId?: number;
  planningProjectId?: string;
  isDocumentReviewerEnable?: boolean;
  kypProjectAuthToken?: string;
}

/**
 * Fields `ProjectController.Update` assigns from the request **unconditionally**, which
 * this form does not edit.
 *
 * Omitting one does not mean "keep it", it means "set it to null": the Razor view only
 * survives that by round-tripping them through `@Html.HiddenFor`. So an update reads them
 * back with `GET api/Building?id=` and echoes them here.
 *
 * `TagData` needs `ticketSourceId` alongside it to survive at all - the service layer does
 * `building.TagData = building.TicketSourceId == EDcontrols ? model.TagData : null`, so
 * echoing the tag rows without the ticket source would still clear them.
 *
 * **Any new field added to {@link IBuildingSaveRequest} must be checked against `Update`
 * for this.**
 */
export interface IPreservedProjectFields {
  /** Immutable after creation: the storage type is never taken from the form. */
  documentStorageType?: number;
  /** Immutable after creation, and orphans the project's folder tree if lost. */
  sharepointFolderId?: string;
  imageBytes?: string;
  tagData?: string;
  ticketSourceId?: number;
  ticketProjectId?: string;
  planningSourcetId?: number;
  planningProjectId?: string;
  isDocumentReviewerEnable?: boolean;
  kypProjectAuthToken?: string;
}

/**
 * Picks the fields an update has to echo back out of what the API returned.
 *
 * Every value is passed through as received, including `undefined`: a column that is
 * genuinely null must stay null, and dropping the key from the request body is how that is
 * expressed.
 */
export function toPreservedFields(details: IBuildingDetails): IPreservedProjectFields {
  return {
    documentStorageType: details.documentStorageType,
    sharepointFolderId: details.sharepointFolderId,
    imageBytes: details.imageBytes,
    tagData: details.tagData,
    ticketSourceId: details.ticketSourceId,
    ticketProjectId: details.ticketProjectId,
    planningSourcetId: details.planningSourcetId,
    planningProjectId: details.planningProjectId,
    isDocumentReviewerEnable: details.isDocumentReviewerEnable,
    kypProjectAuthToken: details.kypProjectAuthToken
  };
}

/**
 * Request body for `POST api/Building`.
 *
 * Every collection is required rather than optional: `BuildingController.Add` calls
 * `model.spatialBreakdownId.Any()` and `Update` calls `model.UserId.Add(...)` with no null
 * check, so an omitted array arrives as `null` and faults the request server-side. The
 * C# defaults (`= new List<int>()`) only apply when the key is absent from the JSON *and*
 * the property initializer runs, which is not something to rely on.
 *
 * Fields the reference posts but this screen does not edit - the ED Controls tag rows and
 * ticket source, the KYP planning fields, `SharePointSiteURL` - are absent, matching the
 * sections `ProjectAddEdit` deliberately does not port.
 */
export interface IBuildingSaveRequest {
  /** `0` inserts; a project id updates. */
  id: number;
  /** `BuildingName`. Required by the API's `ModelState`. */
  projectName: string;
  /** `BuldingID` - the human-facing project code. */
  projectId: string;
  /** Required by the API's `ModelState`. */
  address: string;
  postcode: string;
  /** City. */
  plaats: string;
  /** Contact person. */
  contactpersoon: string;
  /** Function/role of the contact person. */
  functie: string;
  email: string;
  /** Phone number. */
  telefoon: string;
  /** Always `true`: this path never deactivates a project, and the API forces it anyway. */
  isActive: boolean;
  clientId: number;
  clientName: string;
  /**
   * Omitted when no template is chosen, so the API substitutes the base template
   * (`_projectTemplateRepository.GetBaseTemplateId()`). Sending `0` would instead be
   * stored as template `0`, which no template has.
   */
  projectTemplateId?: number;
  /** Co-managers. `UserId` on the server. */
  userId: number[];
  spatialBreakdownId: number[];
  /**
   * {@link DOCUMENT_STORAGE_TYPE}. `sharePoint` on an insert; on an update it is whatever
   * the project already had, because the storage type is immutable after creation.
   */
  documentStorageType?: number;
  /**
   * The project image, Base64-encoded - which is how the server's `byte[]` is expressed on
   * the wire, in both directions.
   *
   * A newly uploaded image on an insert or an update; otherwise the bytes read back from
   * the project, because `Update` would clear the stored image if this key were absent.
   */
  imageBytes?: string;
  isMaintenanceFolderAvailable: boolean;
  isEnableBIMFolder: boolean;
  isEnableISOFormat: boolean;
  isEnableTender: boolean;
  tenderTemplateId: number;
  /** "Save minutes in this folder". `isAddMOMPDF` on the server. */
  isAddMOMPDF: boolean;
  /**
   * The MOM folder path, as a JSON string - the server stores it verbatim in
   * `Building.MOMFolderData`. See {@link toMomData} for the shape.
   */
  momData?: string;
  accountId: number;
  domainName: string;
  companyName: string;
  companyLogo: string;
  /** Echoed back on an update; see {@link IPreservedProjectFields}. */
  sharepointFolderId?: string;
  tagData?: string;
  ticketSourceId?: number;
  ticketProjectId?: string;
  planningSourcetId?: number;
  planningProjectId?: string;
  isDocumentReviewerEnable?: boolean;
  kypProjectAuthToken?: string;
}

/**
 * Response body of `POST api/Building`, unwrapped from its `ResponseDetail` envelope.
 *
 * The API echoes the whole request model back, so many more keys arrive than are declared
 * here; only the two this client acts on are typed.
 */
export interface IBuildingSaveResponse {
  /** Primary key: server-assigned on an insert, echoed unchanged on an update. */
  id: number;
  /** `BuildingName`, echoed back. */
  projectName?: string;
  /**
   * The template the project ended up on.
   *
   * Worth reading rather than assuming: an insert that sent no template gets the base one
   * filled in by `BuildingController.Add`, so this is the resolved value and the request's
   * own is not.
   */
  projectTemplateId?: number;
  /**
   * Business rejection reported *inside* an HTTP 200 - in practice only the subscription
   * cap ("You've reached your account limit...").
   *
   * It has to be read from here rather than from the envelope: `BuildingController.Post`
   * discards what `Add` reported and always answers `success: true` with an empty
   * envelope message, so checking the envelope alone would treat a refused save as a
   * success. See finding 2 in the process document.
   */
  message?: string;
}

/** Outcome of a successful save. */
export interface IProjectSaveResult {
  /** The id the project now has - newly assigned on an insert. */
  projectId: number;
  /** True when the form was an insert, so the caller can pick the right message. */
  isCreated: boolean;
  /** The response body as received, for the development console dump. */
  response: IBuildingSaveResponse;
  /** The template the project is on, as resolved by the save. */
  projectTemplateId?: number;
  /**
   * The template's folder tree, read once the save succeeded.
   *
   * `undefined` when there was no template id to read it under, or when the read failed -
   * neither of which fails the save, because the project exists by then.
   */
  templateFolders?: IProjectTemplateFolders;
  /**
   * What creating those folders in SharePoint did.
   *
   * `undefined` when no folders were provisioned at all - no site was chosen, the template
   * reported no tree, this was an update rather than an insert, or the library could not be
   * resolved. A present result may still carry failures; see {@link IFolderProvisionResult}.
   */
  folderProvision?: IFolderProvisionResult;
}

/** Parses an `<option>` value into a server id; anything unusable becomes `0`. */
function toId(value: string): number {
  const id: number = Number(value);

  return isFinite(id) && id > 0 ? Math.floor(id) : 0;
}

/** Parses a multi-select's chosen values, dropping anything that is not an id. */
function toIdList(values: string[]): number[] {
  return values.map(toId).filter((id: number) => id > 0);
}

/**
 * The MOM folder path, in the shape the reference's `getAllMOMData()` posts: an array of
 * rows, each carrying the four cascade levels and the supplier as strings.
 *
 * This form binds a single path, so the array always holds one row. It is built field by
 * field rather than by serializing {@link IFolderSelection} directly, so a field added to
 * that view-state shape later cannot silently leak into the stored JSON.
 *
 * Returns `undefined` while the toggle is off, which drops the key from the body - the
 * server then stores `null` rather than a row of empty strings.
 */
export function toMomData(form: IBuildingForm): string | undefined {
  if (!form.isAddMomPdf) {
    return undefined;
  }

  const selection: IFolderSelection = form.momFolder;

  return JSON.stringify([
    {
      folderId: selection.folderId,
      subFolderId: selection.subFolderId,
      subSubFolderId: selection.subSubFolderId,
      subSubSubFolderId: selection.subSubSubFolderId,
      supplierId: selection.supplierId
    }
  ]);
}

/** Separates a data URL's media type from its Base64 payload. */
const BASE64_MARKER: string = ';base64,';

/**
 * The Base64 payload of a `data:` URL, which is exactly what the server's `byte[]` wants.
 *
 * `ProjectAddEdit` reads an uploaded file with `FileReader.readAsDataURL`, so the one
 * string serves both the live preview and the bytes posted here - there is no second copy
 * of the image to keep in sync.
 *
 * Anything that is not a Base64 data URL returns `undefined`, so a value that cannot be
 * turned into bytes falls back to whatever the project already has rather than being sent
 * as nonsense. No decoding, resizing or re-encoding happens: the reference's image
 * processing (its 1280x720 resize branches, step 5 of the process document) is not ported,
 * so the file is posted at its original size.
 */
export function toImageBytes(imageDataUrl: string | undefined): string | undefined {
  const value: string = (imageDataUrl || '').trim();
  const marker: number = value.indexOf(BASE64_MARKER);

  if (value.substring(0, 5).toLowerCase() !== 'data:' || marker < 0) {
    return undefined;
  }

  const base64: string = value.substring(marker + BASE64_MARKER.length);

  return base64.length > 0 ? base64 : undefined;
}

/**
 * Host name of a URL, for {@link IProjectSaveContext.domainName}.
 *
 * Written by hand rather than with `new URL(...)`: the value may already be a bare host
 * name rather than a URL, and a bare host makes `URL` throw. The port is kept, because the
 * reference's `Request.Host.Value` includes it.
 *
 * Returns an empty string for anything unusable, so the caller can fall back.
 */
export function toHostName(value: string): string {
  const withoutScheme: string = (value || '').trim().replace(/^[a-z][a-z0-9+.\-]*:\/\//i, '');
  // Credentials, then the path, query and fragment - whichever comes first.
  const authority: string = withoutScheme.replace(/^[^/@]*@/, '').split(/[/?#]/)[0];

  return authority;
}

/**
 * Maps the finished form onto the request body.
 *
 * Text is trimmed on the way out, mirroring the trimming `validateBuildingForm` already
 * applies when it decides whether a required field was filled in - otherwise a project
 * could pass validation on `" "` and be stored with a whitespace name.
 *
 * On an insert `documentStorageType` is always `sharePoint`. The reference decides it by
 * probing an app registration's credentials and falls back to XS Cloud; a SPFx bundle
 * always runs inside SharePoint as the signed-in user, so there is nothing to probe and no
 * fallback to pick - which also keeps that endpoint's Base64-encoded client secret out of
 * the browser.
 *
 * @param preserved - fields read back from the project, supplied on an **update only**.
 *   Passing them is what stops `ProjectController.Update` from clearing the project's
 *   image, its SharePoint folder id and the third-party integration fields this screen
 *   does not edit - see {@link IPreservedProjectFields}. Their absence is therefore what
 *   marks this request as an insert, and the storage type is set accordingly.
 */
export function toBuildingSaveRequest(
  form: IBuildingForm,
  context: IProjectSaveContext,
  preserved?: IPreservedProjectFields
): IBuildingSaveRequest {
  const clientId: number = toId(form.clientId) || context.clientId;
  // A newly chosen image wins; otherwise the stored bytes are echoed back unchanged.
  const uploaded: string | undefined = toImageBytes(form.imageDataUrl);

  return {
    ...preserved,
    id: form.id,
    projectName: form.buildingName.trim(),
    projectId: form.buldingId.trim(),
    address: form.address.trim(),
    postcode: form.postcode.trim(),
    plaats: form.plaats.trim(),
    contactpersoon: form.contactpersoon.trim(),
    functie: form.functie.trim(),
    email: form.email.trim(),
    telefoon: form.telefoon.trim(),
    isActive: true,
    clientId,
    clientName: context.clientName,
    projectTemplateId: toId(form.projectTemplateId) || undefined,
    userId: toIdList(form.userIds),
    spatialBreakdownId: toIdList(form.spatialBreakdownIds),
    documentStorageType: preserved ? preserved.documentStorageType : DOCUMENT_STORAGE_TYPE.sharePoint,
    imageBytes: uploaded || (preserved ? preserved.imageBytes : undefined),
    isMaintenanceFolderAvailable: form.isMaintenanceFolderAvailable,
    isEnableBIMFolder: form.isEnableBimFolder,
    isEnableISOFormat: form.isEnableIsoFormat,
    isEnableTender: form.isEnableTender,
    // The server field is a non-nullable `int`, and a template only matters while the
    // toggle is on, so a stale selection is not carried along with the toggle switched off.
    tenderTemplateId: form.isEnableTender ? toId(form.tenderTemplateId) : 0,
    isAddMOMPDF: form.isAddMomPdf,
    momData: toMomData(form),
    accountId: context.accountId,
    domainName: context.domainName,
    companyName: context.companyName,
    companyLogo: context.companyLogo
  };
}
