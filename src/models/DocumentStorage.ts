/**
 * The folder tree as SharePoint now actually holds it, ready to report back to the Web API
 * through `POST Document/AddDocumentStorageDetails`.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package, so the binding is unit-testable without a tenant or an HTTP client.
 *
 * This is step 4 of the reference application's SharePoint branch (step 9d of
 * `D:\Projects\TaskXS\BMDeskV2\docs\BuildingController-AddEdit-POST-Process.md`):
 * `DoActionForPost<GetProjectTemplateDetailsViewModel>(responseDetail.Result.Data,
 * "Document/AddDocumentStorageDetails")`. There the model is a by-product of provisioning -
 * `SharePointHelper.CreateProjectFolders` sets `SharePointFolderId` on each node as it
 * creates it, and returns the tree it was given. This client provisions through SharePoint
 * REST instead, from a plan of paths, so the ids come back as a flat
 * {@link IProvisionedFolder} list and are folded onto the tree here.
 *
 * ## Which tree
 *
 * **Not** the tree the folders were created from. Folder creation is driven by
 * `ProjectTemplate/GetSubFolder`, whose sub-folder ids are `SubFolder.ID` - the template as
 * authored. `DocumentStorageDetailsService.AddList` matches sub-folders with
 * `projectSubFolderList.FirstOrDefault(x => x.Id == item.Id)` against `ProjectSubFolder`
 * rows scoped to the project, so it needs the ids from `Building/GetProjectsubFolder`
 * instead. Posting the template's ids would either match nothing or, on a numeric
 * collision, write a folder id onto the wrong sub-folder of that project. Top-level folders
 * are unaffected - `foldersId` is the global `Folder.ID` in both reads.
 */

import { DOCUMENT_STORAGE_TYPE } from './BuildingSave';
import {
  IProjectTemplateFolder,
  IProjectTemplateFolders,
  IProjectTemplateSubFolder
} from './ProjectTemplateFolder';
import { IFolderProvisionResult, IProvisionedFolder, sanitizeFolderName } from './SharePointFolder';

/**
 * Request body of `POST Document/AddDocumentStorageDetails`.
 *
 * The same `GetProjectTemplateDetailsViewModel` the read answered with, which is why this
 * extends {@link IProjectTemplateFolders}: the endpoint is handed the model it gave out,
 * with the SharePoint ids filled in. The three fields it narrows to required are the ones
 * `AddList` reads without a null check.
 */
export interface IDocumentStorageRequest extends IProjectTemplateFolders {
  /** Becomes `BuildingId` on every row `AddList` writes. */
  projectId: number;
  /** {@link DOCUMENT_STORAGE_TYPE}; always `sharePoint` by the time this is built. */
  documentStorageType: number;
  /** The project root folder's SharePoint id. */
  sharePointFolderId: string;
}

/**
 * What reporting the provisioned folders to the Web API did.
 *
 * Carried on `IProjectSaveResult` rather than thrown, for the same reason the folder
 * provisioning is: by the time this runs the project and its folders both exist, so a
 * failure here must not read as a failed save and invite the user to create a duplicate.
 */
export interface IDocumentStorageResult {
  /** True when the API accepted the call and wrote the rows. */
  isRecorded: boolean;
  /**
   * The project root folder's name, as SharePoint actually named it.
   *
   * Client-side only - it is not on {@link IDocumentStorageRequest} and so is never posted.
   * It exists because the model carries an id for the root but no name for it, and
   * {@link toFolderIdRows} needs one to label that row.
   */
  rootName: string;
  /** Why it did not, when it did not. Empty on success. */
  message: string;
  /** The model that was posted, so the UI can show what was reported and under which ids. */
  model: IDocumentStorageRequest;
}

/** One folder of a bound model, flattened for display. */
export interface IFolderIdRow {
  /** The folder's own name, not its path. */
  name: string;
  /** The SharePoint id recorded for it. */
  id: string;
  /** How far below the project root it sits; the root itself is `0`. */
  depth: number;
}

/**
 * Flattens a bound model into one row per folder, in tree order, for showing what was
 * recorded.
 *
 * Presentation-shaped but kept here rather than in the component, so the walk over four
 * levels of a tree whose child property is named for the second one is pinned by tests
 * instead of being re-derived in JSX.
 *
 * @param model - the bound model, as {@link toDocumentStorageModel} built it.
 * @param rootName - the project root folder's name. It is not on the model: the model is
 *   posted verbatim, and the reference does not send a name there, so it is passed in
 *   rather than added to the wire.
 */
export function toFolderIdRows(
  model: IDocumentStorageRequest | undefined,
  rootName: string
): IFolderIdRow[] {
  if (!model) {
    return [];
  }

  const rows: IFolderIdRow[] = [
    { name: rootName, id: model.sharePointFolderId, depth: 0 }
  ];

  function addSubFolders(subFolders: IProjectTemplateSubFolder[] | undefined, depth: number): void {
    (subFolders || []).forEach((subFolder: IProjectTemplateSubFolder): void => {
      rows.push({
        name: subFolder.subFolderName || '',
        id: subFolder.sharePointFolderId || '',
        depth
      });
      addSubFolders(subFolder.lstProjectsubsubFolderViewModel, depth + 1);
    });
  }

  (model.lstProjectFolderDetailsViewModel || []).forEach((folder: IProjectTemplateFolder): void => {
    rows.push({
      name: folder.folderName || '',
      id: folder.sharePointFolderId || '',
      depth: 1
    });
    addSubFolders(folder.lstProjectsubFolderViewModel, 2);
  });

  return rows;
}

/**
 * Folds the ids SharePoint handed out onto the project's folder tree.
 *
 * Nodes are matched **by path**, rebuilt exactly as `planFolderTree` built it - the same
 * root, the same {@link sanitizeFolderName} on every segment. Matching on the bare folder
 * name would be ambiguous: a template may well define `Design/Plans` and `Build/Plans`, and
 * whichever id was found first would land on both.
 *
 * A node with no provisioned folder is **dropped, along with its children**, so what is
 * reported is only what SharePoint actually has. That matters twice over: rows carrying an
 * empty folder id would be written for folders that do not exist, and `AddList` reads a
 * top-level folder's storage type as `(int)item.DocumentStorageType` with no null check, so
 * a node left unbound would fault the whole call server-side. Children go with their parent
 * because a folder whose parent could not be created was never attempted.
 *
 * Every other field is passed through untouched. That is deliberate rather than incidental:
 * `GetProjectTemplateDetailsViewModel` initializes its lists in its constructor, so an
 * absent key deserializes to an empty list - but an explicit `null` overwrites the
 * initializer, and `AddList` iterates `ProjectTemplateDocumentList` without a null check.
 * Echoing the API's own payload back keeps those lists as it sent them.
 *
 * @param projectName - the project's name, which is the root folder's name.
 * @param folders - the **project-scoped** tree, from `Building/GetProjectsubFolder`. See
 *   the module docblock for why the template-scoped read will not do.
 * @param provision - what {@link ISharePointFolderService} created.
 * @param projectId - the id the save assigned, which the reference sets on the model
 *   immediately before posting it (`responseDetail.Result.Data.ProjectId = response.Id`).
 * @returns the model to post, or `undefined` when there is nothing worth recording - no
 *   folders were provisioned, the root folder has no id, the project has none, or no folder
 *   of the tree survived the binding.
 */
export function toDocumentStorageModel(
  projectName: string,
  folders: IProjectTemplateFolders | undefined,
  provision: IFolderProvisionResult | undefined,
  projectId: number
): IDocumentStorageRequest | undefined {
  const rootId: string = (provision?.rootId || '').trim();
  const root: string = sanitizeFolderName(projectName);

  if (!rootId || !root || !(projectId > 0)) {
    return undefined;
  }

  /** Provisioned id by lower-cased path - `planFolderTree` dedupes case-insensitively. */
  const idByPath: { [path: string]: string } = {};
  (provision?.folders || []).forEach((provisioned: IProvisionedFolder): void => {
    if (provisioned.id) {
      idByPath[provisioned.path.toLowerCase()] = provisioned.id;
    }
  });

  /** The id SharePoint gave the child of `parent` named `name`, or `''`. */
  function idOf(parent: string, name: string | undefined): string {
    const folderName: string = sanitizeFolderName(name || '');

    return folderName ? idByPath[`${parent}/${folderName}`.toLowerCase()] || '' : '';
  }

  function bindSubFolders(
    parent: string,
    subFolders: IProjectTemplateSubFolder[] | undefined
  ): IProjectTemplateSubFolder[] {
    const bound: IProjectTemplateSubFolder[] = [];

    (subFolders || []).forEach((subFolder: IProjectTemplateSubFolder): void => {
      const id: string = idOf(parent, subFolder.subFolderName);

      if (!id) {
        return;
      }

      const path: string = `${parent}/${sanitizeFolderName(subFolder.subFolderName || '')}`;

      bound.push({
        ...subFolder,
        documentStorageType: DOCUMENT_STORAGE_TYPE.sharePoint,
        sharePointFolderId: id,
        // The same property carries the children at every level below the first - see
        // `IProjectTemplateSubFolder`.
        lstProjectsubsubFolderViewModel: bindSubFolders(
          path,
          subFolder.lstProjectsubsubFolderViewModel
        )
      });
    });

    return bound;
  }

  const bound: IProjectTemplateFolder[] = [];

  (folders?.lstProjectFolderDetailsViewModel || []).forEach((folder: IProjectTemplateFolder): void => {
    const id: string = idOf(root, folder.folderName);

    if (!id) {
      return;
    }

    const path: string = `${root}/${sanitizeFolderName(folder.folderName || '')}`;

    bound.push({
      ...folder,
      documentStorageType: DOCUMENT_STORAGE_TYPE.sharePoint,
      sharePointFolderId: id,
      lstProjectsubFolderViewModel: bindSubFolders(path, folder.lstProjectsubFolderViewModel)
    });
  });

  // `AddList` guards its whole body with `lstProjectFolderDetailsViewModel.Any()`, so a
  // model whose folders all failed to provision would write nothing at all. Reported as
  // "nothing to record" rather than posted, so the caller can say so.
  if (bound.length === 0) {
    return undefined;
  }

  return {
    ...folders,
    projectId,
    documentStorageType: DOCUMENT_STORAGE_TYPE.sharePoint,
    sharePointFolderId: rootId,
    lstProjectFolderDetailsViewModel: bound
  };
}
