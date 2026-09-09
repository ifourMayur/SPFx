/**
 * Turns a project template's folder tree into the list of folders to create in SharePoint.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package, so the planning and the name sanitizing are unit-testable without a tenant.
 *
 * This is the client-side stand-in for the reference application's
 * `SharePointHelper.CreateProjectFolders`, which walked the same tree with app-only Graph
 * credentials. Keeping the walk here as a pure function means the two quirks that matter -
 * what a folder name may contain, and what order folders have to be created in - are
 * pinned by tests rather than buried in a service that needs a live site to run.
 */

import { IProjectTemplateFolder, IProjectTemplateSubFolder } from './ProjectTemplateFolder';

/** One folder that could not be created, and why. */
export interface IFolderProvisionFailure {
  /** Path relative to the document library root, as {@link planFolderTree} planned it. */
  path: string;
  /** SharePoint's own message where it gave one. */
  message: string;
}

/**
 * One folder that now exists in SharePoint, with the two ids it is known by.
 *
 * A folder created through SharePoint REST has a `UniqueId` - a GUID - and that is **not**
 * the identifier Microsoft Graph addresses it by. Graph uses an opaque, drive-scoped
 * `driveItem.id` (`01VL4HET...`), which is a different identifier rather than a different
 * encoding: decoding one yields none of the other's bytes, so there is no local conversion
 * between them. The Web API reaches these folders through Graph with app-only credentials,
 * exactly as the reference's `SharePointHelper` does, so the Graph one is the id worth
 * reporting - {@link id} - and the GUID is kept beside it only for diagnosing a folder in
 * SharePoint's own APIs.
 */
export interface IProvisionedFolder {
  /** Folder name as created: the template's name after {@link sanitizeFolderName}. */
  name: string;
  /** Path relative to the document library root, as {@link planFolderTree} planned it. */
  path: string;
  /**
   * The `driveItem.id` Graph addresses this folder by, read from the site's own drive API.
   *
   * Empty when that read failed, which is deliberately not the same as a failed folder:
   * the folder exists, it simply cannot be reported to the Web API, and an empty id is what
   * keeps `toDocumentStorageModel` from posting a row that would never resolve.
   */
  id: string;
  /** SharePoint's own `UniqueId`. Empty when SharePoint answered without one. */
  uniqueId: string;
  /** Absolute URL of this folder, for a "go and look" link. */
  url: string;
  /** True when this run created it; false when it was already there. */
  wasCreated: boolean;
}

/**
 * What provisioning a project's folders actually did.
 *
 * Declared in the models layer, not beside the service that produces it, because
 * `IProjectSaveResult` carries it - and models may not import from `services`.
 */
export interface IFolderProvisionResult {
  /** Absolute URL of the project's root folder, for the "go and check it" message. */
  rootUrl: string;
  /**
   * The project root folder's Graph `driveItem.id` - the reference's `SharePointFolderId`.
   *
   * Empty when the root could not be created, or its id could not be read, which is what
   * tells the caller there is nothing worth reporting to the Web API.
   */
  rootId: string;
  /**
   * Every folder that now exists, whether this run created it or found it already there.
   *
   * Skipped folders are included deliberately: a retry after a partial failure has to be
   * able to report the ids of the folders the first run created, and those are skips the
   * second time round.
   */
  folders: IProvisionedFolder[];
  /** Folders this run created. */
  created: string[];
  /** Folders that were already there - a re-save, or a retry after a partial failure. */
  skipped: string[];
  /** Folders that could not be created. Empty on a clean run. */
  failed: IFolderProvisionFailure[];
}

/**
 * Characters SharePoint Online refuses in a file or folder name.
 *
 * `#` and `%` are deliberately absent: they were once forbidden but are supported now, and
 * stripping them would rename folders the template legitimately defines.
 */
const INVALID_NAME_CHARACTERS: RegExp = /["*:<>?/\\|]/g;

/**
 * Makes a template's folder name usable as a SharePoint folder name.
 *
 * Invalid characters become spaces rather than being deleted, so `Plans<Old>` reads as
 * `Plans Old` instead of running the words together; the runs of whitespace that leaves
 * are then collapsed. A name may not begin or end with a period, and `_vti_` is reserved
 * anywhere in it.
 *
 * Returns an empty string when nothing usable is left - the caller drops the folder rather
 * than creating one with a made-up name.
 */
export function sanitizeFolderName(name: string): string {
  const stripped: string = (name || '')
    .replace(INVALID_NAME_CHARACTERS, ' ')
    .replace(/_vti_/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // Trimmed again afterwards, so ". Drawings ." reduces to "Drawings" rather than
  // " Drawings ".
  return stripped.replace(/^\.+/, '').replace(/\.+$/, '').trim();
}

/**
 * Plans every folder to create for one project, as paths relative to the document library
 * root.
 *
 * The first entry is always the project's own root folder, and **every parent precedes its
 * children**, which is what lets the service create them with a single forward loop -
 * SharePoint cannot create `A/B` before `A` exists.
 *
 * ```text
 * Roof replacement
 * Roof replacement/Drawings
 * Roof replacement/Drawings/Architectural
 * Roof replacement/Drawings/Architectural/Floor plans
 * Roof replacement/Contracts
 * ```
 *
 * Nodes whose name sanitizes to nothing are skipped along with their children - there is
 * no folder to hang them under. Paths that repeat are planned once: a template can define
 * two folders whose names differ only in the characters
 * {@link sanitizeFolderName} removes, and SharePoint compares names case-insensitively, so
 * they would collide in the library.
 *
 * @param projectName - the project's name, which becomes the root folder. Returns an empty
 *   plan when it sanitizes to nothing: without a root there is nowhere to put the tree.
 * @param folders - `lstProjectFolderDetailsViewModel` from `ProjectTemplate/GetSubFolder`.
 *   The synthetic `BIM` folder the API appends when `isBIM` is set is an ordinary member of
 *   this list and needs no special handling here.
 */
export function planFolderTree(
  projectName: string,
  folders: IProjectTemplateFolder[] | undefined
): string[] {
  const root: string = sanitizeFolderName(projectName);

  if (!root) {
    return [];
  }

  const paths: string[] = [root];
  const planned: { [path: string]: boolean } = {};
  planned[root.toLowerCase()] = true;

  /**
   * Plans one child of `parent`, returning the path children of its own should hang under.
   *
   * An already-planned path is returned without being added twice, so the duplicate's
   * children still attach to the folder that will exist.
   */
  function plan(parent: string, name: string | undefined): string | undefined {
    const folderName: string = sanitizeFolderName(name || '');

    if (!folderName) {
      return undefined;
    }

    const path: string = `${parent}/${folderName}`;
    const key: string = path.toLowerCase();

    if (!planned[key]) {
      planned[key] = true;
      paths.push(path);
    }

    return path;
  }

  function planSubFolders(parent: string, subFolders: IProjectTemplateSubFolder[] | undefined): void {
    (subFolders || []).forEach((subFolder: IProjectTemplateSubFolder): void => {
      const path: string | undefined = plan(parent, subFolder.subFolderName);

      if (path) {
        // The same property carries the children at every level below the first - see
        // `IProjectTemplateSubFolder`.
        planSubFolders(path, subFolder.lstProjectsubsubFolderViewModel);
      }
    });
  }

  (folders || []).forEach((folder: IProjectTemplateFolder): void => {
    const path: string | undefined = plan(root, folder.folderName);

    if (path) {
      planSubFolders(path, folder.lstProjectsubFolderViewModel);
    }
  });

  return paths;
}
