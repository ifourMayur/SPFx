/**
 * The folder tree a project template defines, as `GET ProjectTemplate/GetSubFolder`
 * returns it.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package.
 *
 * This is the tree a newly created project's SharePoint folders are provisioned from - the
 * reference reads its own version of it in step 5 of
 * `docs/BuildingController-AddEdit-POST-Process.md`, through
 * `Building/GetProjectsubFolder`. `ProjectTemplate/GetSubFolder` is the template-only
 * sibling of that call: it takes no project id and reports the template as authored,
 * without any already-created SharePoint ids mixed in.
 *
 * ## Field naming
 *
 * These are the camelCase names System.Text.Json produces from
 * `GetProjectTemplateDetailsViewModel`, and two groups of them do **not** follow the
 * casing you would guess:
 *
 * - The three list properties are already lower-cased in the C# source
 *   (`lstProjectFolderDetailsViewModel`, `lstProjectsubFolderViewModel`,
 *   `lstProjectsubsubFolderViewModel`), and the camelCase policy leaves a name that does
 *   not start with an upper-case letter completely alone - so they arrive spelled exactly
 *   like that, `lst` prefix and all.
 * - `FolderID` becomes `folderID`, not `folderId`: the policy stops lower-casing as soon
 *   as it meets a character followed by a lower-case one, so a trailing `ID` survives.
 *
 * ## Shape
 *
 * The nesting is three levels of sub-folder deep, and every level after the first reuses
 * one type - `ProjectsubFolderViewModel` - whose children property is called
 * `lstProjectsubsubFolderViewModel` regardless of how deep it actually sits:
 *
 * ```text
 * folder                        lstProjectFolderDetailsViewModel[]
 *   sub-folder                    lstProjectsubFolderViewModel[]
 *     sub-sub-folder                lstProjectsubsubFolderViewModel[]
 *       sub-sub-sub-folder            lstProjectsubsubFolderViewModel[]
 * ```
 */

/** One sub-folder at any depth below the top level. */
export interface IProjectTemplateSubFolder {
  id: number;
  /** Id of the top-level folder this sits under, not of the parent sub-folder. */
  folderID: number;
  subFolderName?: string;
  isActive?: boolean;
  /** Children. Named for the second level, but used at every level below the first. */
  lstProjectsubsubFolderViewModel?: IProjectTemplateSubFolder[];
}

/** One top-level folder of the template. */
export interface IProjectTemplateFolder {
  /** Always `0` here; `foldersId` is the folder's actual id. */
  id?: number;
  foldersId: number;
  folderName?: string;
  isActive?: boolean;
  lstProjectsubFolderViewModel?: IProjectTemplateSubFolder[];
}

/** Response body of `GET ProjectTemplate/GetSubFolder`, unwrapped from its envelope. */
export interface IProjectTemplateFolders {
  lstProjectFolderDetailsViewModel?: IProjectTemplateFolder[];
  templateName?: string;
  templateID?: number;
}

/** Top-level folders of the tree, as a plain array. */
export function toTemplateFolders(
  folders: IProjectTemplateFolders | undefined
): IProjectTemplateFolder[] {
  return folders?.lstProjectFolderDetailsViewModel || [];
}

/**
 * How many folders the tree holds in total, across all four levels.
 *
 * Used to say something meaningful about the response in a log line: "12 folder(s)" tells
 * you the template was read, where the raw object dump alone does not.
 */
export function countTemplateFolders(folders: IProjectTemplateFolders | undefined): number {
  function countSubFolders(subFolders: IProjectTemplateSubFolder[] | undefined): number {
    return (subFolders || []).reduce(
      (total: number, subFolder: IProjectTemplateSubFolder): number =>
        total + 1 + countSubFolders(subFolder.lstProjectsubsubFolderViewModel),
      0
    );
  }

  return toTemplateFolders(folders).reduce(
    (total: number, folder: IProjectTemplateFolder): number =>
      total + 1 + countSubFolders(folder.lstProjectsubFolderViewModel),
    0
  );
}
