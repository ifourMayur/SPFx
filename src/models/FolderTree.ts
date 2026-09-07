/**
 * Cascading folder selection rules for the `Folder -> SubFolder -> SubSubFolder ->
 * SubSubSubFolder` dropdowns.
 *
 * Pure and dependency-free, in the models layer, so `FolderCascade` can stay a purely
 * presentational component and these rules can be unit-tested without rendering anything
 * or reaching the network.
 *
 * The reference application fetches one level at a time - `GetProjectFolders`, then
 * `GetProjectSubFolders`, then `GetProjectSubSubFolders`, then
 * `GetProjectSubSubSubFolders` - and resets everything below whenever a parent changes;
 * see `onChangeFolderDD` / `onChnageSubFolderDD` / `onChnageSubSubFolderDD`
 * (AddEdit.cshtml lines 2988-3037). `ILookupService` performs those four calls; the two
 * functions here express what has to be discarded when a parent changes, which is the part
 * worth testing.
 */

import { IFolderSelection, ILookupOption } from './Building';

/** The four cascade levels, shallowest first. Index order is what makes "deeper" meaningful. */
export const FOLDER_LEVELS = [
  'folderId',
  'subFolderId',
  'subSubFolderId',
  'subSubSubFolderId'
] as const;

/** One level of the folder cascade; also the {@link IFolderSelection} key it writes to. */
export type FolderLevel = (typeof FOLDER_LEVELS)[number];

/** The options currently loaded for each cascade level. */
export type IFolderLevelOptions = { [level in FolderLevel]: ILookupOption[] };

/** Every level empty, which is how the cascade renders before anything has loaded. */
export function createEmptyFolderLevelOptions(): IFolderLevelOptions {
  return {
    folderId: [],
    subFolderId: [],
    subSubFolderId: [],
    subSubSubFolderId: []
  };
}

/**
 * Applies a choice at `level` and clears every level below it, returning a new selection.
 *
 * Clearing matters because a retained deeper value would no longer belong to the newly
 * chosen parent. The supplier is untouched: the reference keeps it beside the cascade
 * rather than inside it, and never resets it on a folder change.
 */
export function selectFolderLevel(
  selection: IFolderSelection,
  level: FolderLevel,
  value: string
): IFolderSelection {
  const next: IFolderSelection = { ...selection, [level]: value };
  const depth: number = FOLDER_LEVELS.indexOf(level);

  for (let index: number = depth + 1; index < FOLDER_LEVELS.length; index++) {
    next[FOLDER_LEVELS[index]] = '';
  }

  return next;
}

/**
 * Empties the loaded options for every level below `level`, returning a new object.
 *
 * The counterpart to {@link selectFolderLevel}: those options were the old parent's
 * children, so leaving them in place would let a dropdown keep offering another folder's
 * contents until the replacement fetch lands.
 */
export function clearDeeperLevelOptions(
  options: IFolderLevelOptions,
  level: FolderLevel
): IFolderLevelOptions {
  const next: IFolderLevelOptions = { ...options };
  const depth: number = FOLDER_LEVELS.indexOf(level);

  for (let index: number = depth + 1; index < FOLDER_LEVELS.length; index++) {
    next[FOLDER_LEVELS[index]] = [];
  }

  return next;
}
