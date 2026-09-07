import { IFolderSelection, ILookupOption, createEmptyFolderSelection } from './Building';
import {
  FOLDER_LEVELS,
  FolderLevel,
  IFolderLevelOptions,
  clearDeeperLevelOptions,
  createEmptyFolderLevelOptions,
  selectFolderLevel
} from './FolderTree';

function selection(overrides?: Partial<IFolderSelection>): IFolderSelection {
  return { ...createEmptyFolderSelection(), ...overrides };
}

describe('FOLDER_LEVELS', () => {
  it('lists the four cascade levels from shallowest to deepest', () => {
    expect(FOLDER_LEVELS).toEqual(['folderId', 'subFolderId', 'subSubFolderId', 'subSubSubFolderId']);
  });
});

describe('selectFolderLevel', () => {
  it('stores the chosen value at that level', () => {
    expect(selectFolderLevel(selection(), 'folderId', '10').folderId).toBe('10');
  });

  // `onChangeFolderDD` (line 2988) reloads the sub folders and resets both deeper
  // dropdowns to an empty selection.
  it('clears every deeper level when the folder changes', () => {
    const current: IFolderSelection = selection({
      folderId: '10',
      subFolderId: '11',
      subSubFolderId: '12',
      subSubSubFolderId: '13'
    });

    const next: IFolderSelection = selectFolderLevel(current, 'folderId', '20');

    expect(next.subFolderId).toBe('');
    expect(next.subSubFolderId).toBe('');
    expect(next.subSubSubFolderId).toBe('');
  });

  it('clears the two deeper levels when the sub folder changes', () => {
    const current: IFolderSelection = selection({
      folderId: '10',
      subFolderId: '11',
      subSubFolderId: '12',
      subSubSubFolderId: '13'
    });

    const next: IFolderSelection = selectFolderLevel(current, 'subFolderId', '14');

    expect(next.folderId).toBe('10');
    expect(next.subSubFolderId).toBe('');
    expect(next.subSubSubFolderId).toBe('');
  });

  it('clears only the deepest level when the sub sub folder changes', () => {
    const current: IFolderSelection = selection({
      folderId: '10',
      subFolderId: '11',
      subSubFolderId: '12',
      subSubSubFolderId: '13'
    });

    const next: IFolderSelection = selectFolderLevel(current, 'subSubFolderId', '99');

    expect(next.subFolderId).toBe('11');
    expect(next.subSubSubFolderId).toBe('');
  });

  it('clears nothing when the deepest level changes', () => {
    const current: IFolderSelection = selection({ folderId: '10', subFolderId: '11', subSubFolderId: '12' });

    const next: IFolderSelection = selectFolderLevel(current, 'subSubSubFolderId', '13');

    expect(next).toEqual({ ...current, subSubSubFolderId: '13' });
  });

  // The supplier sits beside the cascade in the reference, not inside it, and survives
  // every folder change.
  it('never disturbs the supplier', () => {
    const current: IFolderSelection = selection({ folderId: '10', subFolderId: '11', supplierId: '77' });

    expect(selectFolderLevel(current, 'folderId', '20').supplierId).toBe('77');
  });

  it('clears the deeper levels when the folder is deselected', () => {
    const current: IFolderSelection = selection({ folderId: '10', subFolderId: '11', subSubFolderId: '12' });

    const next: IFolderSelection = selectFolderLevel(current, 'folderId', '');

    expect(next.folderId).toBe('');
    expect(next.subFolderId).toBe('');
    expect(next.subSubFolderId).toBe('');
  });

  it('leaves the original selection untouched', () => {
    const current: IFolderSelection = selection({ folderId: '10', subFolderId: '11' });

    selectFolderLevel(current, 'folderId', '20');

    expect(current.folderId).toBe('10');
    expect(current.subFolderId).toBe('11');
  });

  it('accepts every declared level', () => {
    FOLDER_LEVELS.forEach((level: FolderLevel) => {
      expect(selectFolderLevel(selection(), level, '1')[level]).toBe('1');
    });
  });
});

describe('createEmptyFolderLevelOptions', () => {
  it('gives every level an empty list, so each dropdown renders before anything is loaded', () => {
    const options: IFolderLevelOptions = createEmptyFolderLevelOptions();

    FOLDER_LEVELS.forEach((level: FolderLevel) => {
      expect(options[level]).toEqual([]);
    });
  });
});

describe('clearDeeperLevelOptions', () => {
  const loaded = (): IFolderLevelOptions => ({
    folderId: [{ id: '10', name: 'Drawings' }],
    subFolderId: [{ id: '11', name: 'Architectural' }],
    subSubFolderId: [{ id: '12', name: 'Floor plans' }],
    subSubSubFolderId: [{ id: '13', name: 'Level 1' }]
  });

  // Options below a changed level belonged to the old parent, so they must go at the same
  // time as the selections `selectFolderLevel` clears - otherwise the dropdown would offer
  // another folder's children until the new fetch lands.
  it('empties every level below the one that changed', () => {
    const result: IFolderLevelOptions = clearDeeperLevelOptions(loaded(), 'subFolderId');

    expect(result.subSubFolderId).toEqual([]);
    expect(result.subSubSubFolderId).toEqual([]);
  });

  it('keeps the changed level and everything above it', () => {
    const result: IFolderLevelOptions = clearDeeperLevelOptions(loaded(), 'subFolderId');

    expect(result.folderId).toHaveLength(1);
    expect(result.subFolderId).toHaveLength(1);
  });

  it('clears all three lower levels when the top level changes', () => {
    const result: IFolderLevelOptions = clearDeeperLevelOptions(loaded(), 'folderId');

    expect(result.subFolderId).toEqual([]);
    expect(result.subSubFolderId).toEqual([]);
    expect(result.subSubSubFolderId).toEqual([]);
  });

  it('changes nothing when the deepest level changes', () => {
    const before: IFolderLevelOptions = loaded();

    expect(clearDeeperLevelOptions(before, 'subSubSubFolderId')).toEqual(before);
  });

  it('leaves the original object untouched', () => {
    const before: IFolderLevelOptions = loaded();
    clearDeeperLevelOptions(before, 'folderId');

    expect(before.subFolderId).toHaveLength(1);
  });

  it('accepts a level whose options were never loaded', () => {
    const empty: IFolderLevelOptions = createEmptyFolderLevelOptions();
    const options: ILookupOption[] = clearDeeperLevelOptions(empty, 'folderId').subFolderId;

    expect(options).toEqual([]);
  });
});
