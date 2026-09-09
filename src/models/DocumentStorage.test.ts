import { IProjectTemplateFolder, IProjectTemplateFolders } from './ProjectTemplateFolder';
import { IFolderProvisionResult, IProvisionedFolder } from './SharePointFolder';
import {
  IDocumentStorageRequest,
  toDocumentStorageModel,
  toFolderIdRows
} from './DocumentStorage';

/** A provision result holding exactly the paths named, each with a predictable id. */
function provision(paths: string[], rootId: string = 'id-roof'): IFolderProvisionResult {
  return {
    rootUrl: 'https://contoso.sharepoint.com/sites/projects/Shared%20Documents/Roof',
    rootId,
    folders: paths.map(
      (path: string): IProvisionedFolder => ({
        name: path.split('/').slice(-1)[0],
        path,
        id: `id-${path.toLowerCase().replace(/[^a-z]+/g, '-')}`,
        uniqueId: `guid-${path}`,
        url: `https://contoso.sharepoint.com/${path}`,
        wasCreated: true
      })
    ),
    created: paths,
    skipped: [],
    failed: []
  };
}

/** One top-level folder holding a chain of sub-folders, as the project-scoped read returns it. */
function folder(foldersId: number, folderName: string, subFolderName?: string): IProjectTemplateFolder {
  return {
    foldersId,
    folderName,
    projectTemplateDocumentList: [],
    lstProjectsubFolderViewModel: subFolderName
      ? [
          {
            id: foldersId * 100,
            folderID: foldersId,
            subFolderName,
            lstProjectsubsubFolderViewModel: [],
            projectTemplateDocumentList: []
          }
        ]
      : []
  } as unknown as IProjectTemplateFolder;
}

const TREE: IProjectTemplateFolders = {
  templateID: 4,
  lstProjectFolderDetailsViewModel: [folder(11, 'Drawings'), folder(12, 'Contracts')]
};

describe('toDocumentStorageModel', () => {
  it('binds the root folder id and the project id onto the tree', () => {
    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      TREE,
      provision(['Roof', 'Roof/Drawings', 'Roof/Contracts']),
      77
    );

    expect(model?.projectId).toBe(77);
    expect(model?.sharePointFolderId).toBe('id-roof');
  });

  it('binds each folder id onto the node of that name', () => {
    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      TREE,
      provision(['Roof', 'Roof/Drawings', 'Roof/Contracts']),
      77
    );
    const folders = model?.lstProjectFolderDetailsViewModel || [];

    expect(folders.map((f) => f.folderName)).toEqual(['Drawings', 'Contracts']);
    expect(folders.map((f) => f.sharePointFolderId)).toEqual(['id-roof-drawings', 'id-roof-contracts']);
  });

  it('sets the SharePoint storage type on the project and on every bound folder', () => {
    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      TREE,
      provision(['Roof', 'Roof/Drawings', 'Roof/Contracts']),
      77
    );
    const folders = model?.lstProjectFolderDetailsViewModel || [];

    // `AddList` reads this as `(int)item.DocumentStorageType`, so an unset one faults it.
    expect(model?.documentStorageType).toBe(1);
    expect(folders.map((f) => f.documentStorageType)).toEqual([1, 1]);
  });

  it('matches a sub-folder by its path, so a repeated name cannot take the wrong id', () => {
    const tree: IProjectTemplateFolders = {
      lstProjectFolderDetailsViewModel: [folder(11, 'Design', 'Plans'), folder(12, 'Build', 'Plans')]
    };

    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      tree,
      provision(['Roof', 'Roof/Design', 'Roof/Design/Plans', 'Roof/Build', 'Roof/Build/Plans']),
      77
    );
    const folders = model?.lstProjectFolderDetailsViewModel || [];

    expect(folders[0].lstProjectsubFolderViewModel?.[0].sharePointFolderId).toBe('id-roof-design-plans');
    expect(folders[1].lstProjectsubFolderViewModel?.[0].sharePointFolderId).toBe('id-roof-build-plans');
  });

  it('drops a folder that was never provisioned, along with its children', () => {
    const tree: IProjectTemplateFolders = {
      lstProjectFolderDetailsViewModel: [folder(11, 'Drawings', 'Plans'), folder(12, 'Locked', 'Deep')]
    };

    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      tree,
      provision(['Roof', 'Roof/Drawings', 'Roof/Drawings/Plans']),
      77
    );
    const folders = model?.lstProjectFolderDetailsViewModel || [];

    expect(folders.map((f) => f.folderName)).toEqual(['Drawings']);
  });

  it('drops a sub-folder that was never provisioned but keeps its provisioned parent', () => {
    const tree: IProjectTemplateFolders = {
      lstProjectFolderDetailsViewModel: [folder(11, 'Drawings', 'Locked')]
    };

    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      tree,
      provision(['Roof', 'Roof/Drawings']),
      77
    );
    const folders = model?.lstProjectFolderDetailsViewModel || [];

    expect(folders.map((f) => f.folderName)).toEqual(['Drawings']);
    expect(folders[0].lstProjectsubFolderViewModel).toEqual([]);
  });

  it('keeps the fields it does not bind, so the API is not sent a null list', () => {
    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      'Roof',
      TREE,
      provision(['Roof', 'Roof/Drawings', 'Roof/Contracts']),
      77
    );
    const first = model?.lstProjectFolderDetailsViewModel?.[0] as { projectTemplateDocumentList?: unknown };

    expect(model?.templateID).toBe(4);
    expect(first.projectTemplateDocumentList).toEqual([]);
  });

  it('reports nothing when the root folder has no id, because there is nothing to record', () => {
    expect(toDocumentStorageModel('Roof', TREE, provision(['Roof'], ''), 77)).toBeUndefined();
  });

  it('reports nothing when no folder survived the binding', () => {
    // `AddList` no-ops unless `lstProjectFolderDetailsViewModel.Any()`, so a model whose
    // folders all failed to provision is a request that would write nothing.
    expect(toDocumentStorageModel('Roof', TREE, provision(['Roof']), 77)).toBeUndefined();
  });

  it('reports nothing when no folders were provisioned at all', () => {
    expect(toDocumentStorageModel('Roof', TREE, undefined, 77)).toBeUndefined();
  });

  it('reports nothing when the project has no id to record the folders against', () => {
    expect(toDocumentStorageModel('Roof', TREE, provision(['Roof']), 0)).toBeUndefined();
  });
});

describe('toFolderIdRows', () => {
  const MODEL = toDocumentStorageModel(
    'Roof',
    {
      lstProjectFolderDetailsViewModel: [folder(11, 'Drawings', 'Plans'), folder(12, 'Contracts')]
    },
    provision(['Roof', 'Roof/Drawings', 'Roof/Drawings/Plans', 'Roof/Contracts']),
    77
  );

  it('lists the project root first, then every folder in tree order', () => {
    expect(toFolderIdRows(MODEL, 'Roof').map((row) => row.name)).toEqual([
      'Roof',
      'Drawings',
      'Plans',
      'Contracts'
    ]);
  });

  it('reports each folder id against its name', () => {
    expect(toFolderIdRows(MODEL, 'Roof').map((row) => row.id)).toEqual([
      'id-roof',
      'id-roof-drawings',
      'id-roof-drawings-plans',
      'id-roof-contracts'
    ]);
  });

  it('reports how deep each folder sits, so the nesting can be shown', () => {
    expect(toFolderIdRows(MODEL, 'Roof').map((row) => row.depth)).toEqual([0, 1, 2, 1]);
  });

  it('lists nothing when there is no model', () => {
    expect(toFolderIdRows(undefined, 'Roof')).toEqual([]);
  });
});
