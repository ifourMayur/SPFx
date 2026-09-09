/**
 * Tests for provisioning a project's folders, against a fake `ISpHttpClient`.
 *
 * The interesting behaviour is all in what gets requested and in what order: SharePoint
 * cannot create a nested folder before its parent, and a create that fails because the
 * folder is already there is not a failure. Neither is visible from the result alone, so
 * these assert the recorded requests.
 */

// `@microsoft/sp-core-library` needs the SPFx page runtime, which plain jsdom has not got;
// this service pulls it in only for `Log`. Must precede the imports: the TypeScript build
// emits CommonJS without Babel's `jest.mock` hoisting.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    verbose: (): void => undefined,
    info: (): void => undefined,
    warn: (): void => undefined,
    error: (): void => undefined
  }
}));

import { IProjectTemplateFolder } from '../models/ProjectTemplateFolder';
import { IFolderProvisionResult } from '../models/SharePointFolder';
import { ISharePointSite } from '../models/SharePointSite';
import { ISpHttpClient } from './SharePointHttpClient';
import { ISharePointFolderService, SharePointFolderService } from './SharePointFolderService';

const SITE: ISharePointSite = {
  title: 'Contoso Project Site',
  url: 'https://contoso.sharepoint.com/sites/projects',
  siteId: 'site-guid'
};

/** The library root the site reports, with the space that makes encoding worth testing. */
const LIBRARY_ROOT: string = '/sites/projects/Shared Documents';

/** Rules a fake transport applies: any URL containing `match` fails with `message`. */
interface IFailureRule {
  match: string;
  message: string;
}

/** One request the fake transport saw. */
interface IRecordedRequest {
  method: 'GET' | 'POST';
  url: string;
}

/**
 * Transport that reports {@link LIBRARY_ROOT} as the default library and otherwise
 * succeeds, except where a rule says to fail.
 *
 * `existing` names folders that are already in the library: a create addressed at one of
 * them fails the way SharePoint fails it, and the existence check that follows succeeds.
 */
function fakeClient(
  requests: IRecordedRequest[],
  options?: { existing?: string[]; failures?: IFailureRule[] }
): ISpHttpClient {
  const existing: string[] = options?.existing || [];
  const failures: IFailureRule[] = options?.failures || [];

  function ruleFor(url: string): IFailureRule | undefined {
    return failures.filter((rule: IFailureRule): boolean => url.indexOf(rule.match) >= 0)[0];
  }

  /**
   * True when the URL addresses a folder the library already holds.
   *
   * The expected literal is encoded here rather than compared raw: the service percent-
   * encodes the path into the OData literal, so `Shared Documents` reaches the transport as
   * `Shared%20Documents`. Encoding it independently keeps this a check of the URL the
   * service actually produces rather than of a helper shared with it.
   */
  function addressesExisting(url: string): boolean {
    return existing.some((path: string): boolean => {
      const literal: string = encodeURIComponent(
        `${LIBRARY_ROOT}/${path}`.replace(/'/g, "''")
      ).replace(/%2F/gi, '/');

      return url.indexOf(`${literal}')`) >= 0;
    });
  }

  return {
    getJson: async <T>(url: string): Promise<T> => {
      requests.push({ method: 'GET', url });

      if (url.indexOf('defaultDocumentLibrary') >= 0) {
        const rule: IFailureRule | undefined = ruleFor(url);
        if (rule) {
          throw new Error(rule.message);
        }

        return { ServerRelativeUrl: LIBRARY_ROOT } as unknown as T;
      }

      // GetFolderByServerRelativePath: only an existing folder answers.
      if (!addressesExisting(url)) {
        throw new Error('404 Not Found');
      }

      return { Exists: true } as unknown as T;
    },
    postJson: async <T>(url: string): Promise<T> => {
      requests.push({ method: 'POST', url });

      const rule: IFailureRule | undefined = ruleFor(url);
      if (rule) {
        throw new Error(rule.message);
      }

      if (addressesExisting(url)) {
        throw new Error("A folder with the name 'x' already exists.");
      }

      return undefined as unknown as T;
    }
  };
}

function build(requests: IRecordedRequest[], options?: { existing?: string[]; failures?: IFailureRule[] }): ISharePointFolderService {
  return new SharePointFolderService(fakeClient(requests, options));
}

/** A two-level template: Drawings/Architectural plus a sibling Contracts. */
const TEMPLATE: IProjectTemplateFolder[] = [
  {
    foldersId: 5,
    folderName: 'Drawings',
    lstProjectsubFolderViewModel: [{ id: 11, folderID: 5, subFolderName: 'Architectural' }]
  },
  { foldersId: 6, folderName: 'Contracts' }
];

/** Just the folder-create POSTs, in the order they were issued. */
function createdPaths(requests: IRecordedRequest[]): string[] {
  return requests
    .filter((request: IRecordedRequest): boolean => request.method === 'POST')
    .map((request: IRecordedRequest): string => request.url);
}

describe('SharePointFolderService - creating the tree', () => {
  it('resolves the default document library instead of assuming its name', async () => {
    const requests: IRecordedRequest[] = [];
    await build(requests).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(requests[0]).toEqual({
      method: 'GET',
      url: 'https://contoso.sharepoint.com/sites/projects/_api/web/defaultDocumentLibrary/rootFolder?$select=ServerRelativeUrl'
    });
  });

  it('creates the project root first, then each folder under it', async () => {
    const requests: IRecordedRequest[] = [];
    await build(requests).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(createdPaths(requests)).toEqual([
      "https://contoso.sharepoint.com/sites/projects/_api/web/folders/AddUsingPath(DecodedUrl='/sites/projects/Shared%20Documents/Roof%20replacement')",
      "https://contoso.sharepoint.com/sites/projects/_api/web/folders/AddUsingPath(DecodedUrl='/sites/projects/Shared%20Documents/Roof%20replacement/Drawings')",
      "https://contoso.sharepoint.com/sites/projects/_api/web/folders/AddUsingPath(DecodedUrl='/sites/projects/Shared%20Documents/Roof%20replacement/Drawings/Architectural')",
      "https://contoso.sharepoint.com/sites/projects/_api/web/folders/AddUsingPath(DecodedUrl='/sites/projects/Shared%20Documents/Roof%20replacement/Contracts')"
    ]);
  });

  it('reports every folder it created', async () => {
    const result: IFolderProvisionResult = await build([]).createProjectFolders(
      SITE,
      'Roof replacement',
      TEMPLATE
    );

    expect(result.created).toEqual([
      'Roof replacement',
      'Roof replacement/Drawings',
      'Roof replacement/Drawings/Architectural',
      'Roof replacement/Contracts'
    ]);
    expect(result.skipped).toEqual([]);
    expect(result.failed).toEqual([]);
  });

  it('reports an openable URL for the project root', async () => {
    const result: IFolderProvisionResult = await build([]).createProjectFolders(
      SITE,
      'Roof replacement',
      TEMPLATE
    );

    expect(result.rootUrl).toBe(
      'https://contoso.sharepoint.com/sites/projects/Shared%20Documents/Roof%20replacement'
    );
  });

  it('doubles an apostrophe so a folder name cannot break out of the OData literal', async () => {
    const requests: IRecordedRequest[] = [];
    await build(requests).createProjectFolders(SITE, "O'Brien site", []);

    expect(createdPaths(requests)[0]).toContain("O''Brien%20site')");
  });

  it('creates nothing when the project name sanitizes away', async () => {
    const requests: IRecordedRequest[] = [];
    const result: IFolderProvisionResult = await build(requests).createProjectFolders(SITE, '///', TEMPLATE);

    expect(createdPaths(requests)).toEqual([]);
    expect(result.created).toEqual([]);
    expect(result.rootUrl).toBe('');
  });
});

describe('SharePointFolderService - folders that are already there', () => {
  it('counts an existing folder as skipped rather than failed', async () => {
    const result: IFolderProvisionResult = await build([], {
      existing: ['Roof replacement', 'Roof replacement/Drawings']
    }).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(result.skipped).toEqual(['Roof replacement', 'Roof replacement/Drawings']);
    expect(result.created).toEqual(['Roof replacement/Drawings/Architectural', 'Roof replacement/Contracts']);
    expect(result.failed).toEqual([]);
  });

  it('verifies existence only after a create has actually failed', async () => {
    const requests: IRecordedRequest[] = [];
    await build(requests, { existing: ['Roof replacement'] }).createProjectFolders(SITE, 'Roof replacement', []);

    // Library lookup, the failed create, then the check that explains it.
    expect(requests.map((request: IRecordedRequest): string => request.method)).toEqual([
      'GET',
      'POST',
      'GET'
    ]);
    expect(requests[2].url).toContain('GetFolderByServerRelativePath');
  });

  it('makes a whole re-save a no-op', async () => {
    const result: IFolderProvisionResult = await build([], {
      existing: [
        'Roof replacement',
        'Roof replacement/Drawings',
        'Roof replacement/Drawings/Architectural',
        'Roof replacement/Contracts'
      ]
    }).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(result.created).toEqual([]);
    expect(result.failed).toEqual([]);
    expect(result.skipped.length).toBe(4);
  });
});

describe('SharePointFolderService - failures', () => {
  it('reports a failed folder with SharePoint\'s own message, and carries on with its siblings', async () => {
    const result: IFolderProvisionResult = await build([], {
      failures: [{ match: 'Drawings\'', message: 'Access denied. You do not have permission.' }]
    }).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(result.failed[0]).toEqual({
      path: 'Roof replacement/Drawings',
      message: 'Access denied. You do not have permission.'
    });
    // The sibling that does not depend on it is still created.
    expect(result.created).toContain('Roof replacement/Contracts');
  });

  it('does not attempt a folder whose parent failed, but still accounts for it', async () => {
    const requests: IRecordedRequest[] = [];
    const result: IFolderProvisionResult = await build(requests, {
      failures: [{ match: 'Drawings\'', message: 'Access denied.' }]
    }).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(createdPaths(requests).filter((url: string): boolean => url.indexOf('Architectural') >= 0)).toEqual([]);
    expect(result.failed.map((failure): string => failure.path)).toEqual([
      'Roof replacement/Drawings',
      'Roof replacement/Drawings/Architectural'
    ]);
    expect(result.failed[1].message).toContain('parent folder could not be created');
  });

  it('never rejects for a folder it could not create - the project is already saved', async () => {
    const result: IFolderProvisionResult = await build([], {
      failures: [{ match: 'AddUsingPath', message: 'Access denied.' }]
    }).createProjectFolders(SITE, 'Roof replacement', TEMPLATE);

    expect(result.failed.length).toBe(4);
    expect(result.created).toEqual([]);
  });

  it('does reject when the library itself cannot be resolved - nothing was attempted', async () => {
    const service: ISharePointFolderService = build([], {
      failures: [{ match: 'defaultDocumentLibrary', message: '403 Forbidden' }]
    });

    await expect(service.createProjectFolders(SITE, 'Roof replacement', TEMPLATE)).rejects.toThrow(
      '403 Forbidden'
    );
  });
});
