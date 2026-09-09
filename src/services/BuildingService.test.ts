/**
 * End-to-end tests for the project save: `BuildingService` over the real `ApiService`,
 * against a fake transport answering the exact JSON the Web API answers.
 *
 * They exist for the failures that are invisible from the outside. The API replies HTTP 200
 * to a refused save, and `ProjectController.Update` treats an omitted field as "set this to
 * null" - so a save that silently destroys the project's image or orphans its SharePoint
 * folder tree looks exactly like a save that worked.
 */

// `@microsoft/sp-core-library` needs the SPFx page runtime, which plain jsdom has not got;
// `ApiService` and `BuildingService` pull it in only for `Log`. Must precede the imports:
// the TypeScript build emits CommonJS without Babel's `jest.mock` hoisting.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    verbose: (): void => undefined,
    info: (): void => undefined,
    warn: (): void => undefined,
    error: (): void => undefined
  }
}));

import type { HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import { IBuildingForm, createEmptyBuildingForm } from '../models/Building';
import { IProjectSaveContext, IProjectSaveResult } from '../models/BuildingSave';
import { IFolderProvisionResult } from '../models/SharePointFolder';
import { ISharePointSite } from '../models/SharePointSite';
import { ApiService } from './ApiService';
import { ApiTokenStore } from './ApiTokenStore';
import { BuildingService, IBuildingService } from './BuildingService';
import { ProjectTemplateService } from './ProjectTemplateService';
import { ISharePointFolderService } from './SharePointFolderService';

/**
 * The shape `ApiService` needs from a transport.
 *
 * Declared here rather than imported from `./ApiHttpClient`, whose module pulls
 * `@microsoft/sp-http` in as a value and so cannot load under jsdom.
 */
interface IFakeHttpClient {
  fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

/** One request the fake transport saw. */
interface IRecordedRequest {
  method: string;
  url: string;
  headers: Record<string, string>;
  body: Record<string, unknown>;
}

function fakeResponse(body: unknown): HttpClientResponse {
  return {
    ok: true,
    status: 200,
    statusText: 'OK',
    text: async (): Promise<string> => JSON.stringify(body)
  } as unknown as HttpClientResponse;
}

/** The three bodies one save can ask for, keyed by which call asked. */
interface IFakeBodies {
  /** `GET api/Building?id=` - the read-back before an update. */
  read?: unknown;
  /** `POST api/Building` - the save itself. */
  write?: unknown;
  /** `GET ProjectTemplate/GetSubFolder` - the folder tree read afterwards. */
  folders?: unknown;
}

/**
 * Transport that answers each of the save's calls with its own body, recording all of them.
 *
 * One canned body would not do: the save makes two calls to the same route with different
 * verbs, and a third to another controller.
 */
function fakeTransport(bodies: IFakeBodies, requests: IRecordedRequest[]): IFakeHttpClient {
  return {
    fetch: async (url: string, options: IHttpClientOptions): Promise<HttpClientResponse> => {
      const method: string = String(options.method);

      requests.push({
        method,
        url,
        headers: (options.headers || {}) as Record<string, string>,
        body: options.body ? JSON.parse(String(options.body)) : {}
      });

      if (url.indexOf('GetSubFolder') >= 0) {
        return fakeResponse(bodies.folders);
      }

      return fakeResponse(method === 'POST' ? bodies.write : bodies.read);
    }
  };
}

/** What one call to the folder service was asked to create. */
interface IRecordedProvision {
  siteUrl: string;
  projectName: string;
  folderNames: string[];
}

/**
 * A folder service that records what it was asked for and reports a clean run.
 *
 * Stubbed rather than driven through `ISpHttpClient`: what the folders end up looking like
 * is `SharePointFolderService`'s own concern and is tested there. What matters here is
 * *whether* the save calls it, *when*, and with what.
 */
function fakeFolderService(
  provisions: IRecordedProvision[],
  failWith?: Error
): ISharePointFolderService {
  return {
    createProjectFolders: async (site, projectName, folders): Promise<IFolderProvisionResult> => {
      provisions.push({
        siteUrl: site.url,
        projectName,
        folderNames: (folders || []).map((folder) => folder.folderName || '')
      });

      if (failWith) {
        throw failWith;
      }

      return {
        rootUrl: `${site.url}/Shared%20Documents/${projectName}`,
        created: [projectName],
        skipped: [],
        failed: []
      };
    }
  };
}

const SITE: ISharePointSite = {
  title: 'Contoso Project Site',
  url: 'https://contoso.sharepoint.com/sites/projects',
  siteId: 'site-guid'
};

function build(
  bodies: IFakeBodies,
  requests: IRecordedRequest[] = [],
  folderService: ISharePointFolderService = fakeFolderService([])
): IBuildingService {
  const store: ApiTokenStore = new ApiTokenStore();
  store.setToken('app-jwt');
  const apiService: ApiService = new ApiService(
    fakeTransport(bodies, requests),
    { baseUrl: 'https://api.example.com/api' },
    store
  );

  return new BuildingService(
    apiService,
    new ProjectTemplateService(apiService, 'ProjectTemplate'),
    folderService,
    'Building'
  );
}

/** Only the folder-tree GETs, so a test can assert the query it was called with. */
function folderRequests(requests: IRecordedRequest[]): IRecordedRequest[] {
  return requests.filter((request: IRecordedRequest) => request.url.indexOf('GetSubFolder') >= 0);
}

/** The `ResponseDetail` envelope, camelCased by System.Text.Json, around `data`. */
function envelope(data: unknown): Record<string, unknown> {
  return { data, message: null, success: true, messageType: 0 };
}

/**
 * A project as `BuildingService.Get` returns it, trimmed to the fields that matter here.
 *
 * `documentStorageType: 0` is XS Cloud - deliberately not what an insert would send, so a
 * test that the storage type survives an update cannot pass by accident.
 */
const STORED_PROJECT = {
  id: 12,
  projectName: 'Roof replacement',
  documentStorageType: 0,
  sharepointFolderId: 'drive-item-id',
  imageBytes: 'c3RvcmVkLWltYWdl',
  tagData: '[{"tags":["asbestos"]}]',
  ticketSourceId: 2,
  ticketProjectId: 'ED-1',
  planningSourcetId: 3,
  planningProjectId: 'KYP-1',
  isDocumentReviewerEnable: true,
  kypProjectAuthToken: 'kyp-token'
};

function form(overrides?: Partial<IBuildingForm>): IBuildingForm {
  return {
    ...createEmptyBuildingForm(),
    buildingName: 'Roof replacement',
    address: 'Keizersgracht 1',
    postcode: '1015 CJ',
    clientId: '42',
    ...overrides
  };
}

const CONTEXT: IProjectSaveContext = {
  clientId: 42,
  clientName: 'Mayur',
  accountId: 7,
  domainName: 'contoso.sharepoint.com',
  companyName: 'Contoso',
  companyLogo: ''
};

/** A context with a chosen site, which is what the app always has by the time it saves. */
function contextWithSite(): IProjectSaveContext {
  return { ...CONTEXT, site: SITE };
}

/**
 * The template's folder tree, in the exact shape `ProjectTemplate/GetSubFolder` returns.
 *
 * Written out literally - `lst...` prefixes, `folderID` with its trailing capital I and D -
 * so a change in the real casing shows up here as a failure rather than as a silently
 * empty tree.
 */
const TEMPLATE_FOLDERS = {
  templateID: 3,
  templateName: 'Base template',
  lstProjectFolderDetailsViewModel: [
    {
      id: 0,
      foldersId: 5,
      folderName: 'Drawings',
      isActive: true,
      lstProjectsubFolderViewModel: [
        {
          id: 11,
          folderID: 5,
          subFolderName: 'Architectural',
          isActive: true,
          lstProjectsubsubFolderViewModel: [
            { id: 21, folderID: 5, subFolderName: 'Floor plans', lstProjectsubsubFolderViewModel: [] }
          ]
        }
      ]
    }
  ]
};

/** A save that inserted project 18 onto template 3. */
const INSERTED = envelope({ id: 18, projectTemplateId: 3 });

/** A save that updated project 12, which is on template 3. */
const UPDATED = envelope({ id: 12, projectTemplateId: 3 });

/** Every call answered: a save that runs all the way through. */
function allBodies(write: unknown = INSERTED, read?: unknown): IFakeBodies {
  return { read, write, folders: envelope(TEMPLATE_FOLDERS) };
}

describe('BuildingService - inserting a project', () => {
  it('posts to the route the reference posts to', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(form(), CONTEXT);

    expect(requests[0].method).toBe('POST');
    expect(requests[0].url).toBe('https://api.example.com/api/Building');
  });

  it('does not read the project back first - there is no project to read yet', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(form(), CONTEXT);

    const readBacks: IRecordedRequest[] = requests.filter(
      (request: IRecordedRequest) => request.url.indexOf('Building?id=') >= 0
    );

    expect(readBacks).toEqual([]);
  });

  it('reports the id the server assigned', async () => {
    const result: IProjectSaveResult = await build(allBodies()).saveProject(form(), CONTEXT);

    expect(result.projectId).toBe(18);
    expect(result.isCreated).toBe(true);
  });

  it('sends the uploaded image as the Base64 bytes the server stores', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(
      form({ imageDataUrl: 'data:image/png;base64,dXBsb2FkZWQ=' }),
      CONTEXT
    );

    expect(requests[0].body.imageBytes).toBe('dXBsb2FkZWQ=');
  });

  it('presents the application JWT, so the [Authorize]d endpoint accepts the call', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(form(), CONTEXT);

    expect(requests[0].headers.Authorization).toBe('Bearer app-jwt');
  });

  it('sends the project name under the name the server renamed it to', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(form(), CONTEXT);

    expect(requests[0].body.projectName).toBe('Roof replacement');
    expect(requests[0].body.buildingName).toBeUndefined();
  });
});

describe('BuildingService - reading the template folders after a save', () => {
  it('reads them once the save succeeded, not before', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(form(), CONTEXT);

    expect(requests.map((request: IRecordedRequest) => request.method)).toEqual(['POST', 'GET']);
  });

  it('reads them under the template the server resolved, not the one the form sent', async () => {
    const requests: IRecordedRequest[] = [];
    // The form chose no template, so `Add` filled in the base one and echoed it back.
    await build(allBodies(), requests).saveProject(form({ projectTemplateId: '' }), CONTEXT);

    expect(folderRequests(requests)[0].url).toBe(
      'https://api.example.com/api/ProjectTemplate/GetSubFolder?id=3&isBIM=false'
    );
  });

  it('passes the BIM switch through, which is what appends the BIM folder', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(), requests).saveProject(form({ isEnableBimFolder: true }), CONTEXT);

    expect(folderRequests(requests)[0].url).toBe(
      'https://api.example.com/api/ProjectTemplate/GetSubFolder?id=3&isBIM=true'
    );
  });

  it('returns the tree with its nesting intact', async () => {
    const result: IProjectSaveResult = await build(allBodies()).saveProject(form(), CONTEXT);

    expect(result.templateFolders).toEqual(TEMPLATE_FOLDERS);
    expect(result.projectTemplateId).toBe(3);
  });

  it('falls back to the template the form sent when the save echoes none', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(envelope({ id: 18 })), requests).saveProject(
      form({ projectTemplateId: '9' }),
      CONTEXT
    );

    expect(folderRequests(requests)[0].url).toBe(
      'https://api.example.com/api/ProjectTemplate/GetSubFolder?id=9&isBIM=false'
    );
  });

  it('skips the read when no template can be resolved, rather than asking for template 0', async () => {
    const requests: IRecordedRequest[] = [];
    const result: IProjectSaveResult = await build(
      allBodies(envelope({ id: 18 })),
      requests
    ).saveProject(form({ projectTemplateId: '' }), CONTEXT);

    expect(folderRequests(requests)).toEqual([]);
    expect(result.templateFolders).toBeUndefined();
    // The save itself still succeeded.
    expect(result.projectId).toBe(18);
  });

  it('does not fail the save when the folders cannot be read', async () => {
    const result: IProjectSaveResult = await build({
      write: INSERTED,
      // The project exists by now, so reporting a failure would invite a duplicate.
      folders: { data: null, success: false, message: 'Template not found.', messageType: 3 }
    }).saveProject(form(), CONTEXT);

    expect(result.projectId).toBe(18);
    expect(result.templateFolders).toBeUndefined();
  });
});

describe('BuildingService - updating a project', () => {
  it('reads the project back before posting', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(UPDATED, envelope(STORED_PROJECT)), requests).saveProject(
      form({ id: 12 }),
      CONTEXT
    );

    expect(requests.map((request: IRecordedRequest) => request.method)).toEqual(['GET', 'POST', 'GET']);
    expect(requests[0].url).toBe('https://api.example.com/api/Building?id=12');
  });

  it('echoes back every field Update would otherwise null', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(UPDATED, envelope(STORED_PROJECT)), requests).saveProject(
      form({ id: 12 }),
      CONTEXT
    );

    const body: Record<string, unknown> = requests[1].body;
    // The image and the folder id are the two that lose user data outright.
    expect(body.imageBytes).toBe('c3RvcmVkLWltYWdl');
    expect(body.sharepointFolderId).toBe('drive-item-id');
    // Tag rows only survive alongside the ticket source they are gated on.
    expect(body.tagData).toBe('[{"tags":["asbestos"]}]');
    expect(body.ticketSourceId).toBe(2);
    expect(body.planningProjectId).toBe('KYP-1');
    expect(body.kypProjectAuthToken).toBe('kyp-token');
    expect(body.isDocumentReviewerEnable).toBe(true);
  });

  it('keeps the storage type the project already had, not the one an insert would use', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(UPDATED, envelope(STORED_PROJECT)), requests).saveProject(
      form({ id: 12 }),
      CONTEXT
    );

    expect(requests[1].body.documentStorageType).toBe(0);
  });

  it('replaces the stored image when a new one was chosen', async () => {
    const requests: IRecordedRequest[] = [];
    await build(allBodies(UPDATED, envelope(STORED_PROJECT)), requests).saveProject(
      form({ id: 12, imageDataUrl: 'data:image/jpeg;base64,cmVwbGFjZWQ=' }),
      CONTEXT
    );

    expect(requests[1].body.imageBytes).toBe('cmVwbGFjZWQ=');
  });

  it('reads the folders under the template the project is on, which the form has locked', async () => {
    const requests: IRecordedRequest[] = [];
    await build(
      allBodies(
        // An update echoes back only what was sent, and the locked dropdown sent nothing.
        envelope({ id: 12 }),
        envelope({ ...STORED_PROJECT, projectTemplateId: 4 })
      ),
      requests
    ).saveProject(form({ id: 12, projectTemplateId: '' }), CONTEXT);

    expect(folderRequests(requests)[0].url).toBe(
      'https://api.example.com/api/ProjectTemplate/GetSubFolder?id=4&isBIM=false'
    );
  });

  it('refuses the save when the project cannot be read back', async () => {
    const requests: IRecordedRequest[] = [];
    const service: IBuildingService = build(allBodies(UPDATED, envelope(null)), requests);

    await expect(service.saveProject(form({ id: 12 }), CONTEXT)).rejects.toThrow(
      /could not be read back/
    );
    // Nothing was posted: a save that blanks the project is worse than no save.
    expect(requests.map((request: IRecordedRequest) => request.method)).toEqual(['GET']);
  });

  it('refuses the save when the read-back reports a failure inside an HTTP 200', async () => {
    const requests: IRecordedRequest[] = [];
    const service: IBuildingService = build(
      allBodies(UPDATED, { data: null, success: false, message: 'Project not found.', messageType: 3 }),
      requests
    );

    await expect(service.saveProject(form({ id: 12 }), CONTEXT)).rejects.toThrow('Project not found.');
    expect(requests.map((request: IRecordedRequest) => request.method)).toEqual(['GET']);
  });
});

describe('BuildingService - creating the folders in SharePoint', () => {
  it('creates them from the tree the template reported, in the chosen site', async () => {
    const provisions: IRecordedProvision[] = [];
    await build(allBodies(), [], fakeFolderService(provisions)).saveProject(
      form(),
      contextWithSite()
    );

    expect(provisions).toEqual([
      {
        siteUrl: 'https://contoso.sharepoint.com/sites/projects',
        projectName: 'Roof replacement',
        // Read straight off `lstProjectFolderDetailsViewModel`.
        folderNames: ['Drawings']
      }
    ]);
  });

  it('reports what provisioning did, root URL included', async () => {
    const result: IProjectSaveResult = await build(allBodies()).saveProject(
      form(),
      contextWithSite()
    );

    expect(result.folderProvision?.rootUrl).toBe(
      'https://contoso.sharepoint.com/sites/projects/Shared%20Documents/Roof replacement'
    );
  });

  it('names the root folder after the project the API stored, not the raw input', async () => {
    const provisions: IRecordedProvision[] = [];
    await build(allBodies(), [], fakeFolderService(provisions)).saveProject(
      form({ buildingName: '  Roof replacement  ' }),
      contextWithSite()
    );

    expect(provisions[0].projectName).toBe('Roof replacement');
  });

  it('creates nothing when no site was chosen', async () => {
    const provisions: IRecordedProvision[] = [];
    const result: IProjectSaveResult = await build(
      allBodies(),
      [],
      fakeFolderService(provisions)
    ).saveProject(form(), CONTEXT);

    expect(provisions).toEqual([]);
    expect(result.folderProvision).toBeUndefined();
    // The save itself is unaffected.
    expect(result.projectId).toBe(18);
  });

  it('creates nothing on an update - a renamed project would get a second root folder', async () => {
    const provisions: IRecordedProvision[] = [];
    const result: IProjectSaveResult = await build(
      allBodies(UPDATED, envelope(STORED_PROJECT)),
      [],
      fakeFolderService(provisions)
    ).saveProject(form({ id: 12 }), contextWithSite());

    expect(provisions).toEqual([]);
    expect(result.folderProvision).toBeUndefined();
    expect(result.projectId).toBe(12);
  });

  it('creates nothing after a refused save', async () => {
    const provisions: IRecordedProvision[] = [];
    const service: IBuildingService = build(
      allBodies(envelope({ id: 0, message: 'Limit reached.' })),
      [],
      fakeFolderService(provisions)
    );

    await expect(service.saveProject(form(), contextWithSite())).rejects.toThrow('Limit reached.');
    expect(provisions).toEqual([]);
  });

  it('does not fail the save when the library cannot be resolved', async () => {
    const result: IProjectSaveResult = await build(
      allBodies(),
      [],
      // The project exists by now, so rejecting would invite a duplicate.
      fakeFolderService([], new Error('403 Forbidden'))
    ).saveProject(form(), contextWithSite());

    expect(result.projectId).toBe(18);
    expect(result.folderProvision).toBeUndefined();
  });
});

describe('BuildingService - printing the responses', () => {
  /** Everything `console.log` was handed, as `[label, payload]` pairs. */
  function captureLogs(): unknown[][] {
    const lines: unknown[][] = [];
    jest.spyOn(console, 'log').mockImplementation((...args: unknown[]): void => {
      lines.push(args);
    });

    return lines;
  }

  /** The payload printed under a label containing `marker`. */
  function printed(lines: unknown[][], marker: string): unknown {
    return lines.filter((line: unknown[]) => String(line[0]).indexOf(marker) >= 0)[0]?.[1];
  }

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('prints both responses, envelope and all', async () => {
    const lines: unknown[][] = captureLogs();
    await build(allBodies()).saveProject(form(), CONTEXT);

    expect(printed(lines, 'BuildingService')).toEqual(INSERTED);
    expect(printed(lines, 'ProjectTemplateService')).toEqual(envelope(TEMPLATE_FOLDERS));
  });

  it('prints a refused save, which is the response most worth seeing', async () => {
    const lines: unknown[][] = captureLogs();
    const refusal = { data: null, success: false, message: 'Project Name is required', messageType: 3 };

    await expect(build(allBodies(refusal)).saveProject(form(), CONTEXT)).rejects.toThrow();

    // The dump has to precede the unwrap that throws, or this line never runs.
    expect(printed(lines, 'BuildingService')).toEqual(refusal);
  });

  it('prints a rejected folder read, whose failure the save deliberately swallows', async () => {
    const lines: unknown[][] = captureLogs();
    const refusal = { data: null, success: false, message: 'Template not found.', messageType: 3 };

    await build({ write: INSERTED, folders: refusal }).saveProject(form(), CONTEXT);

    expect(printed(lines, 'ProjectTemplateService')).toEqual(refusal);
  });
});

describe('BuildingService - failures reported inside an HTTP 200', () => {
  it('treats the subscription cap in data.message as a failure, not a save', async () => {
    // `ProjectController.Post` discards what `Add` reported and always answers
    // `success: true`, so the refusal only appears inside the echoed model.
    const service: IBuildingService = build(
      allBodies(envelope({ id: 0, message: 'Account limit reached, please upgrade.' }))
    );

    await expect(service.saveProject(form(), CONTEXT)).rejects.toThrow(
      'Account limit reached, please upgrade.'
    );
  });

  it('does not read the template folders after a refused save', async () => {
    const requests: IRecordedRequest[] = [];
    const service: IBuildingService = build(
      allBodies(envelope({ id: 0, projectTemplateId: 3, message: 'Limit reached.' })),
      requests
    );

    await expect(service.saveProject(form(), CONTEXT)).rejects.toThrow('Limit reached.');
    expect(folderRequests(requests)).toEqual([]);
  });

  it('treats a ModelState rejection as a failure', async () => {
    const service: IBuildingService = build(
      allBodies({ data: null, success: false, message: 'Project Name is required', messageType: 3 })
    );

    await expect(service.saveProject(form(), CONTEXT)).rejects.toThrow('Project Name is required');
  });

  it('rejects a save that comes back without a project id', async () => {
    const service: IBuildingService = build(allBodies(envelope({ id: 0 })));

    await expect(service.saveProject(form(), CONTEXT)).rejects.toThrow(/without returning its id/);
  });
});
