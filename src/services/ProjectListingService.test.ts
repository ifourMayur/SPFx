/**
 * End-to-end binding tests for the project overview table: `ProjectListingService` over
 * the real `ApiService`, against a fake transport answering the exact JSON the Web API
 * answers.
 *
 * These exist because the failures worth catching here are invisible from the outside.
 * A refused request arrives as an HTTP 200 carrying `success: false`, and the total row
 * count arrives stamped on the rows rather than beside them - get either wrong and the
 * table renders, without an error, showing nothing or paging forever.
 */

// `@microsoft/sp-core-library` needs the SPFx page runtime, which plain jsdom has not got;
// `ApiService` pulls it in only for `Log`. Must precede the imports: the TypeScript build
// emits CommonJS without Babel's `jest.mock` hoisting.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    verbose: (): void => undefined,
    info: (): void => undefined,
    warn: (): void => undefined,
    error: (): void => undefined
  }
}));

import type { HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import { IProjectListPage, IProjectListQuery } from '../models/ProjectListing';
import { ApiService } from './ApiService';
import { ApiTokenStore } from './ApiTokenStore';
import { IProjectListingService, ProjectListingService } from './ProjectListingService';

/**
 * The shape `ApiService` needs from a transport.
 *
 * Declared here rather than imported from `./ApiHttpClient`, whose module pulls
 * `@microsoft/sp-http` in as a value and so cannot load under jsdom.
 */
interface IFakeHttpClient {
  fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

/** One request the fake transport saw, with its body already parsed. */
interface IRecordedRequest {
  url: string;
  options: IHttpClientOptions;
  body: Record<string, unknown>;
}

function fakeResponse(body: unknown, status: number = 200): HttpClientResponse {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? 'OK' : 'Error',
    text: async (): Promise<string> => (typeof body === 'string' ? body : JSON.stringify(body))
  } as unknown as HttpClientResponse;
}

function fakeTransport(
  body: unknown,
  requests: IRecordedRequest[],
  status: number = 200
): IFakeHttpClient {
  return {
    fetch: async (url: string, options: IHttpClientOptions): Promise<HttpClientResponse> => {
      requests.push({
        url,
        options,
        body: options.body ? JSON.parse(String(options.body)) : {}
      });
      return fakeResponse(body, status);
    }
  };
}

/** A query for the second page, as `ProjectList` builds it. */
const QUERY: IProjectListQuery = {
  page: 2,
  pageSize: 25,
  search: 'renovation',
  sortDirection: 'asc',
  showFavoritesOnly: false,
  domainGroup: 2
};

/**
 * Two rows of `Dashboard/GetListAllWithPagination`, inside the `ResponseDetail` envelope
 * every BMDesk action wraps its result in. Note `recordsTotal` on both rows: the API
 * stamps the collection's total onto each one rather than sending it once.
 */
const API_RESPONSE = {
  message: null,
  success: true,
  messageType: 0,
  data: [
    {
      buildingId: 41,
      encrybuildingId: 'NDE=',
      buildingName: 'Office Renovation Plan',
      isFavorite: true,
      contactPerson: 'Jane Dijkstra',
      maintenanceFolderId: 6,
      isMaintenanceFolderAvailable: true,
      recordsTotal: 137,
      lstDashboard: [
        {
          folderId: 1,
          encryfolderId: 'MQ==',
          folderName: 'Administratie',
          colorCode: '#22b14c',
          borderColorCode: '2.0px solid#d9d9d9',
          isAccessRights: true
        }
      ]
    },
    {
      buildingId: 42,
      encrybuildingId: 'NDI=',
      buildingName: 'Tender Response A12',
      isFavorite: false,
      contactPerson: null,
      maintenanceFolderId: 6,
      isMaintenanceFolderAvailable: false,
      recordsTotal: 137,
      lstDashboard: []
    }
  ]
};

/** A service whose transport answers with `body`, signed in as `app-jwt`. */
function build(
  body: unknown,
  requests: IRecordedRequest[] = [],
  token: string | undefined = 'app-jwt',
  status: number = 200
): IProjectListingService {
  const store = new ApiTokenStore();
  store.setToken(token);

  return new ProjectListingService(
    new ApiService(
      fakeTransport(body, requests, status),
      { baseUrl: 'https://api.example.com/api' },
      store
    )
  );
}

describe('ProjectListingService - binding the overview table end to end', () => {
  it('calls the endpoint the reference controller proxies to', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getProjectPage(QUERY);

    expect(requests[0].url).toBe('https://api.example.com/api/Dashboard/GetListAllWithPagination');
  });

  /** `[HttpPost]` on the API action; a GET would not bind the `PaginationModel` at all. */
  it('posts, as the action requires', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getProjectPage(QUERY);

    expect(requests[0].options.method).toBe('POST');
  });

  it('sends the pagination fields the repository pages and filters on', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getProjectPage(QUERY);

    expect(requests[0].body).toMatchObject({
      Start: '25',
      Length: '25',
      SearchValue: 'renovation',
      SortColumnDirection: 'asc',
      DomainGroup: 2,
      IsShowFavoriteProject: false
    });
  });

  /** The API overwrites both from the JWT, so sending them would be misleading to read. */
  it('does not send the identity fields the API takes from the token', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getProjectPage(QUERY);

    expect(requests[0].body).not.toHaveProperty('UserId');
    expect(requests[0].body).not.toHaveProperty('UserRoleId');
  });

  it('presents the application JWT, so an [Authorize]d endpoint accepts the call', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getProjectPage(QUERY);

    expect((requests[0].options.headers as Record<string, string>).Authorization).toBe('Bearer app-jwt');
  });

  it('unwraps the envelope into the rows the table renders', async () => {
    const page: IProjectListPage = await build(API_RESPONSE).getProjectPage(QUERY);

    expect(page.rows.map((row) => row.buildingName)).toEqual([
      'Office Renovation Plan',
      'Tender Response A12'
    ]);
  });

  /**
   * The count that drives the pager. Two rows came back out of 137; reading the page's own
   * length instead would collapse the table to a single page.
   */
  it('reports the collection total, not the page length', async () => {
    const page: IProjectListPage = await build(API_RESPONSE).getProjectPage(QUERY);

    expect(page.rows).toHaveLength(2);
    expect(page.recordsTotal).toBe(137);
  });

  it('keeps the colours the server computed for each phase cell', async () => {
    const page: IProjectListPage = await build(API_RESPONSE).getProjectPage(QUERY);

    expect(page.rows[0].cells[0]).toMatchObject({
      folderId: 1,
      colorCode: '#22b14c',
      borderColorCode: '2.0px solid#d9d9d9'
    });
  });

  it('reports an empty result as an empty page rather than failing', async () => {
    const page: IProjectListPage = await build({
      message: null,
      success: true,
      messageType: 0,
      data: []
    }).getProjectPage(QUERY);

    expect(page).toEqual({ rows: [], recordsTotal: 0 });
  });

  /**
   * The gotcha this whole layer exists for: the API answers HTTP 200 and reports the
   * failure inside the envelope. Passing the payload straight through would render an
   * empty table, which is indistinguishable from a tenant with no projects.
   */
  it('rejects a refusal the API reports inside an HTTP 200', async () => {
    const service: IProjectListingService = build({
      data: null,
      success: false,
      message: 'Invalid data.',
      messageType: 3
    });

    await expect(service.getProjectPage(QUERY)).rejects.toThrow('Invalid data.');
  });

  it('rejects an HTTP failure', async () => {
    const service: IProjectListingService = build('Server error', [], 'app-jwt', 500);

    await expect(service.getProjectPage(QUERY)).rejects.toThrow();
  });
});
