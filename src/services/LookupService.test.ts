/**
 * End-to-end binding tests for the dropdown data: `LookupService` over the real
 * `ApiService`, against a fake transport returning the exact JSON the Web API returns.
 *
 * These exist because stubbing `ILookupService` (as the component tests do) skips the
 * layer most likely to break a dropdown - URL building, the `ResponseDetail` unwrap and the
 * `LookupItem` mapping. A fault in any of those renders an empty dropdown with no error,
 * which is indistinguishable from "the API returned nothing".
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

import { ILookupOption } from '../models/Building';
import type { IApiEndpoints } from '../models/Environment';
import { ApiService } from './ApiService';
import { IAccessTokenProvider } from './AccessTokenProvider';
import { ApiTokenStore } from './ApiTokenStore';
import { LoginService } from './LoginService';
import { ILookupService, LookupService } from './LookupService';

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
  url: string;
  options: IHttpClientOptions;
}

/**
 * Minimal stand-in for an SPFx `HttpClientResponse`.
 *
 * Only the members `ApiService` touches are implemented - `ok`, `status`, `statusText` and
 * `text()` - so the real response type does not have to be constructed.
 */
function fakeResponse(body: unknown, status: number = 200): HttpClientResponse {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? 'OK' : 'Error',
    text: async (): Promise<string> => (typeof body === 'string' ? body : JSON.stringify(body))
  } as unknown as HttpClientResponse;
}

/** Transport that answers every call with `body` and records what it was asked. */
function fakeTransport(
  body: unknown,
  requests: IRecordedRequest[],
  status: number = 200
): IFakeHttpClient {
  return {
    fetch: async (url: string, options: IHttpClientOptions): Promise<HttpClientResponse> => {
      requests.push({ url, options });
      return fakeResponse(body, status);
    }
  };
}

/**
 * The exact envelope the Web API returns: `ResponseDetail` camelCased by System.Text.Json,
 * wrapping `List<LookupItem>`. Written out literally so a change in the real shape shows up
 * here as a failure.
 */
const API_RESPONSE = {
  message: null,
  success: true,
  data: [
    { id: 1, strId: null, name: 'Base template', description: null, selected: false },
    { id: 2, strId: null, name: 'Renovation', description: null, selected: false }
  ],
  messageType: 0
};

/** A service whose transport answers with `body`, signed in as `app-jwt`. */
function build(body: unknown, requests: IRecordedRequest[] = []): ILookupService {
  return buildWith(body, requests, 'app-jwt');
}

/** As {@link build}, but with an explicit token - omit it for a signed-out caller. */
function buildWith(
  body: unknown,
  requests: IRecordedRequest[],
  token?: string,
  status: number = 200
): ILookupService {
  const store = new ApiTokenStore();
  store.setToken(token);

  return new LookupService(
    new ApiService(
      fakeTransport(body, requests, status),
      { baseUrl: 'https://api.example.com/api' },
      store
    )
  );
}

describe('LookupService - binding a dropdown end to end', () => {
  it('turns the real API envelope into dropdown options', async () => {
    const options: ILookupOption[] = await build(API_RESPONSE).getProjectTemplates();

    expect(options).toEqual([
      { id: '1', name: 'Base template' },
      { id: '2', name: 'Renovation' }
    ]);
  });

  it('calls the same route the reference controller proxies to', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getProjectTemplates();

    expect(requests[0].url).toBe('https://api.example.com/api/ProjectTemplate/GetTemplateList');
  });

  it('presents the application JWT, so an [Authorize]d endpoint accepts the call', async () => {
    const requests: IRecordedRequest[] = [];
    await build(API_RESPONSE, requests).getUsers();

    expect((requests[0].options.headers as Record<string, string>).Authorization).toBe('Bearer app-jwt');
  });

  it('omits the header entirely when no one has signed in', async () => {
    const requests: IRecordedRequest[] = [];
    await buildWith(API_RESPONSE, requests).getUsers();

    expect((requests[0].options.headers as Record<string, string>).Authorization).toBeUndefined();
  });

  it('reaches each of the five flat lookup routes', async () => {
    const requests: IRecordedRequest[] = [];
    const service: ILookupService = build(API_RESPONSE, requests);

    await service.getProjectTemplates();
    await service.getUsers();
    await service.getSpatialBreakdowns();
    await service.getTenderTemplates();
    await service.getSuppliers();

    expect(requests.map((request: IRecordedRequest) => request.url)).toEqual([
      'https://api.example.com/api/ProjectTemplate/GetTemplateList',
      'https://api.example.com/api/Users/GetAllList',
      'https://api.example.com/api/SpatialBreakdown/GetAllLookup',
      'https://api.example.com/api/TenderTemplate/GetTenderTemplateLookupItems',
      'https://api.example.com/api/Issuer/GetIssuerList'
    ]);
  });

  it('sends the folder cascade parameters the API declares', async () => {
    const requests: IRecordedRequest[] = [];
    const service: ILookupService = build(API_RESPONSE, requests);

    await service.getProjectSubSubSubFolders(12, 10, 11, 13);

    // `subfolderId` and `subsubfolderId` are lower-cased in the API signature, unlike
    // `folderId` - getting that wrong would bind an empty dropdown.
    expect(requests[0].url).toBe(
      'https://api.example.com/api/Building/GetProjectSubSubSubFolders' +
        '?projectId=12&folderId=10&subfolderId=11&subsubfolderId=13'
    );
  });

  // The API answers 200 with `success: false`, so a failure must not arrive as an empty
  // dropdown - that would look like "the data just isn't there".
  it('raises a failure reported inside a 200 response', async () => {
    const failure = { message: 'Account is not active.', success: false, data: null, messageType: 3 };

    await expect(build(failure).getUsers()).rejects.toThrow('Account is not active.');
  });

  it('raises a transport failure', async () => {
    const service: ILookupService = buildWith({ message: 'Unauthorized' }, [], 'app-jwt', 401);

    await expect(service.getUsers()).rejects.toThrow();
  });

  it('returns no options for a successful but empty collection', async () => {
    const empty = { message: null, success: true, data: [], messageType: 0 };

    expect(await build(empty).getSuppliers()).toEqual([]);
  });
});

describe('LookupService - endpoint composition', () => {
  // The routes come from `src/config/environment.ts`, not from constants inside the
  // service, so renaming a controller is a one-line change there. This proves the
  // configured value is what is actually used.
  it('builds the URL from the configured controller route plus the action', async () => {
    const requests: IRecordedRequest[] = [];
    const store = new ApiTokenStore();
    const endpoints: IApiEndpoints = {
      projects: 'projects',
      authenticate: 'Authenticate',
      building: 'RenamedBuilding',
      users: 'RenamedUsers',
      projectTemplates: 'ProjectTemplate',
      spatialBreakdowns: 'SpatialBreakdown',
      tenderTemplates: 'TenderTemplate',
      suppliers: 'Issuer'
    };

    const service = new LookupService(
      new ApiService(
        fakeTransport(API_RESPONSE, requests),
        { baseUrl: 'https://api.example.com/api' },
        store
      ),
      endpoints
    );

    await service.getUsers();
    await service.getProjectFolders(7);

    expect(requests[0].url).toBe('https://api.example.com/api/RenamedUsers/GetAllList');
    expect(requests[1].url).toBe(
      'https://api.example.com/api/RenamedBuilding/GetProjectFolders?projectId=7'
    );
  });
});

describe('the application token issued at sign-in', () => {
  /** A login response carrying `token`, which is all these tests need from it. */
  const LOGIN_RESPONSE = {
    id: 42,
    isExist: true,
    token: 'jwt-from-login',
    expiration: '2030-01-01T00:00:00Z',
    userRoleId: 4,
    userName: 'mayur@ifourtechnolab.com'
  };

  function tokenProvider(): IAccessTokenProvider {
    return {
      resourceUri: 'https://graph.microsoft.com',
      getAccessToken: async (): Promise<string> => 'entra-access-token'
    };
  }

  // The point of the shared store: nobody passes the token around, yet every later call
  // presents it.
  it('is presented on a lookup call made after a successful login', async () => {
    const store = new ApiTokenStore();
    const loginRequests: IRecordedRequest[] = [];
    const lookupRequests: IRecordedRequest[] = [];
    const configuration = { baseUrl: 'https://api.example.com/api' };

    const loginService = new LoginService(
      new ApiService(fakeTransport(LOGIN_RESPONSE, loginRequests), configuration, store),
      tokenProvider(),
      'Authenticate',
      'https://contoso.sharepoint.com',
      store
    );
    const lookupService: ILookupService = new LookupService(
      new ApiService(fakeTransport(API_RESPONSE, lookupRequests), configuration, store)
    );

    // Nothing is signed in yet, so nothing is presented.
    expect(store.getToken()).toBeUndefined();

    await loginService.loginAsCurrentUser();
    await lookupService.getUsers();

    expect((lookupRequests[0].options.headers as Record<string, string>).Authorization).toBe(
      'Bearer jwt-from-login'
    );
  });

  it('is composed the same way login composes its own endpoint', async () => {
    const store = new ApiTokenStore();
    const requests: IRecordedRequest[] = [];

    const loginService = new LoginService(
      new ApiService(
        fakeTransport(LOGIN_RESPONSE, requests),
        { baseUrl: 'https://api.example.com/api' },
        store
      ),
      tokenProvider(),
      'Authenticate',
      'https://contoso.sharepoint.com',
      store
    );

    await loginService.loginAsCurrentUser();

    // baseUrl + controller + action - the pattern the lookups now follow.
    expect(requests[0].url).toBe('https://api.example.com/api/Authenticate/SPFxLogin');
  });

  it('is not stored when the sign-in fails', async () => {
    const store = new ApiTokenStore();
    const loginService = new LoginService(
      new ApiService(
        fakeTransport({ errorMessage: 'Your account is locked.' }, []),
        { baseUrl: 'https://api.example.com/api' },
        store
      ),
      tokenProvider(),
      'Authenticate',
      'https://contoso.sharepoint.com',
      store
    );

    await expect(loginService.loginAsCurrentUser()).rejects.toThrow('Your account is locked.');
    expect(store.getToken()).toBeUndefined();
  });
});
