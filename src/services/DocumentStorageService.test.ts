/**
 * End-to-end tests for reporting provisioned folders back to the Web API:
 * `DocumentStorageService` over the real `ApiService`, against a fake transport answering
 * the exact JSON `Document/AddDocumentStorageDetails` answers.
 *
 * The real one is worth running because this endpoint reports its failures in two
 * incompatible ways inside an HTTP 200 - see the `messageType` test below - and a client
 * that reads only the status code would record "folders saved" for a call that wrote
 * nothing.
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

import { IDocumentStorageRequest } from '../models/DocumentStorage';
import { ApiService } from './ApiService';
import { ApiTokenStore } from './ApiTokenStore';
import { DocumentStorageService, IDocumentStorageService } from './DocumentStorageService';

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

/** Minimal stand-in for an SPFx `HttpClientResponse`. */
function fakeResponse(body: unknown): HttpClientResponse {
  return {
    ok: true,
    status: 200,
    statusText: 'OK',
    text: async (): Promise<string> => JSON.stringify(body)
  } as unknown as HttpClientResponse;
}

function fakeTransport(body: unknown, requests: IRecordedRequest[]): IFakeHttpClient {
  return {
    fetch: async (url: string, options: IHttpClientOptions): Promise<HttpClientResponse> => {
      requests.push({ url, options });
      return fakeResponse(body);
    }
  };
}

function build(body: unknown, requests: IRecordedRequest[] = []): IDocumentStorageService {
  const store = new ApiTokenStore();
  store.setToken('app-jwt');

  return new DocumentStorageService(
    new ApiService(fakeTransport(body, requests), { baseUrl: 'https://api.example.com/api' }, store),
    'Document'
  );
}

/** A bound model, as `toDocumentStorageModel` produces one. */
const MODEL: IDocumentStorageRequest = {
  projectId: 77,
  documentStorageType: 1,
  sharePointFolderId: 'guid:Roof',
  templateID: 4,
  lstProjectFolderDetailsViewModel: [
    {
      foldersId: 11,
      folderName: 'Drawings',
      documentStorageType: 1,
      sharePointFolderId: 'guid:Roof/Drawings',
      lstProjectsubFolderViewModel: []
    }
  ]
};

/** The envelope a successful call answers with: the model echoed back. */
const ACCEPTED = { data: MODEL, success: true, message: '', messageType: 0 };

describe('DocumentStorageService', () => {
  it('posts to the action the reference posts to', async () => {
    const requests: IRecordedRequest[] = [];
    await build(ACCEPTED, requests).addDocumentStorageDetails(MODEL);

    expect(requests[0].url).toBe('https://api.example.com/api/Document/AddDocumentStorageDetails');
    expect(requests[0].options.method).toBe('POST');
  });

  it('sends the bound model as the request body', async () => {
    const requests: IRecordedRequest[] = [];
    await build(ACCEPTED, requests).addDocumentStorageDetails(MODEL);

    expect(JSON.parse(requests[0].options.body as string)).toEqual(MODEL);
  });

  it('resolves when the API accepts the call', async () => {
    await expect(build(ACCEPTED).addDocumentStorageDetails(MODEL)).resolves.toBeUndefined();
  });

  it('rejects when the API reports a failure in the envelope', async () => {
    const refused = { data: null, success: false, message: 'Invalid data.', messageType: 1 };

    await expect(build(refused).addDocumentStorageDetails(MODEL)).rejects.toThrow('Invalid data.');
  });

  it('rejects when the API reports an error while still claiming success', async () => {
    // `DocumentController.AddDocumentStorageDetails` catches its own exception and answers
    // `success: true` with `messageType: Error` - so the envelope's own flag says the
    // folders were recorded when nothing was written at all.
    const faulted = {
      data: MODEL,
      success: true,
      message: 'Something went wrong, please try after sometime',
      messageType: 1
    };

    await expect(build(faulted).addDocumentStorageDetails(MODEL)).rejects.toThrow(
      'Something went wrong, please try after sometime'
    );
  });

  it('accepts a message the API sends alongside a success', async () => {
    const chatty = { data: MODEL, success: true, message: 'Saved.', messageType: 0 };

    await expect(build(chatty).addDocumentStorageDetails(MODEL)).resolves.toBeUndefined();
  });
});
