import { AadHttpClient, AadHttpClientFactory, HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

/**
 * Minimal transport abstraction over the SPFx HTTP clients.
 *
 * {@link ApiService} depends on this interface rather than on a concrete SPFx client, which
 * keeps the request/response handling logic independent of the authentication scheme and
 * trivially mockable in unit tests.
 */
export interface IApiHttpClient {
  fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse>;
}

/**
 * Calls a Web API secured by Entra ID (Azure AD). SPFx acquires and attaches the
 * bearer token; the client instance is created lazily and cached for reuse.
 */
export class AadApiHttpClient implements IApiHttpClient {
  private readonly _clientFactory: AadHttpClientFactory;
  private readonly _resourceUri: string;
  private _clientPromise: Promise<AadHttpClient> | undefined;

  public constructor(clientFactory: AadHttpClientFactory, resourceUri: string) {
    this._clientFactory = clientFactory;
    this._resourceUri = resourceUri;
  }

  public async fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse> {
    const client: AadHttpClient = await this._getClient();
    return client.fetch(url, AadHttpClient.configurations.v1, options);
  }

  private async _getClient(): Promise<AadHttpClient> {
    if (!this._clientPromise) {
      this._clientPromise = this._clientFactory.getClient(this._resourceUri);
    }

    return this._clientPromise;
  }
}

/** Calls a Web API that does not require an Entra ID token (anonymous or cookie based). */
export class AnonymousApiHttpClient implements IApiHttpClient {
  private readonly _httpClient: HttpClient;

  public constructor(httpClient: HttpClient) {
    this._httpClient = httpClient;
  }

  public async fetch(url: string, options: IHttpClientOptions): Promise<HttpClientResponse> {
    return this._httpClient.fetch(url, HttpClient.configurations.v1, options);
  }
}
