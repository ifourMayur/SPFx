import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Minimal transport abstraction over SharePoint's own REST API.
 *
 * The counterpart of {@link IApiHttpClient} - which talks to the XSProject Web API - for
 * the calls that go to SharePoint itself. Services depend on this interface rather than on
 * `SPHttpClient`, which keeps them free of the SPFx page runtime and so unit-testable
 * under jsdom, where `@microsoft/sp-http` cannot load as a value.
 */
export interface ISpHttpClient {
  /**
   * Issues a `GET` and parses the JSON body.
   *
   * Rejects when SharePoint answers a failure status, so a caller never mistakes an error
   * page for an empty result.
   */
  getJson<T>(url: string): Promise<T>;

  /**
   * Issues a `POST` and parses the JSON body, tolerating an empty one.
   *
   * Rejects on a failure status, carrying SharePoint's own message - which is the only way
   * to tell "a folder of that name already exists" from "you may not write here".
   */
  postJson<T>(url: string, body?: unknown): Promise<T>;
}

/**
 * Calls SharePoint REST as the signed-in user, through SPFx's `SPHttpClient`.
 *
 * SPFx attaches the user's credentials and the `X-RequestDigest` a write needs itself, so
 * there is no token or digest handling here - which is also why this needs no app
 * registration and no admin consent, and why every response is security-trimmed to what
 * the current user may see.
 */
export class SpRestHttpClient implements ISpHttpClient {
  private readonly _client: SPHttpClient;

  public constructor(client: SPHttpClient) {
    this._client = client;
  }

  public async getJson<T>(url: string): Promise<T> {
    const response: SPHttpClientResponse = await this._client.get(
      url,
      SPHttpClient.configurations.v1,
      SpRestHttpClient._options()
    );

    return SpRestHttpClient._readJson<T>(response);
  }

  public async postJson<T>(url: string, body?: unknown): Promise<T> {
    const response: SPHttpClientResponse = await this._client.post(
      url,
      SPHttpClient.configurations.v1,
      SpRestHttpClient._options(body)
    );

    return SpRestHttpClient._readJson<T>(response);
  }

  /**
   * `odata=nometadata` is asked for explicitly so responses arrive as plain JSON rather
   * than wrapped in the verbose `d`/`results` envelope.
   */
  private static _options(body?: unknown): ISPHttpClientOptions {
    return {
      headers: { Accept: 'application/json;odata=nometadata' },
      body: body === undefined ? undefined : JSON.stringify(body)
    };
  }

  /**
   * Reads the body once, as text, then decides what it was.
   *
   * Text first because the body is needed on both paths and a response may only be read
   * once: on failure it carries SharePoint's error message, and on success it may be empty
   * (some writes answer `204`), which `response.json()` would reject on.
   */
  private static async _readJson<T>(response: SPHttpClientResponse): Promise<T> {
    let text: string = '';

    try {
      text = await response.text();
    } catch {
      text = '';
    }

    if (!response.ok) {
      throw new Error(
        SpRestHttpClient._toErrorMessage(text) ||
          `SharePoint request failed: ${response.status} ${response.statusText}`
      );
    }

    if (!text) {
      return undefined as unknown as T;
    }

    try {
      return JSON.parse(text) as T;
    } catch {
      return text as unknown as T;
    }
  }

  /**
   * Pulls the message out of a SharePoint REST error body.
   *
   * The shape depends on the metadata level: `odata=nometadata` answers
   * `{ "error": { "message": { "value": "..." } } }`, while other levels use the
   * `odata.error` key. Both are read, because the header above is a request, not a
   * guarantee, and a message like "The folder ... already exists" is what makes a failed
   * folder create diagnosable at all.
   */
  private static _toErrorMessage(text: string): string {
    if (!text) {
      return '';
    }

    let payload: { [key: string]: unknown };
    try {
      payload = JSON.parse(text) as { [key: string]: unknown };
    } catch {
      return '';
    }

    const error: unknown = payload.error || payload['odata.error'];
    const message: unknown = error && typeof error === 'object'
      ? (error as { [key: string]: unknown }).message
      : undefined;

    if (typeof message === 'string') {
      return message;
    }

    const value: unknown = message && typeof message === 'object'
      ? (message as { [key: string]: unknown }).value
      : undefined;

    return typeof value === 'string' ? value : '';
  }
}
