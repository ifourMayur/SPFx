/**
 * Holds the application JWT the Web API issues at sign-in, so every later call can present
 * it.
 *
 * `POST Authenticate/SPFxLogin` returns a `token` in `ILoginResponse`, and every other
 * controller is `[Authorize]`d and expects it as `Authorization: Bearer <token>`. Rather
 * than threading that string through the component tree - where it would end up in props,
 * state and React DevTools - it is kept here and read by `ApiService` when it builds
 * request headers. `LoginService` writes it on a successful sign-in, so no component ever
 * handles it.
 *
 * The token stays in memory only. It is deliberately never written to `localStorage`,
 * `sessionStorage` or a cookie: a page reload re-runs the SPFx sign-in handshake, which is
 * cheap, whereas a persisted bearer token would outlive the session and be readable by any
 * script on the page.
 */

/** Reads and writes the bearer token used for Web API calls. */
export interface IApiTokenStore {
  /** The current token, or `undefined` when the user is not signed in to the Web API. */
  getToken(): string | undefined;

  /** Stores the token. An empty, blank or `undefined` value clears it. */
  setToken(token: string | undefined): void;

  /** Forgets the token, so subsequent calls are made unauthenticated. */
  clear(): void;
}

/** In-memory {@link IApiTokenStore}. */
export class ApiTokenStore implements IApiTokenStore {
  private _token: string | undefined;

  public getToken(): string | undefined {
    return this._token;
  }

  /**
   * A blank token is stored as `undefined` rather than as an empty string: it would
   * otherwise produce a header of `Authorization: Bearer ` and turn a missing token into a
   * confusing 401 downstream. Same reasoning as `AadAccessTokenProvider`, which treats an
   * empty access token as a failure instead of a value.
   */
  public setToken(token: string | undefined): void {
    const trimmed: string = (token || '').trim();
    this._token = trimmed.length > 0 ? trimmed : undefined;
  }

  public clear(): void {
    this._token = undefined;
  }
}

/**
 * The process-wide token store.
 *
 * One instance per loaded bundle, which is what makes the token "global": `LoginService`
 * writes it once and every `ApiService` instance - including those built for a second web
 * part on the same page - reads the same value. Inject a fresh `ApiTokenStore` instead
 * wherever isolation is wanted, as the unit tests do.
 */
export const apiTokenStore: IApiTokenStore = new ApiTokenStore();
