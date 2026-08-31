import { AadTokenProvider, AadTokenProviderFactory } from '@microsoft/sp-http';

import { TokenAcquisitionError } from '../models/Auth';

/**
 * Minimal abstraction over the SPFx token providers.
 *
 * {@link LoginService} depends on this interface rather than on the SPFx runtime, which
 * keeps the sign-in flow independent of how the token is obtained and trivially mockable
 * in unit tests - the same reasoning as {@link IApiHttpClient} for the transport.
 */
export interface IAccessTokenProvider {
  /** Resource the token is issued for. */
  readonly resourceUri: string;

  /**
   * Access token of the signed-in Microsoft 365 user.
   *
   * The value is passed straight to the caller: never log it, render it or write it to
   * `localStorage` / `sessionStorage`.
   *
   * @throws TokenAcquisitionError - when no token can be issued for the current user.
   */
  getAccessToken(): Promise<string>;
}

/**
 * Acquires the access token of the signed-in Microsoft 365 user through the SPFx
 * `AadTokenProvider`.
 *
 * SPFx owns the token cache, expiry and refresh, so nothing is cached here beyond the
 * provider instance itself. A failed provider lookup is deliberately not cached, so a
 * retry after a transient failure can succeed.
 */
export class AadAccessTokenProvider implements IAccessTokenProvider {
  private readonly _providerFactory: AadTokenProviderFactory;
  private readonly _resourceUri: string;
  private _providerPromise: Promise<AadTokenProvider> | undefined;

  /**
   * @param providerFactory - `context.aadTokenProviderFactory` from the hosting web part.
   * @param resourceUri - resource to request the token for; comes from
   *   `auth.tokenResourceUri` in `src/config/environment.ts`.
   */
  public constructor(providerFactory: AadTokenProviderFactory, resourceUri: string) {
    this._providerFactory = providerFactory;
    this._resourceUri = resourceUri;
  }

  public get resourceUri(): string {
    return this._resourceUri;
  }

  public async getAccessToken(): Promise<string> {
    let token: string;

    try {
      const provider: AadTokenProvider = await this._getProvider();
      token = await provider.getToken(this._resourceUri);
    } catch (error) {
      const reason: string = error instanceof Error ? error.message : '';
      throw new TokenAcquisitionError(
        `SPFx could not issue an access token for ${this._resourceUri}. ${reason}`.trim(),
        this._resourceUri
      );
    }

    // An empty string would be sent to the API as a valid-looking token and rejected
    // there with a confusing 401, so it is caught as a token failure instead.
    if (!token) {
      throw new TokenAcquisitionError(
        `SPFx returned an empty access token for ${this._resourceUri}. Make sure you are signed in to Microsoft 365.`,
        this._resourceUri
      );
    }

    return token;
  }

  private async _getProvider(): Promise<AadTokenProvider> {
    if (!this._providerPromise) {
      this._providerPromise = this._providerFactory.getTokenProvider().catch((error: unknown): never => {
        // Drop the rejected promise so the next attempt asks SPFx again instead of
        // replaying the failure for the lifetime of the component.
        this._providerPromise = undefined;
        throw error;
      });
    }

    return this._providerPromise;
  }
}
