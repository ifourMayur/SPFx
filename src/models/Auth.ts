import { getErrorMessage, isAuthenticationError, isAuthorizationError } from './ApiError';

/**
 * Data contracts for the `AuthenticateController` of the .NET Core Web API, plus the
 * view state shared by the sign-in components.
 *
 * Part of the models layer: no imports from `src/config`, `src/services` or
 * `src/components`, and no SPFx packages, so the shapes stay reusable and testable.
 */

/**
 * Body sent to `POST {baseUrl}/Authenticate/SPFxLogin`.
 *
 * The access token proves to the Web API that the call really comes from a signed-in
 * Microsoft 365 user. It is acquired at runtime, passed straight through to the request
 * and never persisted, logged or rendered.
 */
export interface ISPFxLoginRequest {
  /** Access token of the signed-in Microsoft 365 user, acquired by SPFx. */
  token: string;
  /** Absolute address of the XSProject application the caller is signing in to. */
  domainUrl: string;
}

/**
 * `BuildingRedirect`: the building to land on when the user has access to exactly one.
 */
export interface IBuildingRedirect {
  buildingId: number;
  isSingleBuilding: boolean;
}

/**
 * `LoginModel`, returned by every login endpoint of the Web API including
 * `POST {baseUrl}/Authenticate/SPFxLogin`.
 *
 * Names are camelCase because the API serializes with the ASP.NET Core System.Text.Json
 * defaults, which also write nulls rather than omitting them: a value the server left unset
 * arrives as an explicit `null`, not as an absent key. Those properties are declared
 * optional here - the house style prefers `undefined` over `null`, and every consumer has
 * to treat "absent", "null" and "empty" the same way regardless.
 *
 * The `Password` property of the server model is deliberately **not** declared here:
 * nothing in the SPFx solution should read it, and `SPFxLogin` now clears it on the server
 * before responding. The other login endpoints still return it to their own clients.
 */
export interface ILoginResponse {
  id: number;
  isExist: boolean;
  /** Application JWT issued by the Web API, for use as the bearer token on later calls. */
  token?: string;
  /** ISO 8601 expiry of `token`, as serialized by System.Text.Json. */
  expiration: string;
  refreshtoken?: string;

  /** Company ("Bedrijf"). */
  bedrijf?: string;
  /** Contact person ("Contactpersoon"). */
  contactpersoon?: string;
  /** Job title ("Functie"). */
  functie?: string;
  /** Phone number ("Telefoon"). */
  telefoon?: string;
  email?: string;
  userName?: string;
  /** Initials used for the avatar placeholder. */
  initial?: string;
  isActive: boolean;
  userRoleId: number;
  languageId: number;
  supplierId: number;
  accountId: number;
  subscriptionsId: number;
  /** Base64 string, which is how System.Text.Json serializes the server's `byte[]`. */
  profilePicture?: string;
  singleBuildingModel?: IBuildingRedirect;

  isNotification: boolean;
  tasksNotification: boolean;
  messagesNotification: boolean;
  approvalsNotification: boolean;
  taskEmailNotification: boolean;
  approveEmailNotification: boolean;

  domainUrl?: string;
  domainGroup: number;
  companyName?: string;
  masterCompanyName?: string;
  userAccessURL?: string;
  userMultipleDomain: number;
  enableAzureSSO: boolean;

  isNewProjectListLayout: boolean;
  isFavoriteProject: boolean;
  isDocumentPreview: boolean;

  /**
   * Business failure reported on an otherwise successful HTTP response, for example a
   * locked account. A 200 with this set does not mean the user is signed in.
   */
  errorMessage?: string;
}

/**
 * Why a sign-in attempt did not succeed.
 *
 * `authentication` and `authorization` correspond to HTTP 401 and 403 and are reported
 * separately from a generic API failure, because they are the user's or the tenant
 * administrator's to resolve rather than a transport problem.
 */
export type LoginFailureKind = 'tokenAcquisition' | 'authentication' | 'authorization' | 'api';

/** A failed sign-in attempt, classified and reduced to a message that is safe to display. */
export interface ILoginFailure {
  kind: LoginFailureKind;
  message: string;
}

/**
 * View state used by components that run the sign-in handshake.
 *
 * The access token is deliberately absent: it lives only for the duration of the service
 * call and never enters component state.
 */
export interface ILoginState {
  isLoading: boolean;
  isAuthenticated: boolean;
  response?: ILoginResponse;
  failure?: ILoginFailure;
}

/**
 * Error thrown when SPFx cannot hand out an access token for the signed-in user.
 *
 * Kept separate from `ApiError` because nothing was sent to the Web API: the request
 * failed before it could be built.
 */
export class TokenAcquisitionError extends Error {
  /** Resource the token was requested for. */
  public readonly resourceUri: string;

  public constructor(message: string, resourceUri: string) {
    super(message);

    this.name = 'TokenAcquisitionError';
    this.resourceUri = resourceUri;

    // The compiler targets ES5, where extending a built-in breaks the prototype
    // chain. Restoring it keeps `instanceof TokenAcquisitionError` working at runtime.
    Object.setPrototypeOf(this, TokenAcquisitionError.prototype);
  }
}

/** Type guard so callers can narrow an `unknown` catch variable. */
export function isTokenAcquisitionError(error: unknown): error is TokenAcquisitionError {
  return error instanceof TokenAcquisitionError;
}

/**
 * Classifies a value caught during sign-in and reduces it to a displayable message.
 *
 * Components call this instead of inspecting status codes themselves, so the 401 / 403
 * distinction is made in exactly one place.
 */
export function resolveLoginFailure(error: unknown): ILoginFailure {
  if (isTokenAcquisitionError(error)) {
    return {
      kind: 'tokenAcquisition',
      message: `Could not obtain your Microsoft 365 access token. ${error.message}`.trim()
    };
  }

  if (isAuthenticationError(error)) {
    return {
      kind: 'authentication',
      message: `XSProject did not accept your Microsoft 365 sign-in. ${getErrorMessage(error)}`.trim()
    };
  }

  if (isAuthorizationError(error)) {
    return {
      kind: 'authorization',
      message: `Your account is not allowed to sign in to XSProject. ${getErrorMessage(error)}`.trim()
    };
  }

  return { kind: 'api', message: getErrorMessage(error, 'Signing in to XSProject failed.') };
}
