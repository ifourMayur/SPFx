import { IProblemDetails } from './ApiResponse';

/** Everything known about a failed API call. */
export interface IApiErrorDetails {
  /** HTTP status code, or `0` when the request never reached the server. */
  status: number;
  statusText?: string;
  url?: string;
  problemDetails?: IProblemDetails;
}

/**
 * Error thrown by the service layer for any non-successful API call.
 *
 * Components catch this type to render meaningful messages without knowing
 * anything about HTTP or the transport used to reach the Web API.
 */
export class ApiError extends Error {
  public readonly status: number;
  public readonly statusText: string | undefined;
  public readonly url: string | undefined;
  public readonly problemDetails: IProblemDetails | undefined;

  public constructor(message: string, details: IApiErrorDetails) {
    super(message);

    this.name = 'ApiError';
    this.status = details.status;
    this.statusText = details.statusText;
    this.url = details.url;
    this.problemDetails = details.problemDetails;

    // The compiler targets ES5, where extending a built-in breaks the prototype
    // chain. Restoring it keeps `instanceof ApiError` working at runtime.
    Object.setPrototypeOf(this, ApiError.prototype);
  }

  /** Validation messages reported by `ValidationProblemDetails`, flattened for display. */
  public get validationMessages(): string[] {
    const errors = this.problemDetails?.errors;
    if (!errors) {
      return [];
    }

    const messages: string[] = [];
    Object.keys(errors).forEach((field: string): void => {
      (errors[field] || []).forEach((message: string): void => {
        messages.push(`${field}: ${message}`);
      });
    });

    return messages;
  }
}

/** Type guard so callers can narrow an `unknown` catch variable. */
export function isApiError(error: unknown): error is ApiError {
  return error instanceof ApiError;
}

/**
 * True when the API rejected the call because the caller is not authenticated (HTTP 401).
 *
 * Reported separately from a generic failure: the credentials, not the request, are the
 * problem.
 */
export function isAuthenticationError(error: unknown): boolean {
  return isApiError(error) && error.status === 401;
}

/**
 * True when the API recognised the caller but refused the operation (HTTP 403).
 *
 * Retrying will not help; the account needs to be granted access.
 */
export function isAuthorizationError(error: unknown): boolean {
  return isApiError(error) && error.status === 403;
}

/**
 * Turns any thrown value into a message that is safe to show in the UI.
 *
 * @param error - the value caught by the caller.
 * @param fallback - message used when nothing useful can be extracted.
 */
export function getErrorMessage(error: unknown, fallback: string = 'An unexpected error occurred.'): string {
  if (isApiError(error)) {
    const validationMessages = error.validationMessages;
    return validationMessages.length > 0 ? `${error.message} ${validationMessages.join(' ')}` : error.message;
  }

  if (error instanceof Error) {
    return error.message || fallback;
  }

  return typeof error === 'string' && error.length > 0 ? error : fallback;
}
