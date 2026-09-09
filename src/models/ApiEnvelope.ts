/**
 * The `ResponseDetail` envelope every BMDesk Web API action wraps its result in.
 *
 * Part of the models layer: pure and dependency-free apart from `ApiError`, which is also
 * a model.
 *
 * This exists because the API **answers HTTP 200 even when the operation failed**, putting
 * the real outcome in a `success` flag:
 *
 * ```json
 * { "data": null, "success": false, "message": "...", "messageType": 3 }
 * ```
 *
 * `ApiService` only knows about status codes, so it hands such a response back as a
 * success. Every service that talks to this API must therefore unwrap through
 * {@link unwrapResponseDetail} rather than using the payload directly - otherwise a
 * failure is silently rendered as empty data. `LoginService` already handles its own
 * version of this for the sign-in response.
 */

/* eslint-disable @rushstack/no-new-null -- These interfaces describe the BMDesk Web API's
   wire format, which is the "legacy API" case the rule exempts: it serializes with
   System.Text.Json defaults, so absent values arrive as an explicit `null` rather than
   being omitted. Declaring them as `| null` is what makes the null checks below type-check
   instead of looking redundant. */

import { ApiError } from './ApiError';

/**
 * `DropMessageType` on the server - what kind of message {@link IResponseDetail.message} is.
 *
 * Worth reading on endpoints whose `catch` block answers `success: true`, where this is the
 * only field that still says the call failed. `Document/AddDocumentStorageDetails` is one.
 */
export const MESSAGE_TYPE = {
  success: 0,
  error: 1,
  warning: 2,
  info: 3
};

/** The envelope shape, with every field optional because a failure omits `data`. */
export interface IResponseDetail<T> {
  data?: T | null;
  success?: boolean;
  message?: string;
  /** `DropMessageType` on the server; unused here, kept so the shape is complete. */
  messageType?: number;
}

/** True when `value` is a non-null object, and so could be an envelope. */
function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Returns the payload carried by an API response, or throws when the response reports a
 * failure.
 *
 * @param payload - the parsed response body, exactly as `ApiService` returned it.
 * @param operation - what was being fetched, used in the error message when the API
 *   reports a failure without one of its own (for example "the project templates").
 *
 * A body that is not an envelope is passed straight through, so an endpoint that returns a
 * bare array is not mistaken for a failure. A missing `success` flag counts as success, for
 * the same reason. A successful envelope whose `data` is `null` or absent returns
 * `undefined` - System.Text.Json writes `null` rather than omitting keys, so "absent" and
 * "null" have to mean the same thing.
 *
 * @throws ApiError - when `success` is explicitly `false`. Status `200` is carried on the
 *   error because that genuinely is what the transport saw.
 */
export function unwrapResponseDetail<T>(payload: unknown, operation: string): T | undefined {
  if (!isObject(payload) || !('data' in payload || 'success' in payload)) {
    return payload as T | undefined;
  }

  const envelope: IResponseDetail<T> = payload as IResponseDetail<T>;

  if (envelope.success === false) {
    const message: string = (envelope.message || '').trim();

    throw new ApiError(message || `The Web API could not return ${operation}.`, { status: 200 });
  }

  return envelope.data === null ? undefined : envelope.data;
}
