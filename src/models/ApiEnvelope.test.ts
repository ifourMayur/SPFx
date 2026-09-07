import { isApiError } from './ApiError';
import { unwrapResponseDetail } from './ApiEnvelope';

describe('unwrapResponseDetail', () => {
  it('returns the payload from inside the envelope', () => {
    const result = unwrapResponseDetail<number[]>({ data: [1, 2], success: true }, 'lookups');

    expect(result).toEqual([1, 2]);
  });

  // The API answers HTTP 200 even when the operation failed, so the flag - not the status
  // code - is what decides. ApiService cannot see this; only the caller can.
  it('throws when the envelope reports failure', () => {
    expect(() =>
      unwrapResponseDetail({ data: undefined, success: false, message: 'Account is locked.' }, 'lookups')
    ).toThrow('Account is locked.');
  });

  it('throws an ApiError, so callers can narrow it like any other API failure', () => {
    let thrown: unknown;
    try {
      unwrapResponseDetail({ success: false, message: 'Nope.' }, 'lookups');
    } catch (error) {
      thrown = error;
    }

    expect(isApiError(thrown)).toBe(true);
  });

  it('reports the operation when a failure carries no message', () => {
    expect(() => unwrapResponseDetail({ success: false }, 'the project templates')).toThrow(
      /the project templates/
    );
  });

  it('treats a whitespace-only message as no message', () => {
    expect(() => unwrapResponseDetail({ success: false, message: '   ' }, 'the suppliers')).toThrow(
      /the suppliers/
    );
  });

  // System.Text.Json writes `null` rather than omitting a key, so an absent collection and
  // an explicit null must behave the same way.
  it('returns undefined for a successful envelope whose data is null', () => {
    expect(unwrapResponseDetail({ data: null, success: true }, 'lookups')).toBeUndefined();
  });

  it('returns undefined for a successful envelope with no data key at all', () => {
    expect(unwrapResponseDetail({ success: true }, 'lookups')).toBeUndefined();
  });

  // Not every endpoint is guaranteed to wrap its result; a bare payload is passed through
  // rather than being mistaken for a failure.
  it('passes through a payload that is not wrapped in an envelope', () => {
    expect(unwrapResponseDetail<number[]>([1, 2, 3], 'lookups')).toEqual([1, 2, 3]);
  });

  it('treats a missing success flag as success', () => {
    expect(unwrapResponseDetail<string>({ data: 'ok' }, 'lookups')).toBe('ok');
  });

  it('returns undefined for an empty response body', () => {
    expect(unwrapResponseDetail(undefined, 'lookups')).toBeUndefined();
  });
});
