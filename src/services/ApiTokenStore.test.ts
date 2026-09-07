import { ApiTokenStore, apiTokenStore } from './ApiTokenStore';

describe('ApiTokenStore', () => {
  it('starts with no token', () => {
    expect(new ApiTokenStore().getToken()).toBeUndefined();
  });

  it('returns the token it was given', () => {
    const store = new ApiTokenStore();
    store.setToken('jwt-abc');

    expect(store.getToken()).toBe('jwt-abc');
  });

  it('replaces an earlier token, so a second sign-in wins', () => {
    const store = new ApiTokenStore();
    store.setToken('first');
    store.setToken('second');

    expect(store.getToken()).toBe('second');
  });

  it('forgets the token on clear, so signing out stops authenticating calls', () => {
    const store = new ApiTokenStore();
    store.setToken('jwt-abc');
    store.clear();

    expect(store.getToken()).toBeUndefined();
  });

  it('treats undefined as clearing the token', () => {
    const store = new ApiTokenStore();
    store.setToken('jwt-abc');
    store.setToken(undefined);

    expect(store.getToken()).toBeUndefined();
  });

  // An empty or whitespace-only token would produce `Authorization: Bearer ` and a
  // confusing 401. It is treated as no token at all, the same way `AadAccessTokenProvider`
  // treats an empty access token as a failure rather than a value.
  it('treats an empty token as no token', () => {
    const store = new ApiTokenStore();
    store.setToken('');

    expect(store.getToken()).toBeUndefined();
  });

  it('treats a whitespace-only token as no token', () => {
    const store = new ApiTokenStore();
    store.setToken('   ');

    expect(store.getToken()).toBeUndefined();
  });

  it('trims a token that arrived with surrounding whitespace', () => {
    const store = new ApiTokenStore();
    store.setToken('  jwt-abc  ');

    expect(store.getToken()).toBe('jwt-abc');
  });
});

describe('apiTokenStore', () => {
  afterEach(() => {
    apiTokenStore.clear();
  });

  it('is a single shared instance, so a token set once is visible everywhere', () => {
    apiTokenStore.setToken('shared-jwt');

    // Re-imported rather than re-constructed: this is the whole point of the singleton.
    expect(apiTokenStore.getToken()).toBe('shared-jwt');
  });

  it('is independent of any freshly constructed store', () => {
    apiTokenStore.setToken('shared-jwt');

    expect(new ApiTokenStore().getToken()).toBeUndefined();
  });
});
