# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

SharePoint Framework (SPFx v1.23.2) client-side solution named `xs-project` ("XSProject"). It contains a single web part, `XsProjectWebPart` (alias in the manifest: `XsProjectWebPart`), built with React 17 and Fluent UI v8. The web part signs the current Microsoft 365 user in to a separate .NET Core Web API and will render project data from it. The Web API itself is not part of this repository — only the SPFx client that talks to it.

Build tooling is **Heft** (`@rushstack/heft`), not Gulp (`useGulp: false` in `.yo-rc.json`). Config files under `config/` and `tsconfig.json` extend a shared rig package, `@microsoft/spfx-web-build-rig` (see `config/rig.json`), which is where the actual Heft task/phase definitions and Jest config live — `config/*.json` in this repo only overrides pieces of that rig profile.

## Commands

Run from the repo root.

```bash
npm install                 # install dependencies (Node ^22.14.x required, see package.json "engines")
npm run start                # heft start --clean: build-watch + serve, for use with the hosted workbench
npm run build                 # heft test --clean --production && heft package-solution --production
npm run clean                # heft clean
npm run eject-webpack        # heft eject-webpack
heft --help                    # list every available Heft action
```

- `npm run start` serves at the URL/port configured in `config/serve.json` (defaults to the hosted workbench at `https://ifourtechnolab.sharepoint.com/_layouts/workbench.aspx`, port 4321, HTTPS). Debug via the "Hosted workbench" launch config in `.vscode/launch.json`.
- There is no standalone `npm test` script; testing runs through `heft test`, which builds first, then runs Jest via `@rushstack/heft-jest-plugin`. No test files exist in `src/` yet — Jest picks up `*.test.ts`/`*.test.tsx` colocated with source.
- Run a single test file: `heft test --test-path-pattern <regex>` (use `/` for path separators, even on Windows).
- Run tests by name: `heft test --test-name-pattern <regex>` (short flag `-t`).
- Update snapshots: `heft test --update-snapshots` (short flag `-u`).
- Linting runs as part of the `build` phase (via `@microsoft/eslint-config-spfx` / `@microsoft/eslint-plugin-spfx`, configured in `eslint.config.js`) — it isn't a separate script; it runs automatically inside `heft build` / `heft test` / `heft start`.

## Architecture

### Dependency layering

`src/index.ts` documents and enforces (by convention, not tooling) a strict one-way dependency flow:

```
models  ->  config  ->  services  ->  components (src/webparts/*)
```

- **`src/models/`** — transport-agnostic data contracts and pure view-state shapes (`Project`, `Auth`, `ApiError`, `ApiResponse`, `Environment`). Must not import from `config`, `services`, `components`, or any SPFx package, so these stay reusable and unit-testable in isolation.
- **`src/config/environment.ts`** — the single place environment-specific settings live (API base URL, auth resource URIs, UI defaults, per-environment logging). SPFx bundles run in the browser, so there's no `.env`/`process.env`; this typed module is the substitute. `EnvironmentName` is `'local' | 'development' | 'test' | 'production'`, resolved from (in priority order) an explicit override, `isServedFromLocalhost`/`localhost`/a debug build, a host-name-to-environment map (`HOST_NAME_TO_ENVIRONMENT`), then a fallback. Call `initializeEnvironment(...)` once from a web part's `onInit` with SPFx context hints; `getEnvironment()` is safe to call anywhere afterward (it auto-resolves if `initializeEnvironment` hasn't run yet). To onboard a new tenant/host, add an entry to `HOST_NAME_TO_ENVIRONMENT` — no other file needs to change.
- **`src/services/`** — the API/auth layer, each piece behind an interface so components depend on abstractions, not SPFx concretes:
  - `IApiHttpClient` (`ApiHttpClient.ts`) — thin transport abstraction over SPFx's `AadHttpClient` (`AadApiHttpClient`) vs. plain `HttpClient` (`AnonymousApiHttpClient`).
  - `IAccessTokenProvider` (`AccessTokenProvider.ts`) — wraps SPFx's `AadTokenProviderFactory` to acquire the signed-in user's access token; throws `TokenAcquisitionError` (from `models/Auth.ts`) rather than letting SPFx failures leak out raw.
  - `IApiService` / `ApiService` — generic HTTP gateway: URL building (`resolveUrl`), JSON (de)serialization, request timeout via `AbortController`, and translating failed responses into `ApiError` (carries RFC 7807 `ProblemDetails` when the API returns them). All resource-specific services go through this instead of touching `fetch`/SPFx clients directly.
  - `ILoginService` / `LoginService` — the sign-in handshake: acquires an SPFx access token for `auth.tokenResourceUri` (Microsoft Graph by default) and posts it to `POST {baseUrl}/Authenticate/SPFxLogin`. Treats a 200 response carrying `errorMessage` (e.g. a locked account) or missing `token` as a failure, not a success — status code alone doesn't mean the user is signed in.
  - `IProjectService` / `ProjectService` — CRUD against the `ProjectsController` REST endpoints (`GET/POST/PUT/DELETE {baseUrl}/projects[...]`).
  - `ServiceFactory` — the only place concrete service classes are wired up from `BaseComponentContext`. Call `ServiceFactory.getServices(context, overrides?)`; the container is cached per `(instanceId, baseUrl, resourceUri)` so re-renders reuse service instances (and the cached Entra ID token) instead of rebuilding the API layer. `overrides` lets a web part instance replace `baseUrl`/`resourceUri` (e.g. from property pane fields) without touching `src/config/environment.ts`.
- **`src/webparts/xsProject/`** — the SPFx web part and its React components. `XsProjectWebPart.ts` is the only place that calls `ServiceFactory`/`initializeEnvironment`; it injects the resolved `projectService`/`loginService` into `XsProject` (`components/XsProject/XsProject.tsx`) as props, which passes `loginService` down to `Login` (`components/Login/Login.tsx`). Components never import SPFx HTTP/token APIs directly — only the service interfaces from `src/services`. Every component gets its own folder named after it (`components/<Component>/<Component>.tsx`, `I<Component>Props.ts`, `<Component>.module.scss`) — this is a documented convention (see the docblock in `src/index.ts`), not just an existing pattern, so new components should follow it too.

### Auth flow

1. `Login.tsx` runs the handshake on mount (unless `autoLogin={false}`) via `loginService.loginAsCurrentUser()`.
2. `LoginService` asks `AadAccessTokenProvider` for a token scoped to `auth.tokenResourceUri` (defaults to Microsoft Graph, just to prove the caller is a signed-in M365 user — not the Web API's own resource).
3. The token is POSTed to `Authenticate/SPFxLogin` with `domainUrl` (the XSProject app instance the API should associate the session with); the response is an application JWT (`ILoginResponse.token`) for later calls to the Web API — currently just logged to the console (`LoginService._logResponse`) for development verification, not yet stored/used by other services.
4. Any failure (token acquisition, 401, 403, or another API error) is normalized by `resolveLoginFailure` (`models/Auth.ts`) into an `ILoginFailure` with a `kind` (`'tokenAcquisition' | 'authentication' | 'authorization' | 'api'`) and a display-safe `message` — components branch on `kind` (e.g. a 403 doesn't offer a retry button) rather than inspecting status codes themselves.

Access tokens and the application JWT are handled as transient values only: never logged (besides the explicit dev-only console dump), rendered, or persisted to storage outside of what the API response echoes back into component state for display.

### Adding a new API resource

Follow the `Project`/`ProjectService` pattern: define request/response shapes in `src/models/<Resource>.ts` (no imports outside `models/`), add an interface + implementation in `src/services/<Resource>Service.ts` built on `IApiService`, export both from the relevant `index.ts` barrels, then wire the new service into `IServiceContainer` in `ServiceFactory.ts` so components receive it as an injected prop instead of constructing it themselves.

### Multi-environment / multi-tenant packaging

`skipFeatureDeployment: true` and a single `package-solution.json` (`config/package-solution.json`) mean one `.sppkg` is built per `npm run build` and deployed to multiple tenants; per-tenant behavior is driven entirely by `HOST_NAME_TO_ENVIRONMENT` in `src/config/environment.ts`, not by separate packages. The web part also exposes `apiBaseUrl`/`apiResourceUri` property pane fields as a per-instance escape hatch (see `XsProjectWebPart.getPropertyPaneConfiguration`), which only take effect when non-empty.
