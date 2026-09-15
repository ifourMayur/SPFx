# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

SharePoint Framework (SPFx v1.23.2) client-side solution named `xs-project` ("XSProject"). It contains a single web part, `XsProjectWebPart`, built with React 17, Fluent UI v8 and `react-router-dom` v6. The web part signs the current Microsoft 365 user in to a separate .NET Core Web API ("BMDesk.API") and renders/edits project data from it. Neither the Web API nor the reference ASP.NET MVC application is part of this repository — only the SPFx client.

Much of the work here is a **port of that reference MVC application** (`D:\Projects\TaskXS\BMDeskV2`). Source comments cite reference files, controllers, resource strings and `ViewBag` flags by name, and deliberately record where this implementation diverges and why. Preserve that reasoning when editing — several "obvious corrections" are documented mistakes (see `src/models/Permissions.ts`).

Approved design specs live in [docs/superpowers/specs/](docs/superpowers/specs/); `2026-09-07-projectaddedit-design.md` supersedes `2026-09-03-addeditproject-design.md` and is the current spec for the project form.

Build tooling is **Heft** (`@rushstack/heft`), not Gulp (`useGulp: false` in `.yo-rc.json`). Config files under `config/` and `tsconfig.json` extend a shared rig package, `@microsoft/spfx-web-build-rig` (see `config/rig.json`), which is where the actual Heft task/phase definitions and Jest config live — `config/*.json` in this repo only overrides pieces of that rig profile.

## Commands

Run from the repo root.

```bash
npm install                  # install dependencies (Node ^22.14.x required, see package.json "engines")
npm run start                # heft start --clean: build-watch + serve, for use with the hosted workbench
npm run build                # heft test --clean --production && heft package-solution --production
npm run clean                # heft clean
npm run eject-webpack        # heft eject-webpack
heft test                    # build, then run Jest
heft --help                  # list every available Heft action
```

- `npm run start` serves at the URL/port configured in `config/serve.json` (defaults to the hosted workbench at `https://ifourtechnolab.sharepoint.com/_layouts/workbench.aspx`, port 4321, HTTPS). Debug via the "Hosted workbench" launch config in `.vscode/launch.json`.
- There is no `npm test` script; testing runs through `heft test`, which builds first, then runs Jest via `@rushstack/heft-jest-plugin`. Results land in `jest-output/` (JUnit XML + coverage, collected on every run).
- **Jest runs against compiled output, never `src`.** The rig points Jest's `roots`/`testMatch` at `lib-commonjs/**/*.test.js`, and the `test` phase depends on `build`. Two consequences: a TypeScript or lint error anywhere fails the build before a single test runs (so "no tests ran" usually means the build broke, not that the tests are missing), and path patterns match the **compiled** path — `--test-path-pattern services/BuildingService` matches, anything anchored on `src/` matches nothing. Run `heft clean` after renaming or deleting a test, or its stale compiled copy keeps running.
- Run a single test file: `heft test --test-path-pattern <regex>` (use `/` for path separators, even on Windows).
- Run tests by name: `heft test --test-name-pattern <regex>` (short flag `-t`).
- Update snapshots: `heft test --update-snapshots` (short flag `-u`). `--disable-code-coverage` speeds up a focused run; `--detect-open-handles`, `--silent`, `--max-workers` and `--test-timeout-ms` are also available.
- `npm run build` writes the package to `sharepoint/solution/xs-project.sppkg` (`paths.zippedPackage` in `config/package-solution.json`); `npm run start` writes debug manifests under `sharepoint/solution/debug/`. Every generated directory (`lib`, `lib-commonjs`, `dist`, `temp`, `release`, `sharepoint`, `jest-output`) is gitignored.
- Linting runs inside the `build` phase (`@microsoft/eslint-config-spfx` react flat profile, see `eslint.config.js`) — it isn't a separate script; it runs automatically inside `heft build` / `heft test` / `heft start`.
- `core.autocrlf` is `true` with no `.gitattributes`, so most `git` commands print `LF will be replaced by CRLF` warnings. That's noise, not a pending change — don't "fix" the line endings.

## Architecture

### Dependency layering

`src/index.ts` documents and enforces (by convention, not tooling) a strict one-way dependency flow:

```
models  ->  config  ->  services  ->  components (src/webparts/*)
```

- **`src/models/`** — transport-agnostic data contracts, pure view-state shapes and pure functions (validation, permission rules, folder-tree planning, route table). Must not import from `config`, `services`, `components`, or any SPFx package, so these stay reusable and unit-testable under plain jsdom. Most of the repo's test coverage lives at this layer.
- **`src/config/environment.ts`** — the single place environment-specific settings live (API base URL, controller routes, auth resource URIs, UI defaults). SPFx bundles run in the browser, so there's no `.env`/`process.env`; this typed module is the substitute. `EnvironmentName` is `'local' | 'development' | 'test' | 'production'`, resolved from (in priority order) an explicit override, `isServedFromLocalhost`/`localhost`/a debug build, a host-name-to-environment map (`HOST_NAME_TO_ENVIRONMENT`), then a fallback. Call `initializeEnvironment(...)` once from a web part's `onInit`; `getEnvironment()` is safe to call anywhere afterward (it auto-resolves if `initializeEnvironment` hasn't run yet). `api.endpoints` holds **controller routes, not full endpoints** — services append their own action (`Authenticate` + `SPFxLogin`, `Users` + `GetAllList`) and `ApiService.resolveUrl` prefixes `baseUrl`. To onboard a new tenant/host, add an entry to `HOST_NAME_TO_ENVIRONMENT` — no other file needs to change.
- **`src/services/`** — the API/auth layer, each piece behind an interface so components depend on abstractions, not SPFx concretes. See "Services" below.
- **`src/webparts/xsProject/`** — the SPFx web part and its React components. `XsProjectWebPart.ts` is the only place that calls `ServiceFactory`/`initializeEnvironment`; it injects the resolved services into `XsProject` as props. Components never import SPFx HTTP/token APIs directly — only the service interfaces from `src/services`.

### Two transports, two APIs

The solution talks to two different back ends, and keeping them apart is deliberate:

- **The XSProject Web API** — `IApiHttpClient` (`ApiHttpClient.ts`) over SPFx's `AadHttpClient` (`AadApiHttpClient`) or plain `HttpClient` (`AnonymousApiHttpClient`), wrapped by `IApiService`.
- **SharePoint itself** — `ISpHttpClient` (`SharePointHttpClient.ts`) over SPFx's `SPHttpClient`. Runs as the signed-in user, so it needs no app registration, no client secret and no tenant-admin API approval, and every response is security-trimmed. Requests `odata=nometadata` and reads bodies as text first (a body can only be read once, and SharePoint's own error message is the only way to tell "folder already exists" from "you may not write here").

### The `ResponseDetail` envelope — HTTP 200 can mean failure

**The most important gotcha in the codebase.** Every BMDesk Web API action wraps its result in `{ data, success, message, messageType }` and answers **HTTP 200 even when the operation failed**. `ApiService` only knows status codes, so it hands such a response back as a success.

Every service that talks to this API must unwrap through `unwrapResponseDetail` (`src/models/ApiEnvelope.ts`) rather than using the payload directly — otherwise a refused save is silently rendered as empty data. A body that is not an envelope passes straight through (so a bare array isn't mistaken for a failure), and a missing `success` flag counts as success. `LoginService` handles its own version of the same problem.

### Services

- `IApiService` / `ApiService` — generic HTTP gateway: URL building (`resolveUrl`), JSON (de)serialization, request timeout via `AbortController`, the `Authorization: Bearer` header from the token store, and translating failed responses into `ApiError` (carrying RFC 7807 `ProblemDetails` when present).
- `IAccessTokenProvider` (`AccessTokenProvider.ts`) — wraps SPFx's `AadTokenProviderFactory`; throws `TokenAcquisitionError` (from `models/Auth.ts`) rather than letting SPFx failures leak out raw. Treats an empty token as a failure, not a value.
- `IApiTokenStore` / `apiTokenStore` — holds the application JWT the Web API issues at sign-in. `LoginService` writes it; `ApiService` reads it when building headers. **In-memory only, module-level singleton, never threaded through props/state** (where it would surface in React DevTools) and deliberately never written to `localStorage`/`sessionStorage`/a cookie — a reload re-runs the cheap SPFx handshake. Inject a fresh `ApiTokenStore` for isolation, as the tests do.
- `ILoginService` / `LoginService` — the sign-in handshake (see "Auth flow").
- `IProjectService` / `ProjectService` — the REST-style `api/projects` the dashboard reads.
- `IBuildingService` / `BuildingService` — the project add/edit **write** path against `api/Building` (a `ProjectController` alias). Distinct from `IProjectService`; don't merge them. See "The project save" below.
- `IProjectTemplateService` — what a project template defines, currently its folder tree (`ProjectTemplate/GetSubFolder`).
- `ILookupService` / `LookupService` — reference data the project form's dropdowns bind to. Each method maps onto the same endpoint the reference MVC view used, calling the Web API directly instead of via the MVC proxy.
- `ISharePointFolderService` — creates a project's folder tree in the chosen site's default document library and reports the id of each folder; the SPFx-native replacement for the reference's app-only-Graph `SharePointHelper.CreateProjectFolders`. **A folder has two unrelated ids.** `AddUsingPath` answers with SharePoint's `UniqueId` (a GUID), but the Web API reaches these folders through Graph, which addresses items by an opaque drive-scoped `driveItem.id` (`01VL4HET…`) — a different identifier, not a different encoding, so there is no local conversion between them. The Graph id is read per folder from `{siteUrl}/_api/v2.0/drive/root:/{path}`, the same drive API Graph serves but hosted by the site, so `SPHttpClient` can call it as the signed-in user with no app registration or admin consent. That read failing is not a folder failure: the folder exists, its `id` is empty, and `toDocumentStorageModel` then leaves it out rather than posting an id that would never resolve.
- `IDocumentStorageService` — tells the Web API which SharePoint folder each of a project's folders became (`Document/AddDocumentStorageDetails`). Its `catch` block answers `success: true` with `messageType: Error`, so the envelope's own flag says "recorded" on the one response that means nothing was written — the service checks `messageType` as well.
- `ISharePointSiteService` — the site collections the user may work in, via the SharePoint **Search** API (`contentclass:STS_Site`, one page of 500) rather than Graph `/sites`, again to avoid app registration and admin consent.
- `ServiceFactory` — the only place concrete service classes are wired up from `BaseComponentContext`. Call `ServiceFactory.getServices(context, overrides?)`; the container is cached per `(instanceId, baseUrl, resourceUri)` so re-renders reuse service instances (and the cached Entra ID token). `overrides` lets a web part instance replace `baseUrl`/`resourceUri` from property pane fields without touching `src/config/environment.ts`. Note the two separately-resolved URLs: `_resolveDomainUrl` (the app identity sent as `domainUrl` at sign-in) and `_resolveWebUrl` (the site whose `_api` endpoint SharePoint calls are addressed to) — their fallbacks differ, so don't collapse them.

### Auth flow

1. `Login.tsx` runs the handshake on mount (unless `autoLogin={false}`) via `loginService.loginAsCurrentUser()`.
2. `LoginService` asks `AadAccessTokenProvider` for a token scoped to `auth.tokenResourceUri` (Microsoft Graph by default, just to prove the caller is a signed-in M365 user — not the Web API's own resource).
3. The token is POSTed to `Authenticate/SPFxLogin` with `domainUrl`. The response (`ILoginResponse`) carries the application JWT, plus `userRoleId`, `domainGroup` and user details the UI needs — so no extra call is needed for permissions or tenant behaviour. `LoginService` writes the JWT into `apiTokenStore` and hands the rest to the component.
4. A 200 response carrying `errorMessage` (e.g. a locked account) or missing `token` is a **failure**, not a success — status code alone doesn't mean the user is signed in.
5. Any failure is normalized by `resolveLoginFailure` (`models/Auth.ts`) into an `ILoginFailure` with a `kind` (`'tokenAcquisition' | 'authentication' | 'authorization' | 'api'`) and a display-safe `message` — components branch on `kind` (e.g. a 403 doesn't offer a retry button) rather than inspecting status codes themselves.

Tokens are transient values only: never rendered, never persisted, and not logged beyond the explicit dev-only response dump in `LoginService._logResponse`.

### Rendering and routing

`XsProject.tsx` (class component) gates the app in **two steps** before the router appears: `Login` must report success, then `SharePointSites` must report the site the user chose to work in. Only then does it render `HashRouter` → `AppShell`.

- Routes live in the **URL fragment** (`.../workbench.aspx#/projects/add`). A SPFx web part is a guest on a SharePoint page whose path SharePoint owns and navigates; `HashRouter` keeps the app's routes real and linkable without triggering a page load.
- The route table is declared once in `src/models/Navigation.ts` (`ROUTE_PATHS`, `DEFAULT_ROUTE_PATH`, `projectEditPath`) — models layer, so the router, the shared `Menu` and every page agree without depending on each other. Build concrete edit paths with `projectEditPath(id)` rather than interpolating `:id`.
- `AppShell.tsx` renders `Menu` **once, outside `<Routes>`**, so adding a page cannot forget the navigation. Unrecognised paths `Navigate ... replace` to the default.
- **There are two project surfaces, not one.** `/dashboard` renders `Project`, which reads live data through `IProjectService.getProjects()`; `/projects` renders `ProjectList`, the reference's overview table, whose rows are still `SAMPLE_PROJECT_ROWS` (`models/ProjectListRow.ts`) because `GetProjectList` exposes no per-phase status. Don't conflate the two when wiring data in.
- `XsProject` owns those listing rows and the forms saved during the session, because routing unmounts the page that would otherwise hold them — replace the sample rows with a service call following the `Project`/`ProjectService` pattern once the data is available.
- `XsProject.render` still ends with the Yeoman generator's boilerplate ("Well done, …" plus the SPFx links), *below* the gated app. It's leftover scaffolding, not part of the ported UI.

### The project save

`BuildingService.saveProject` fans out to several calls whose **order is a business rule**, which is why it lives in the service and not a component: read the project back (update only) → `POST api/Building` (routes on `model.Id`, so `0` inserts) → read the template's folder tree → create those folders in SharePoint (insert only, and only with a chosen site) → `POST api/Building` **again** with the root folder's id bound on (the "rebind") → read the project's **own** folder tree → `POST Document/AddDocumentStorageDetails` with the SharePoint ids bound onto it.

**The folder tree is read twice on purpose.** `ProjectTemplate/GetSubFolder` decides *what to create*: it reports the template as authored and handles the synthetic `BIM` folder. `Building/GetProjectsubFolder?...&isCreateProject=true` decides *what to report*: `DocumentStorageDetailsService.AddList` stamps the SharePoint ids onto `ProjectSubFolder` rows matched by id, and only that read knows them — the template read answers with `SubFolder.ID` instead, which would match nothing or, on a numeric collision, write an id onto the wrong sub-folder of the project. Driving creation from the project-scoped read as the reference does would create every active folder in the system rather than the template's, so the two stay separate. The binding itself is `toDocumentStorageModel` in `src/models/DocumentStorage.ts`, which matches nodes **by path** (a template may define both `Design/Plans` and `Build/Plans`) and drops any folder that was not provisioned — `AddList` reads a top-level folder's storage type as `(int)item.DocumentStorageType` with no null check, so an unbound node would fault the whole call.

Everything pure about that fan-out lives in `src/models/BuildingSave.ts` — `toBuildingSaveRequest` (form → wire), `toPreservedFields`, `toRebindRequest`, and the `IProjectSaveResult` the UI reads — so the shapes can be tested without a transport and `BuildingService` holds only the ordering.

Two rules to preserve when touching it:

- An **update is refused** when the read-back comes back empty. `ProjectController.Update` assigns the image, SharePoint folder id and third-party fields from the request *unconditionally*, so a POST without them blanks the image and orphans the folders (see `IPreservedProjectFields`). The reference has no such guard because its Razor view round-trips them through hidden fields.
- The folder steps are **never allowed to fail the save**. The project genuinely exists by then, so reporting failure would invite a duplicate. What couldn't be done comes back in the result instead (absent `templateFolders`, a `folderProvision` carrying failures, a `rebind` saying the project was not given its folder id, or a `documentStorage` saying the ids were not recorded).

**The project is posted twice on an insert, and the second POST is not a retry.** A project cannot be told which SharePoint folder it lives in until both exist, and the folders cannot be created until the insert has assigned a project id to name them under — each value is unknowable at the moment the other is needed. So the insert posts `id: 0` with no folder, and `toRebindRequest` re-sends **the whole body** with the assigned id (which routes the call to `Update`) and the root folder's Graph `driveItem.id`. Re-sending the whole body is the point: `Update` assigns unconditionally, so a two-field patch would blank the image and storage type the insert just stored — the same trap `IPreservedProjectFields` exists for. It is skipped when there is nothing to bind: an update, no chosen site, or a root folder whose Graph id couldn't be read.

The save reports itself on the listing as a **single self-dismissing toast** carrying the reference's own `msgAddProject`/`msgUpdateProject` text, and nothing else. Folders that *failed* to be created are the one exception and stay named in it — the folder steps never fail the save, so that toast is the only signal the user gets that a folder they will look for isn't there. What the Web API answered, including the SharePoint ids bound onto the folder tree, is a development detail: `BuildingService` logs it, the UI doesn't show it.

Still not implemented from the reference: tender folders.

### The project form

`ProjectAddEdit` is the largest component here and the port of `Views/Building/AddEdit.cshtml`. Its rules deliberately live in the models layer, so they can be tested without rendering anything:

- `models/Building.ts` — `IBuildingForm` (everything the form edits) and `createEmptyBuildingForm`. `id: 0` means insert, matching the single `POST api/Building` endpoint. Field names keep the reference's Dutch-origin spellings (`plaats`, `contactpersoon`, `functie`, `telefoon`) so the mapping back to `BuildingModel`/`BuildingAddModel` stays obvious — the visible labels are English; don't rename the fields to match them.
- `models/BuildingValidation.ts` — every rule from both places the reference enforces them (`BuildingAddModel` data annotations server-side, `saveProject` in the Razor view client-side). All rules are evaluated on each attempt rather than returning on the first failure, so the user isn't fixing one field per round trip.
- `models/FolderTree.ts` — the `Folder → SubFolder → SubSubFolder → SubSubSubFolder` cascade: which levels to discard when a parent changes, keeping `FolderCascade` purely presentational. `ILookupService` performs the four level-by-level fetches the reference does.
- `FolderCascade` and `MultiSelect` are the form's two shared controls.

**Message strings are copied verbatim from the reference's resource files, double spaces included** — `'Project  Name is required'`, `'Project  Added successfully!'`. They look like typos and aren't; a ported screen is meant to read identically to the one it replaces.

ED Controls (tag rows, ticket source) and KYP planning are absent on purpose: third-party integration sections, excluded from this screen.

### Permissions and per-tenant behaviour

Both are derived from the sign-in response, in the models layer:

- `src/models/Permissions.ts` — the `UserRole` enum plus `canAddEditProject` / `canOpenProjectForm`, evaluated against `ILoginResponse.userRoleId`. It **deliberately grants `MaintenanceManager`**, which the reference denies, because `AuthenticateController.Add` hard-codes that role for every user `SPFxLogin` creates — denying it left no SPFx user able to open the form at all. Read the docblock before "fixing" this. A UI check is a convenience, never a security boundary; the Web API remains the authority.
- `src/models/Building.ts` — `toDomainGroup` maps the numeric `domainGroup` onto `'mcD' | 'taskXs' | 'xsWallet'`, which drives the surviving server-side gates (McD skips the postcode rule and hides Spatial Breakdown; TaskXs alone shows the ISO 19650 toggle). An unknown tenant grouping falls back to the strictest, most featureful treatment rather than silently losing a required field.

### Component conventions

Every component gets its own folder named after it: `components/<Component>/<Component>.tsx`, plus `I<Component>Props.ts` / `I<Component>State.ts` and `<Component>.module.scss` as needed (a few components declare their props interface inline in the `.tsx` instead). This is a documented convention (see the docblock in `src/index.ts`), not just an existing pattern.

Components are **class components**. There are exactly two function components, both intentional and both documented in place, because React Router 6 exposes location/params/navigate through hooks only:

- `AppShell.tsx` — the routed shell.
- `ProjectAddEditRoute.tsx` — a thin route adapter that reads `:id` and `navigate` and hands them to the `ProjectAddEdit` class as plain props, so the form itself knows nothing about routing.

### Testing conventions

Tests are colocated as `*.test.ts`/`*.test.tsx`. Two jsdom constraints shape every test that touches a service or component:

- `jest.mock('@microsoft/sp-core-library', ...)` (stubbing `Log`) **must appear above the imports**. The TypeScript build emits CommonJS without Babel's `jest.mock` hoisting, so the mock has to be registered before the module graph is required.
- `@microsoft/sp-http` **cannot load as a value** under jsdom. Import from it with `import type` only, and declare a local fake-transport interface (e.g. `IFakeHttpClient`) rather than importing `IApiHttpClient` from a module that pulls it in as a value.

Service tests run the real `ApiService` over a fake transport answering the exact JSON the Web API answers, because the interesting failures (HTTP 200 on a refused save, a field silently nulled) are invisible from the outside. Component tests drive the real component tree with stubbed service interfaces and unmount every `ReactDOM.render` in `afterEach`.

### Adding a new API resource

Follow the `Project`/`ProjectService` pattern: define request/response shapes in `src/models/<Resource>.ts` (no imports outside `models/`), add an interface + implementation in `src/services/<Resource>Service.ts` built on `IApiService` — unwrapping through `unwrapResponseDetail` — add the controller route to `ENDPOINTS` in `src/config/environment.ts`, export both from the relevant `index.ts` barrels, then wire the new service into `IServiceContainer` in `ServiceFactory.ts` so components receive it as an injected prop instead of constructing it themselves.

### Adding a new page

Three coordinated edits and nothing else: a path in `ROUTE_PATHS` (`src/models/Navigation.ts`), a `<Route>` in `AppShell.tsx`, and an entry in `MENU_ITEMS` (`Menu.tsx`) — a leaf for a top-level item, or a `children` entry for a submenu. `Menu` is data-driven and rendered once by the shell, so no page carries navigation markup. `Document`, `Search` and `Suppliers` are "Coming soon" placeholders in exactly this shape, waiting on a backing endpoint.

### Multi-environment / multi-tenant packaging

`skipFeatureDeployment: true` and a single `config/package-solution.json` mean one `.sppkg` is built per `npm run build` and deployed to multiple tenants; per-tenant behavior is driven entirely by `HOST_NAME_TO_ENVIRONMENT` in `src/config/environment.ts`, not by separate packages. The web part also exposes `apiBaseUrl`/`apiResourceUri` property pane fields as a per-instance escape hatch (see `XsProjectWebPart.getPropertyPaneConfiguration`), which only take effect when non-empty.

`auth.tokenResourceUri` is Microsoft Graph today, so the API is called anonymously and `resourceUri` is `undefined` in every environment. Pointing it at an Application ID URI also requires a matching `webApiPermissionRequests` entry in `config/package-solution.json`.
