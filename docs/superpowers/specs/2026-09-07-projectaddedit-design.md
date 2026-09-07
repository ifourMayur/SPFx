# ProjectAddEdit — design

Date: 2026-09-07
Status: approved and implemented. Supersedes `2026-09-03-addeditproject-design.md`, whose
reference analysis remains accurate but whose four scope decisions do not.

## Goal

Port the reference application's project form,
`D:\Projects\TaskXS\BMDeskV2\BMDesk\Views\Building\AddEdit.cshtml` (3143 lines; markup
68–748, jQuery 754–3143), to React in the `xs-project` SPFx web part, reachable from an
Add button on the project listing and from a per-row Edit action.

Scope decisions taken with the user before design:

| Decision | Choice |
| --- | --- |
| Sections | Core fields, Project Settings, **Tender** and **MOM**. ED Controls, KYP and SharePoint Site URL excluded |
| Data | Initially UI-only. **Amended 2026-09-07:** dropdowns are now bound to the real Web API (see "Dropdown binding" below); saving is still local |
| Routing | **`react-router-dom` v6 with `HashRouter`** — real `/projects`, `/projects/add`, `/projects/edit/:id` paths |
| Listing | Keep the sample rows; add or update the saved project in local state |

## Reference analysis

### The API that exists but is not called

The MVC controller is a thin proxy over a real Web API (`BMDesk.API`), the same one this
web part already signs in to. Recorded here because it is what a later change will wire up:

| Purpose | Endpoint |
| --- | --- |
| Create **and** update | `POST api/Building` — one action; `Id <= 0` adds, otherwise updates |
| Get one project | `GET api/Building?id={id}` |
| Project list | `GET api/Building/GetProjectList` |
| Folder cascade | `GET api/Building/GetProjectFolders` → `…SubFolders` → `…SubSubFolders` → `…SubSubSubFolders` (each takes `projectId`) |
| Project templates | `GET api/ProjectTemplate/GetTemplateList` |
| User list | `GET api/Users/GetAllList` |
| Spatial breakdown | `GET api/SpatialBreakdown/GetAllLookup` |
| Tender templates | `GET api/TenderTemplate/GetTenderTemplateLookupItems` |
| Suppliers | `GET api/Issuer/GetIssuerList` |

Three facts about that API matter for whoever wires it up:

- Every response is wrapped `{ data, success, message, messageType }` (`ResponseDetail`)
  and returns **HTTP 200 even on failure** — callers unwrap `.data` and branch on
  `success`. `ApiService` treats only a non-OK status as an error, so it would read a
  failure as a success. `LoginService` already sets the precedent for handling this.
- It requires `Authorization: Bearer <token>` — the application JWT `SPFxLogin` already
  returns in `ILoginResponse.token`, which nothing consumes yet — plus a `RequestUrl`
  header the auth filter uses to resolve the tenant's database.
- `POST api/Building` dereferences `UserId` and `spatialBreakdownId` without a null check,
  so those arrays must always be sent, and it serializes `BuildingName` as `projectName`
  and `BuldingID` as `projectId` (`[JsonPropertyName]`), while `GET` returns both
  `buildingName`/`project` and `buldingID`/`pId`.

### Validation, from two places

| Field | Rule | Source |
| --- | --- | --- |
| Project name | required | `BuildingAddModel` `[Required]` |
| Address | required | `BuildingAddModel` `[Required]` |
| Postcode | required **unless** domain group is McD | `saveProject` only — the server's `[Required]` is commented out |
| Email | optional; ≤255 and `^[\w-\.]+@([\w-]+\.)+[\w-]{2,}$` | `Constant.EmailRegex`, `[MaxLength(255)]` |
| Phone | optional; `^[\d \+]*$` | `Constant.MobileNumberRegex` |
| Tender template | required when the Tender toggle is on | `saveProject` (line 2377) |

**MOM has no validation.** `getAllMOMData` (line 2431) initialises `allValid = true` and
never clears it, and `#MOMfolderNotSelect` is never written to, so the reference saves a
MOM selection with no folder chosen. The superseded spec claimed otherwise.

`[Required]` on `ClientId` is a no-op: it is a non-nullable `int`, which satisfies the
attribute at `0`.

### Permissions

`BuildingController.Index` (lines 112–140) sets `ViewBag.AddEditAccessRights`, which gates
both the Add button and the per-row Edit link:

- **Allowed** — Admin (1), SuperAdmin (2), Client (3), ClientAdmin (5)
- **Denied** — MaintenanceManager (4), Consultant (7), Franchisenemer (8), McDOffice (10),
  and any unlisted role

`BuildingController.AddEdit` (line 531) is stricter: it returns the error view for a
SuperAdmin, who therefore sees an Add button that leads nowhere. Reproduced deliberately
as two functions rather than silently reconciled.

Every one of these facts is available client-side from the sign-in response
(`userRoleId`, `domainGroup`), so the gates need no extra call.

**One deliberate divergence, found in testing.** Applying the matrix verbatim hid the Add
button from every user. `AuthenticateController.Add` (line 707) hard-codes
`UserRole = (int)UserRoleEnum.MaintenanceManager` for every user it creates, and that is
the path `SPFxLogin` takes - so all SPFx users are MaintenanceManagers, one of the four
roles the reference denies. `MaintenanceManager` is therefore allowed here.

That role means "Uitvoerende partij", an external executing party, in the MVC
application's user population; users of this web part are internal Microsoft 365 users of
the tenant who merely inherit it by accident of provisioning. It grants nothing the API
would refuse either: `BuildingController.Add`'s guard,
`userRoleId != MaintenanceManager || userRoleId != Consultant || ...`, is a tautology (one
integer cannot differ from two values at once), so the API applies no role restriction to
creating a project. Consultant, Franchisenemer and McDOffice stay denied.

The alternative - changing the provisioning at `AuthenticateController.cs:707` so SPFx
users get a project-managing role - is the more faithful fix, but it is a backend change
affecting every tenant.

### Behaviour ported from the jQuery

- Tender template dropdown hidden unless the toggle is on (`getTenderTemplate`).
- MOM cascade hidden unless `isAddMOMPDF` (`#formMomBinding`'s initial state).
- Choosing a folder clears every deeper level and leaves the supplier alone
  (`onChangeFolderDD` / `onChnageSubFolderDD` / `onChnageSubSubFolderDD`, lines 2988–3037).
- Project Template is disabled when `Model.Id > 0`; Reset renders only when `Model.Id == 0`.
- Image upload previews via `URL.createObjectURL` (`loadImage`).
- Success redirects to the listing and shows a toast that auto-hides after 3000 ms.

## Design

### Layering

Follows `src/index.ts` (`models → config → services → components`). The screen is UI-only,
so nothing is added to `config` or `services`; all pure logic lives in `models`, where it
is unit-tested.

### Routing

`react-router-dom` 6.30.6 — the last v6 line, since v7 requires React 18 and this solution
is on React 17 — with `HashRouter`. A SPFx web part is a guest on a SharePoint page whose
path SharePoint itself owns and navigates, so the app's routes live in the fragment, where
they are real, linkable URLs that never cause a page load.

`ROUTE_PATHS` in `src/models/Navigation.ts` declares every path once. `AppView` is removed:
a path is now the only notion of "which view", so there is no second, parallel one to keep
in step.

Hooks are confined to two small function components — `AppShell` and
`ProjectAddEditRoute` — so every other component stays a class, as the repo convention
documents. React Router 6 has no class-component API, and this is the smallest bridge that
does not force the convention to change.

### Shared menu

The existing `Menu` already was the shared menu: one central `MENU_ITEMS` tree, active
highlighting, generic `IMenuItem[]` rendering. It is extended, not duplicated — `view:
AppView` becomes `path: string`, and `IMenuProps` becomes `{ activePath, onNavigate }`.

`AppShell` renders it **once, outside `<Routes>`**, so every page shares it, no page holds
navigation markup, and adding a page cannot forget the menu. This is why the superseded
spec's `VIEWS_WITH_OWN_MENU` guard is unnecessary.

`_isActive` matches an item's own path and anything nested below it, on a `/` boundary, so
"Project" stays highlighted on `/projects/add` and `/projects/edit/12` without
`/projects-archive` counting as nested.

No "Add Project" menu entry: creating a project is permission-dependent, and a menu link
would offer restricted roles a dead end. The gated listing button is the only way in, as
in the reference.

### Where the rows live

`ProjectList` no longer owns its rows. Routing to the form unmounts it, which would discard
a newly saved project, so `XsProject` owns `rows` and the favourite toggle. Search, sort,
paging and the favourites filter stay local to `ProjectList` — they are presentation only.

`XsProject` also keeps `savedForms`, the forms saved this session keyed by project id, so a
project created here can be reopened with its values intact.

### New files

```
src/models/Building.ts                  form shapes, feature gates, empty-form factory
src/models/BuildingValidation.ts        pure validateBuildingForm() / hasErrors()
src/models/FolderTree.ts                cascade options + clear-deeper-levels rules
src/models/Permissions.ts               UserRole, canAddEditProject, canOpenProjectForm
src/models/ProjectListRow.ts            row shape, sample rows, upsertProjectRow()
src/models/{BuildingValidation,FolderTree,Permissions,ProjectListRow}.test.ts
src/webparts/xsProject/components/ProjectAddEdit/
    ProjectAddEdit.tsx  ProjectAddEditRoute.tsx
    IProjectAddEditProps.ts  IProjectAddEditState.ts
    ProjectAddEdit.module.scss  lookups.ts
src/webparts/xsProject/components/FolderCascade/
    FolderCascade.tsx  IFolderCascadeProps.ts  FolderCascade.module.scss
src/webparts/xsProject/components/XsProject/AppShell.tsx
```

### Modified files

- `src/models/Navigation.ts` — `ROUTE_PATHS`, `projectEditPath()`, `IMenuItem.path`;
  `AppView` removed.
- `src/models/index.ts` — export the new modules.
- `components/Menu/{Menu.tsx,IMenuProps.ts}` — path-based navigation and highlighting.
- `components/ProjectList/ProjectList.tsx` (+ `.module.scss`, new `IProjectListProps.ts`) —
  rows and favourites via props, gated Add button and Edit row action, success banner.
- `components/XsProject/{XsProject.tsx,IXsProjectState.ts}` — hosts `HashRouter`, owns rows
  and saved forms, captures the sign-in response for the permission and tenant gates.
- `package.json` — `react-router-dom` 6.30.6.

Removed: the empty `components/AddEditProject/` and `components/TagRow/` folders left by the
superseded spec (`TagRow` existed only for the excluded ED Controls section).

### Validation and error presentation

`validateBuildingForm(form, features)` returns a message per broken rule; `hasErrors`
reports whether saving may proceed. Messages are the reference's own resource strings
verbatim, including the double space in `Project  Name is required`.

Two deliberate differences, both widening what the user is told without changing what
counts as valid: every rule is evaluated so all messages show at once (the reference
returns on the first failure), and the email/phone patterns are checked in the browser
rather than only on the round trip.

### Loading and error states

`onSave` returns a promise. `ProjectAddEdit` awaits it with `isSaving` true — disabling the
actions and showing "Saving…" — and renders a rejection as a form-level error while keeping
the user's input. That is the exact seam where a `BuildingService.save()` call belongs, so
wiring the API in later changes no JSX.

There is deliberately no spinner for loading project data or dropdowns: with no API there is
nothing to wait for, and a fake one would misrepresent the state. Edit mode instead shows an
explicit notice when a project's saved values are not in this session's memory.

## Testing

71 unit tests over the four pure modules, written test-first:

- `BuildingValidation.test.ts` (27) — each required field, postcode enforced except for
  McD, the email length/format and phone format in both directions, tender template in
  both toggle states, MOM proven *not* to be validated, all-errors-at-once, `hasErrors`.
- `FolderTree.test.ts` (16) — options per level, empty when an ancestor is unselected or
  unknown, and which deeper levels each level's change clears, plus the supplier surviving.
- `Permissions.test.ts` (16) — every allowed and denied role, the SuperAdmin discrepancy
  between the two gates, unknown and undefined roles.
- `ProjectListRow.test.ts` (12) — id allocation, append vs. rename in place, and phase
  progress and the favourite star surviving an update.

**Amended during implementation: render tests were added after all.** The plan was to
verify the JSX only through `npm run build`. Two render suites were written instead,
because a build that compiles says nothing about whether the router mounts - and they
immediately earned their place by catching a real defect (see below). They need no new
dependency: `react-dom/test-utils` and the jsdom environment are already there.

- `ProjectAddEdit.test.tsx` (6) - the form mounts with its core fields, Project Template is
  editable and Reset is offered in add mode, the excluded ED Controls / KYP / SharePoint
  sections are genuinely absent, an empty save is blocked with all three required
  messages, a valid save hands over the typed values with both arrays non-null, the MOM
  cascade appears only once its toggle is on with deeper levels disabled, and a
  disallowed role gets the refusal instead of the form.
- `XsProject.test.tsx` (9) - the whole flow against the real component with only the two
  injected services stubbed: sign-in lands on `#/dashboard`; an Admin sees the Add button
  and reaches `#/projects/add`; a MaintenanceManager does not; creating a project returns
  to `#/projects` with the success message, the new row and 18 entries; that project
  reopens for editing with its values intact, template locked and Reset gone; saving the
  edit renames in place and stays at 18; McD needs no postcode and hides Spatial
  Breakdown; a failed validation keeps the user on the form with their input; and a
  SuperAdmin sees the Add button but is refused the form.

**The defect they caught.** New rows were originally appended. The listing pages at ten
rows and the user is returned to the first page after saving, so a newly created project
landed on page two - invisible, which is exactly what requirement 11 ("do not leave the
user on a stale Project Listing screen") is about. `upsertProjectRow` now puts a new row
first. Nothing is contradicted by this: these rows are local sample data, so there is no
server ordering to preserve. An edit still renames in place.

The production bundle is still verified by `npm run build`, which runs ESLint, the
TypeScript build, Jest and the production webpack pass, and writes the `.sppkg`.

## Amendment: authenticated calls and real dropdown binding

Added after the initial implementation, at the user's request.

### The application JWT, held globally

`SPFxLogin` already returned a `token` that nothing consumed. `LoginService` now publishes
it to `apiTokenStore` (`src/services/ApiTokenStore.ts`) — a module-level singleton — as the
last step of a successful sign-in, after the empty-response, `errorMessage` and
missing-token checks, so a failed sign-in never leaves a stale token behind.
`ApiService._buildHeaders` then adds `Authorization: Bearer <token>` to every request while
a token is held, and omits the header entirely when none is, so an unauthenticated call
fails as a clear 401 rather than as a malformed header.

Both the store and the header are injectable (`ApiService`'s third constructor parameter,
`LoginService`'s fifth), defaulting to the singleton, so tests can isolate them.

The token is deliberately **in memory only** — never `localStorage`, `sessionStorage` or a
cookie. A reload re-runs the SPFx handshake, which is cheap; a persisted bearer token would
outlive the session and be readable by any script on the page. No component ever handles it.

**Trap for later:** `resourceUri` must stay unset. It is `undefined` today, so the transport
is `AnonymousApiHttpClient` (plain SPFx `HttpClient`), which forwards our `Authorization`
header untouched. A non-empty `resourceUri` switches to `AadApiHttpClient`, and
`AadHttpClient` sets its *own* `Authorization` header from Entra ID — silently replacing the
application JWT, which the BMDesk API would reject. The two auth models are mutually
exclusive as things stand.

### The `ResponseDetail` envelope

Every action wraps its result as `{ data, success, message, messageType }` and **answers
HTTP 200 even on failure**, so `ApiService` — which only knows status codes — hands a
failure back as a success. `unwrapResponseDetail` (`src/models/ApiEnvelope.ts`) returns
`data` and throws an `ApiError` when `success === false`; a body that is not an envelope
passes straight through, so an endpoint returning a bare array is not mistaken for a
failure. `toLookupOptions` (`src/models/Lookup.ts`) then maps `LookupItem` onto `<option>`
values, dropping unusable entries and rendering a `null` name as an empty label rather than
the string "null". Both are pure and unit-tested, which keeps `LookupService` a thin
mapping of operations onto routes.

### Dropdown binding

`ILookupService` binds each dropdown to the endpoint the reference binds it to. Three come
from the `AddEdit` GET action itself; three are the AJAX handlers the page calls:

| Dropdown | Endpoint | Where the reference calls it |
| --- | --- | --- |
| Project Template | `ProjectTemplate/GetTemplateList` | `AddEdit` GET → `ViewBag.ProjecTemplateList` |
| User List | `Users/GetAllList` | `AddEdit` GET → `ViewBag.AccessRightsUserList` |
| Spatial Breakdown | `SpatialBreakdown/GetAllLookup` | `AddEdit` GET → `ViewBag.spatialBreakdownViewModelsList` |
| Tender Templates | `TenderTemplate/GetTenderTemplateLookupItems` | `getTenderTemplate()` → `Building/GetTenderTemplates` |
| Supplier | `Issuer/GetIssuerList` | `getSupplierList()` → `Building/GetSupplierList` |
| Folder cascade | `Building/GetProjectFolders` + three `…SubFolders` siblings | `getMOMProjectFolderList()` and friends |

The MVC controller is only a proxy — `DoActionForGet<List<LookupItem>>(query, route)` calls
the Web API route named above and returns `ResponseDetail.data` — so calling the API
directly removes a hop rather than changing the data. All six return `List<LookupItem>`,
which is why one mapper serves them all.

The five flat lookups load in parallel on mount, because they are independent endpoints and
sequencing them would make the form wait for the sum of five round trips instead of the
slowest. A single failure fails the whole load and shows one message with a **Try again**
button; a form with some dropdowns silently empty invites saving a half-filled project.

The folder cascade loads one level at a time, on demand, exactly as the reference does. This
replaced the previous local tree walk: `getFolderLevelOptions` is gone, `FolderCascade` now
takes `options` per level plus `loadingLevels` and reports which level changed, and
`clearDeeperLevelOptions` discards the options below a changed parent alongside the
selections `selectFolderLevel` already cleared. Each level shows the reference's own
"Loading..." placeholder while in flight (`setDropdownLoading` in the cshtml). A failed
level is left empty rather than failing the form, since MOM has no validation.

`components/ProjectAddEdit/lookups.ts` — the sample data — is deleted.

## Known limitations

- **Edit prefills only what is in memory.** With no `GET api/Building?id=`, a project
  created this session round-trips fully; the 17 sample rows have only a name, and the
  form says so rather than showing blanks as if they were loaded.
- **Saving is still local.** The dropdowns are real; `POST api/Building` is not wired, so a
  saved project exists only in this session. `onSave` is the seam for it.
- **MOM's folder cascade only populates in edit mode**, because every folder endpoint takes
  a `projectId` and a project being added has none. The dropdowns render but stay disabled,
  matching the reference.
- **`isGrandChildFolderEnabled` is hard-coded true.** The Web API does not expose
  `ViewBag.IsGrandChildFolderEnabled` to this client, so the fourth level always shows.
- **Two web part instances on one page would share the URL fragment**, and so would
  navigate together.
- The SPFx starter boilerplate still renders below the app, and `Login` still renders its
  panel above it — both pre-existing, both flagged in `CLAUDE.md`, neither in scope here.
