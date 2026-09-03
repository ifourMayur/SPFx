# AddEditProject component — design

Date: 2026-09-03
Status: approved, ready for implementation planning

## Goal

Add an `AddEditProject` view to the `xs-project` SPFx web part that reproduces the
BMDesk MVC project form at
`D:\Projects\TaskXS\BMDeskV2\BMDesk\Views\Building\AddEdit.cshtml` (3143 lines; markup
68–748, jQuery 754–3143), and reach it from the shared `Menu`.

Scope decisions taken with the user before design:

| Decision | Choice |
| --- | --- |
| Fidelity | **Full** — every section of the cshtml, including ED Controls tag rows, KYP planning, Tender, SharePoint URL and MOM |
| Data | **UI-only, local React state** — no POST, no new service; dropdowns read in-memory lookup data |
| Menu | **Both** — a `MENU_ITEMS` entry *and* `AddEditProject` renders its own `<Menu>` |
| Controls | **Plain HTML + SCSS module**, matching `Login` / `Menu` / `ProjectList` |

Out of scope: any Web API call, `IProjectService` integration, a `BuildingService`, and
the 14 `/Building/*` AJAX endpoints the cshtml uses. Those are a later change; this one is
a faithful, self-contained UI.

## Reference analysis

### Form model

The cshtml binds `BMDesk.Business.ViewModels.Building.BuildingModel` and posts to
`Building/AddEdit` as `multipart/form-data`, with hidden fields for `TagData`, `MOMData`,
`TicketSourceId`, `PlanningSourcetId`, `ImageBytes` and (edit only) `Id`.

### Sections, in render order

1. **Header** — Back button, "Project" title.
2. **Core fields** — image upload button with live preview (`ImageFile` → `loadImage`),
   then two-per-row: Project Name \*, Project ID · Postcode \*, Address \* · City
   (`Plaats`), Contactpersoon · Functie, Email · Telefoon, Client (read-only text +
   hidden `ClientId`) · Project Template (locked when editing), User List (multi-select)
   · Spatial Breakdown (multi-select, hidden for the McD domain group).
3. **ED Controls / Audits** — rendered only when `EDControlsConfigurationModel` has
   ClientId + ClientSecretId + Username + Password. Contains the `IsTicketSource` toggle
   and a `TicketProjectId` dropdown; when editing (`Model.Id > 0`) it also renders
   repeatable tag rows — Periodic checkbox, Document Name (revealed by Periodic), Tag
   multi-select, an "+" add-row button on the first row only, and a folder cascade group
   of Folder → SubFolder → SubSubFolder → SubSubSubFolder (the last gated on
   `ViewBag.IsGrandChildFolderEnabled`) plus Supplier.
4. **Planning / KYP** — rendered only when `KYPProjectConfigurationModel.AuthToken` is
   set. `IsPlanningSource` toggle, `IsUseClientKYPToken` checkbox, `KypProjectAuthToken`
   password field, an Authorize button, and a `PlanningProjectId` dropdown.
5. **Project Settings** — Maintenance folder, BIM folder, and (TaskXs domain group only)
   ISO-19650 code toggles. The cshtml duplicates this whole block across the
   `DocumentStorageType == XSCloud` branch and its `else`; the two branches are
   identical in the parts that survive into this design.
6. **Tender** — `IsEnableTender` toggle revealing a required `TenderTemplateId` dropdown.
7. **SharePoint** — `SharePointSiteURL` text field, when `Model.IsShowShareSiteSiteURL`.
8. **MOM** — `isAddMOMPDF` toggle revealing its own folder cascade + supplier.
9. **Actions** — Save (client-side validation, then submit) and, in add mode only, Reset.

### Initial visibility rules (`$(document).ready`, lines 814–967)

- KYP token box and the "use alternative token" checkbox start hidden; the planning
  project dropdown starts hidden.
- Tender template dropdown hidden unless `Model.IsEnableTender`.
- ED tag-row container hidden; existing `TagData` rows are rendered instead when present.
- MOM cascade hidden unless `Model.isAddMOMPDF`.
- Toggling `IsPlanningSource` on shows the alt-token checkbox (and the token box when
  alt-token is checked); toggling it off hides both.
- `Periodic` reveals / hides Document Name (`readOnlyTextBox`, line 3040).

### Save validation order (`saveProject`, line 2323)

1. Clear prior messages.
2. Postcode required — **skipped** for the McD domain group.
3. ED Controls project required when `IsTicketSource` is checked.
4. Planning project required when `IsPlanningSource` is checked.
5. Abort if a SharePoint auth error is displayed.
6. Collect tag rows (`getAllTagData`, line 2395): only rows with at least one tag are
   kept; `validateAllRowsOnEdit` requires a tag *and* a folder per row.
7. Tender template required when `IsEnableTender` is checked.
8. Collect MOM rows (`getAllMOMData`, line 2431).
9. Submit.

Adding a row (`addNewRow`, line 1105) refuses to append until every existing row
validates.

## Design

### Layering

Follows the documented flow in `src/index.ts` (`models → config → services →
components`). Because the view is UI-only, nothing is added to `config` or `services`.

### New files

```
src/models/Building.ts                 form/view-state shapes, no outside imports
src/models/BuildingValidation.ts       pure validateBuildingForm(), the unit under test
src/models/BuildingValidation.test.ts  Jest, colocated per CLAUDE.md
src/webparts/xsProject/components/AddEditProject/
    AddEditProject.tsx
    IAddEditProjectProps.ts
    IAddEditProjectState.ts
    AddEditProject.module.scss
    lookups.ts                         sample dropdown/folder-tree data
src/webparts/xsProject/components/FolderCascade/
    FolderCascade.tsx  IFolderCascadeProps.ts  FolderCascade.module.scss
src/webparts/xsProject/components/TagRow/
    TagRow.tsx  ITagRowProps.ts  TagRow.module.scss
```

`FolderCascade` is extracted because the cascade appears twice (per tag row, and in MOM);
`TagRow` because it repeats N times and owns the Periodic → Document Name rule. Both are
controlled components — value in, `onChange` out, no state of their own.
Each gets its own folder, per the convention documented in `src/index.ts`.

### Modified files

- `src/models/Navigation.ts` — add `'addEditProject'` to `AppView`.
- `src/models/index.ts` — export the new `Building` / `BuildingValidation` members.
- `components/Menu/Menu.tsx` — add
  `{ key: 'addEditProject', label: 'Add / Edit Project', view: 'addEditProject' }`
  to the `children` of the `project` item.
- `components/XsProject/XsProject.tsx` — add the render branch, and a
  `VIEWS_WITH_OWN_MENU: AppView[] = ['addEditProject']` guard so the outer `<Menu>` is
  skipped for views that render their own. `AddEditProject` receives `activeView` and
  `onNavigate` and renders `<Menu>` itself, mirroring the cshtml's
  `@Html.Partial("MenuBar")`. Exactly one sidebar is on screen at a time.

`XsProject.module.scss` is **not** modified: for a view that owns its menu, `XsProject`
renders the component directly as a child of `.xsProject` rather than wrapping it in
`.appLayout`, and `AddEditProject` supplies its own flex row (`.layout`) holding `<Menu>`
plus the form.

### Model shapes (`src/models/Building.ts`)

```ts
export type DomainGroup = 'taskXs' | 'mcD' | 'other';

export interface ILookupOption { id: string; name: string; }

export interface IFolderSelection {
  folderId: string; subFolderId: string; subSubFolderId: string;
  subSubSubFolderId: string; supplierId: string;
}

export interface ITagRow extends IFolderSelection {
  key: string;            // stable React key, not part of the payload
  periodic: boolean;
  documentName: string;
  tags: string[];
}

export interface IBuildingForm {
  id: number;
  imageDataUrl?: string;
  buildingName: string; buildingId: string; postcode: string; address: string;
  plaats: string; contactpersoon: string; functie: string; email: string;
  telefoon: string; clientName: string; clientId: string;
  projectTemplateId: string; userIds: string[]; spatialBreakdownIds: string[];
  isTicketSource: boolean; ticketProjectId: string; tagRows: ITagRow[];
  isPlanningSource: boolean; isUseClientKypToken: boolean;
  kypProjectAuthToken: string; planningProjectId: string;
  isMaintenanceFolderAvailable: boolean; isEnableBimFolder: boolean;
  isEnableIsoFormat: boolean;
  isEnableTender: boolean; tenderTemplateId: string;
  sharePointSiteUrl: string;
  isAddMomPdf: boolean; momFolder: IFolderSelection;
}
// The reference serializes MOM as a one-element `MOMData` array but only ever renders a
// single row, so a single `IFolderSelection` is the honest shape here.

/** Stands in for the cshtml's server-side @if gates. All default to true/'other'. */
export interface IProjectFormFeatures {
  domainGroup: DomainGroup;
  isEdControlsConfigured: boolean;
  isKypConfigured: boolean;
  isGrandChildFolderEnabled: boolean;
  isShowSharePointSiteUrl: boolean;
}

export interface IBuildingFormErrors {
  buildingName?: string; postcode?: string; address?: string;
  ticketProjectId?: string; planningProjectId?: string; tenderTemplateId?: string;
  momFolderId?: string;
  /** Keyed by `ITagRow.key`. */
  tagRows?: Record<string, { tags?: string; folderId?: string }>;
}
```

Field names keep the reference's Dutch-origin spellings (`plaats`, `contactpersoon`,
`functie`, `telefoon`) so the mapping back to `BuildingModel` stays obvious — the same
precedent `ProjectList` sets by keeping Dutch phase-column labels. Visible labels are
English ("City", "Contact person", "Role", "Phone").

### Validation (`src/models/BuildingValidation.ts`)

```ts
export function validateBuildingForm(
  form: IBuildingForm, features: IProjectFormFeatures
): IBuildingFormErrors;
export function hasErrors(errors: IBuildingFormErrors): boolean;
export function validateTagRows(rows: ITagRow[]): IBuildingFormErrors['tagRows'];
```

Pure, dependency-free, in the models layer so it is unit-testable in isolation. Rules, in
the reference's order: `buildingName` and `address` required; `postcode` required unless
`domainGroup === 'mcD'`; `ticketProjectId` required when `isTicketSource`;
`planningProjectId` required when `isPlanningSource`; `tenderTemplateId` required when
`isEnableTender`; every tag row needs at least one tag and a `folderId`; `momFolder`
needs a `folderId` when `isAddMomPdf`. `validateTagRows` is exported separately because
the "+" add-row button runs exactly that check before appending, matching
`validateAllRowsOnEdit`.

### Lookup data

A `LOOKUPS` constant in a colocated `AddEditProject/lookups.ts` holds sample options for project templates, users, spatial breakdowns,
ED projects, planning projects, tender templates, tags, suppliers, and a nested folder
tree:

```ts
interface IFolderNode extends ILookupOption { children?: IFolderNode[]; }
```

`FolderCascade` derives each level's options from the level above by walking that tree, and
clears the deeper levels when a parent changes — the observable behavior of
`getProjectSubFolderList` / `getProjectSubSubFolderList` / `getProjectSubSubSubFolderList`
without the network. This mirrors how `ProjectList` renders `SAMPLE_ROWS` today, and each
lookup carries a comment naming the `/Building/*` endpoint that will replace it.

### Component behavior

`AddEditProject` is a class component (matching every other component here) holding one
`IBuildingForm` in state plus `IBuildingFormErrors` and a `statusMessage`.

- Handlers are assigned as instance properties, following the `Login` / `Menu` pattern; a
  single `_onFieldChange` keyed off `data-field` covers the text inputs rather than one
  handler per field.
- Toggle-driven visibility is derived at render time from the form state, reproducing the
  jQuery show/hide rules listed above — including the KYP token box appearing only when
  planning *and* alt-token are both on, and Document Name appearing only with Periodic.
- The "+" button appends a `ITagRow` only when `validateTagRows` returns no errors;
  otherwise it sets the per-row errors. Consistent with the reference, the button renders
  on the first row only.
- Save runs `validateBuildingForm`, renders the messages next to their fields, and on
  success sets a saved-status message carrying the collected payload shape. No request is
  made; a comment marks where a `BuildingService` call will go.
- Reset restores the initial form state instead of `location.reload()` — reloading the
  hosting SharePoint page would discard the whole web part. Reset is hidden when
  `projectId > 0`, and Project Template is disabled then, matching `@if (Model.Id > 0)`.
- Image upload sets a preview via `URL.createObjectURL`, as `loadImage` does, and revokes
  the URL on unmount / replacement.
- A Back button above the title calls `onNavigate('projectList')`, standing in for the
  reference's `#backButton`.

**One deliberate divergence.** The reference renders the ED Controls tag rows only when
`Model.Id > 0`, so in add mode they are invisible. Since this view defaults to add mode
and the agreed scope is full fidelity, `AddEditProject` renders the tag rows whenever ED
Controls is configured and the `isTicketSource` toggle is on, in both add and edit mode —
otherwise a whole section of the form could never be seen or reviewed. Everything else
follows the reference's gating exactly.

### Props (`IAddEditProjectProps`)

`activeView: AppView`, `onNavigate: (view: AppView) => void` (both for the embedded
`<Menu>`), `projectId?: number` (0/undefined = add mode), and
`features?: Partial<IProjectFormFeatures>` merged over an all-sections-visible default so
the whole form is reviewable in the workbench.

### Styling

`AddEditProject.module.scss` uses the `"[theme:token, default:…]"` + `var(--token)` pattern
the existing modules use, reuses ProjectList's toggle-switch styles, ports
`.folderCascadeGroup` from the cshtml's inline `<style>` (grey panel, 6px radius, 12px
uppercase 600-weight labels), and lays out the `row formProject` / `col-lg-6` pairs as a
responsive two-column grid collapsing to one column on narrow widths. Required-field
asterisks and `text-danger` messages get their own classes rather than Bootstrap's.

## Testing

- `src/models/BuildingValidation.test.ts` is written first (TDD) and covers: each required
  field; postcode skipped for McD and enforced otherwise; the three toggle-conditional
  requirements in both toggle states; tag-row validation with a missing tag, a missing
  folder, and a valid row; MOM folder required only when the toggle is on; `hasErrors` on
  an empty object.
- The JSX is verified by `npm run build` completing clean (Heft runs ESLint and the
  TypeScript build), plus a look at the form in the hosted workbench via `npm run start`.
- No component-level render tests: the repo has none today, and this change is not the
  place to introduce that infrastructure.

## Risks and follow-ups

- The lookup data is sample data. Every dropdown is wired to shapes that a future
  `BuildingService` can fill without touching the JSX, but until then Save proves only the
  validation path.
- `IProject` in `src/models/Project.ts` covers a much smaller field set than
  `BuildingModel`. Reconciling the two (or adding `BuildingService` against the
  `/Building` endpoints) is deliberately left for a separate change.
- `VIEWS_WITH_OWN_MENU` introduces one layout special case in `XsProject`. If more views
  come to own their menu, the guard is the single place that changes; if none do, it can
  fold back into a plain condition.
