/**
 * The project overview table's wire contract and its presentation rules.
 *
 * Part of the models layer: pure, with no imports outside `models`, so the awkward parts
 * of this endpoint can be tested without a transport.
 *
 * The table is the port of `Views/Building/_ProjectListingScreen.cshtml`, whose DataTable
 * runs with `serverSide: true` and posts to `BuildingController.GetAllProjectWithPagination`.
 * That action is a **proxy and nothing more**: it decodes DataTables' form-encoded
 * `draw` / `start` / `length` / `order[0][dir]` / `search[value]` into a `PaginationModel`,
 * stamps `DomainGroup` on from the MVC session, and POSTs it to the real endpoint,
 * `Dashboard/GetListAllWithPagination`. There is no DataTables here and no MVC session, so
 * this module builds that same `PaginationModel` directly and the proxy disappears - the
 * same hop `LookupService` already skips for the dropdowns.
 *
 * Two details of that contract are invisible from the outside and wrong by default:
 *
 * - **`UserId` and `UserRoleId` must not be sent.** `DashboardController` assigns both
 *   from the caller's JWT before the repository sees them, so a value sent here is
 *   discarded - and reads as though the client chose whose projects to list.
 * - **The row count is not sent beside the rows.** `GetDashboardListWithPagination`
 *   stamps `RecordsTotal` onto *every* row, and the reference reads it off
 *   `data.FirstOrDefault()`. Take the page's length for the total instead and the pager
 *   silently believes there is exactly one page.
 */

/* eslint-disable @rushstack/no-new-null -- The `IApi*` interfaces below describe the
   BMDesk Web API's wire format, which is the "legacy API" case the rule exempts: it
   serializes with System.Text.Json defaults, so absent values arrive as an explicit
   `null` rather than being omitted. `| null` is what makes the guards below type-check
   instead of looking redundant. Same reasoning as `models/Lookup.ts`. */

/** Page size used when a caller asks for a nonsensical one. Matches the reference's default. */
const DEFAULT_PAGE_SIZE: number = 25;

/**
 * The column the reference's DataTable names when it sorts.
 *
 * Sent for fidelity only: `GetDashboardListWithPagination` branches on
 * `SortColumnDirection` alone and orders by `BuildingName` either way, so the project name
 * is the only thing this table can actually be sorted by.
 */
const SORT_COLUMN: string = 'buildingName';

/** What `ProjectList` asks for: a page, in this app's own terms rather than the wire's. */
export interface IProjectListQuery {
  /** 1-based, as the pager shows it. */
  page: number;
  pageSize: number;
  search: string;
  /**
   * `''` is not "unsorted": it selects the repository's `OrderByDescending(x => x.IsFavorite)`
   * branch, which is the table's default order.
   */
  sortDirection: 'asc' | 'desc' | '';
  showFavoritesOnly: boolean;
  /**
   * `ILoginResponse.domainGroup`. The one field of the request the API does not fill in
   * for itself - the reference reads it from the MVC session - and it drives the colour
   * rules for every cell in the row.
   */
  domainGroup?: number;
}

/**
 * `BMDesk.Business.Model.PaginationModel`, in its own PascalCase.
 *
 * ASP.NET Core binds a JSON body case-insensitively, so camelCase would also work; the
 * server's own casing is kept so this reads as the C# type it becomes.
 *
 * `UserId` and `UserRoleId` are deliberately absent - see this file's header.
 */
export interface IPaginationRequest {
  /** DataTables' echo counter. The API ignores it; only the MVC proxy echoed it back. */
  Draw: string;
  /** Row offset, as a string: the server does `Convert.ToInt32(Start)` then `Skip`. */
  Start: string;
  /** Page size, as a string: `Convert.ToInt32(Length)` then `Take`. */
  Length: string;
  SortColumn: string;
  SortColumnDirection: string;
  SearchValue: string;
  DomainGroup?: number;
  IsShowFavoriteProject: boolean;
  /** Only the card screen pages by number; the table pages by offset. */
  PageNo: number;
  /** Selects the card screen's ordering branch in the repository. Always false here. */
  IsNewView: boolean;
}

/** One phase cell of one row: `DashboardDetailModel`, reduced to what the table renders. */
export interface IDashboardCell {
  /** What a column is matched on. The reference indexes `lstDashboard` positionally instead. */
  folderId: number;
  /** Base64 folder id, for the folder link the reference builds. */
  encryFolderId: string;
  folderName: string;
  /** Server-computed fill, from the folder's document-completion percentage. */
  colorCode: string;
  /** A complete CSS `border` shorthand, e.g. `2.0px solid#d9d9d9`, not just a colour. */
  borderColorCode: string;
  isAccessRights: boolean;
}

/** One row of the table: `DashboardModel`, reduced to what the table renders. */
export interface IProjectListingRow {
  buildingId: number;
  /** Base64 project id, which every reference link is built from. */
  encryBuildingId: string;
  buildingName: string;
  isFavorite: boolean;
  contactPerson: string;
  /** Id of the system's Maintenance folder, used by the dash rule in {@link resolveCell}. */
  maintenanceFolderId: number;
  /** Whether *this project* has that folder. */
  isMaintenanceFolderAvailable: boolean;
  cells: IDashboardCell[];
}

/**
 * One phase column of the table.
 *
 * Built from `Folder/GetFolderWithoutBIM`, which is what makes the columns dynamic: the
 * reference loops over that same list to render its `<thead>`. `folderId` is numeric
 * because that is what a row's cells are matched on.
 */
export interface IListingColumn {
  folderId: number;
  label: string;
}

/** A page of the table: the rows to show, and how many there are in total. */
export interface IProjectListPage {
  rows: IProjectListingRow[];
  recordsTotal: number;
}

/**
 * What one phase cell renders as.
 *
 * A discriminated union rather than an optional colour, so "no dot here" is a state the
 * table has to handle rather than a colour that happens to be empty.
 */
export type ICellRender =
  | { kind: 'dash' }
  | { kind: 'dot'; colorCode: string; borderColorCode: string };

/** `DashboardDetailModel` as System.Text.Json writes it. */
interface IApiDashboardCell {
  folderId?: number | null;
  /** Already lower-cased in the C# type, so the camelCase policy leaves it alone. */
  encryfolderId?: string | null;
  folderName?: string | null;
  colorCode?: string | null;
  borderColorCode?: string | null;
  isAccessRights?: boolean | null;
}

/** `DashboardModel` as System.Text.Json writes it. */
interface IApiDashboardRow {
  buildingId?: number | null;
  encrybuildingId?: string | null;
  buildingName?: string | null;
  isFavorite?: boolean | null;
  contactPerson?: string | null;
  maintenanceFolderId?: number | null;
  isMaintenanceFolderAvailable?: boolean | null;
  /** Stamped onto every row rather than sent once. See this file's header. */
  recordsTotal?: number | null;
  lstDashboard?: IApiDashboardCell[] | null;
}

/** Trimmed text, or `''` - never the string "null". */
function text(value: string | null | undefined): string {
  return (value || '').trim();
}

/** A finite number, or `fallback`. */
function count(value: number | null | undefined, fallback: number = 0): number {
  return typeof value === 'number' && isFinite(value) ? value : fallback;
}

/**
 * Maps the folder lookup onto the table's phase columns.
 *
 * `ILookupOption.id` is a string because it was mapped for `<select>` use, but a cell is
 * matched on a numeric folder id - so a column whose id is not a number could never match
 * anything and would show a dash on every row forever. Dropping it costs one column;
 * keeping it silently widens the table with a permanently empty one.
 *
 * Typed against the option shape rather than importing `ILookupService`, which is a
 * services-layer concern this layer may not depend on.
 */
export function toListingColumns(options: { id: string; name: string }[] | undefined): IListingColumn[] {
  if (!Array.isArray(options)) {
    return [];
  }

  const columns: IListingColumn[] = [];

  options.forEach((option: { id: string; name: string }): void => {
    const folderId: number = Number(option.id);

    if (isFinite(folderId) && text(option.id).length > 0) {
      columns.push({ folderId, label: text(option.name) });
    }
  });

  return columns;
}

/**
 * Builds the request for one page.
 *
 * `Start` is an offset rather than a page number because that is what the repository
 * skips by; the 1-based page the pager shows exists only on this side.
 *
 * The page size is clamped because `Take(-1)` faults the server query, and -1 is exactly
 * what the reference's own length menu offers as "All".
 */
export function toPaginationRequest(query: IProjectListQuery): IPaginationRequest {
  const pageSize: number = query.pageSize > 0 ? Math.floor(query.pageSize) : DEFAULT_PAGE_SIZE;
  const page: number = query.page > 0 ? Math.floor(query.page) : 1;

  return {
    Draw: '1',
    Start: String((page - 1) * pageSize),
    Length: String(pageSize),
    SortColumn: SORT_COLUMN,
    SortColumnDirection: query.sortDirection,
    SearchValue: text(query.search),
    DomainGroup: query.domainGroup,
    IsShowFavoriteProject: query.showFavoritesOnly,
    PageNo: 0,
    IsNewView: false
  };
}

/**
 * Maps the unwrapped `data` array onto a page.
 *
 * Anything that is not an array of rows becomes an empty page rather than throwing: a
 * failed unwrap hands `undefined` through, and an empty table beats a render that dies.
 */
export function toProjectListPage(payload: unknown): IProjectListPage {
  if (!Array.isArray(payload)) {
    return { rows: [], recordsTotal: 0 };
  }

  const rows: IProjectListingRow[] = payload
    .filter((entry: unknown): boolean => typeof entry === 'object' && entry !== null)
    .map(toRow);

  return {
    rows,
    // Off the first row, exactly as the reference's proxy reads it.
    recordsTotal: count((payload[0] as IApiDashboardRow | undefined)?.recordsTotal)
  };
}

function toRow(entry: unknown): IProjectListingRow {
  const row: IApiDashboardRow = entry as IApiDashboardRow;
  const cells: IApiDashboardCell[] = Array.isArray(row.lstDashboard) ? row.lstDashboard : [];

  return {
    buildingId: count(row.buildingId),
    encryBuildingId: text(row.encrybuildingId),
    buildingName: text(row.buildingName),
    isFavorite: row.isFavorite === true,
    contactPerson: text(row.contactPerson),
    maintenanceFolderId: count(row.maintenanceFolderId),
    isMaintenanceFolderAvailable: row.isMaintenanceFolderAvailable === true,
    cells: cells
      .filter((cell: unknown): boolean => typeof cell === 'object' && cell !== null)
      .map(toCell)
  };
}

function toCell(cell: IApiDashboardCell): IDashboardCell {
  return {
    folderId: count(cell.folderId),
    encryFolderId: text(cell.encryfolderId),
    folderName: text(cell.folderName),
    colorCode: text(cell.colorCode),
    borderColorCode: text(cell.borderColorCode),
    isAccessRights: cell.isAccessRights === true
  };
}

/**
 * What the cell for one column renders as, following the reference view's three branches:
 *
 * ```js
 * if (!row.isMaintenanceFolderAvailable && row.maintenanceFolderId == cell.folderId) "-"
 * else if (cell.isAccessRights == true)                                              dot
 * else                                                                               "-"
 * ```
 *
 * The cell is found **by folder id**, where the reference indexes `lstDashboard[0..5]`
 * positionally against headers it built from a separate call. That holds only while every
 * project's folder set matches the system's exactly and is six long; where it does not,
 * the reference shows one folder's status under another folder's heading. A column with no
 * cell is a dash here rather than the reference's empty string, so a short row still reads
 * as a row.
 */
export function resolveCell(row: IProjectListingRow, folderId: number): ICellRender {
  const cell: IDashboardCell | undefined = row.cells.filter(
    (candidate: IDashboardCell) => candidate.folderId === folderId
  )[0];

  if (!cell) {
    return { kind: 'dash' };
  }

  // A maintenance column on a project that has no maintenance folder, however the colour
  // rules coloured it.
  if (!row.isMaintenanceFolderAvailable && row.maintenanceFolderId === cell.folderId) {
    return { kind: 'dash' };
  }

  if (!cell.isAccessRights) {
    return { kind: 'dash' };
  }

  return { kind: 'dot', colorCode: cell.colorCode, borderColorCode: cell.borderColorCode };
}
