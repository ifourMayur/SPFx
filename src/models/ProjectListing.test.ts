/**
 * The pure rules behind the project overview table: what goes on the wire, what comes
 * back off it, and what each phase cell renders as.
 *
 * These matter more than they look. The request is a `PaginationModel` whose fields the
 * API partly overwrites and partly trusts, and the total row count is not sent alongside
 * the page - it is stamped onto every row. Getting either wrong produces a table that
 * renders without complaint and is quietly wrong: the first page repeated forever, or a
 * pager that thinks there is one page.
 */

import {
  IDashboardCell,
  ICellRender,
  IListingColumn,
  IPaginationRequest,
  IProjectListPage,
  IProjectListingRow,
  resolveCell,
  toListingColumns,
  toPaginationRequest,
  toProjectListPage
} from './ProjectListing';

/** A query for the second page, as `ProjectList` would build it. */
const QUERY = {
  page: 2,
  pageSize: 25,
  search: '',
  sortDirection: 'asc' as const,
  showFavoritesOnly: false,
  domainGroup: 1
};

/**
 * One row of `Dashboard/GetListAllWithPagination`, written out as the Web API actually
 * answers it - `ResponseDetail` camelCased by System.Text.Json, wrapping
 * `List<DashboardModel>`. Written literally so a change in the real shape fails here.
 */
const API_ROW = {
  buildingId: 41,
  encrybuildingId: 'NDE=',
  buildingName: 'Office Renovation Plan',
  isFavorite: true,
  contactPerson: 'Jane Dijkstra',
  openActions: 0,
  maintenanceFolderId: 6,
  isMaintenanceFolderAvailable: false,
  recordsTotal: 137,
  lstDashboard: [
    {
      folderId: 1,
      encryfolderId: 'MQ==',
      folderName: 'Administratie',
      colorCode: '#22b14c',
      borderColorCode: '2.0px solid#d9d9d9',
      isAccessRights: true,
      isSubFolderExist: true,
      index: 1
    },
    {
      folderId: 2,
      encryfolderId: 'Mg==',
      folderName: 'Voorbereiding',
      colorCode: '#000000',
      borderColorCode: '2.0px solid#d9d9d9',
      isAccessRights: false,
      isSubFolderExist: false,
      index: 2
    },
    {
      folderId: 6,
      encryfolderId: 'Ng==',
      folderName: 'Maintenance',
      colorCode: '#22b14c',
      borderColorCode: '2.0px solid#d9d9d9',
      isAccessRights: true,
      isSubFolderExist: true,
      index: 3
    }
  ]
};

describe('toPaginationRequest - what the listing puts on the wire', () => {
  /**
   * `DashboardRepository.GetDashboardListWithPagination` does `Skip(Start).Take(Length)`
   * on a 1-based page number it never sees, so the offset has to be computed here.
   */
  it('turns a page number into the row offset the server skips', () => {
    const request: IPaginationRequest = toPaginationRequest(QUERY);

    expect(request.Start).toBe('25');
    expect(request.Length).toBe('25');
  });

  it('asks for the first page from offset zero', () => {
    expect(toPaginationRequest({ ...QUERY, page: 1 }).Start).toBe('0');
  });

  /**
   * `DashboardController.GetListAllWithPagination` assigns `model.UserId` and
   * `model.UserRoleId` from the caller's JWT before the repository sees them. Sending
   * them would be at best ignored and at worst misleading to read.
   */
  it('omits the identity fields the API overwrites from the JWT', () => {
    const request: IPaginationRequest = toPaginationRequest(QUERY);

    expect(request).not.toHaveProperty('UserId');
    expect(request).not.toHaveProperty('UserRoleId');
  });

  /**
   * `DomainGroup` is the one field the API does *not* fill in for itself - the reference
   * reads it from the MVC session, and this app has it on the sign-in response. It drives
   * the colour rules for the whole row.
   */
  it('sends the domain group, which the API does not derive for itself', () => {
    expect(toPaginationRequest(QUERY).DomainGroup).toBe(1);
  });

  /**
   * `Take(-1)` throws server-side, so the "All" case the reference's length menu offers
   * can never reach the API as -1.
   */
  it('clamps a non-positive page size, which would fault the server query', () => {
    expect(Number(toPaginationRequest({ ...QUERY, pageSize: -1 }).Length)).toBeGreaterThan(0);
  });

  /**
   * An empty `SortColumnDirection` is meaningful, not missing: it selects the repository's
   * `OrderByDescending(x => x.IsFavorite)` branch, which is the table's default order.
   */
  it('passes an empty sort direction through as the favourites-first default', () => {
    expect(toPaginationRequest({ ...QUERY, sortDirection: '' }).SortColumnDirection).toBe('');
  });

  it('trims the search term, which the server matches with Contains', () => {
    expect(toPaginationRequest({ ...QUERY, search: '  renovation  ' }).SearchValue).toBe('renovation');
  });

  /**
   * `IsNewView` selects a different ordering branch in the repository and is only set by
   * the card screen's own endpoint. The listing table must leave it false.
   */
  it('stays on the table branch rather than the card screen branch', () => {
    expect(toPaginationRequest(QUERY).IsNewView).toBe(false);
  });

  it('carries the favourites-only filter the toggle drives', () => {
    expect(toPaginationRequest({ ...QUERY, showFavoritesOnly: true }).IsShowFavoriteProject).toBe(true);
  });
});

describe('toProjectListPage - reading the page back', () => {
  /**
   * The total is not sent beside the rows: `GetDashboardListWithPagination` stamps it onto
   * every row, and the reference's MVC proxy reads `data.FirstOrDefault().RecordsTotal`.
   * Miss this and the pager believes there is exactly one page.
   */
  it('reads the record count off the first row, where the API stamps it', () => {
    const page: IProjectListPage = toProjectListPage([API_ROW]);

    expect(page.recordsTotal).toBe(137);
  });

  it('reports no records when the page came back empty', () => {
    expect(toProjectListPage([])).toEqual({ rows: [], recordsTotal: 0 });
  });

  /** A failed unwrap hands `undefined` through; an empty table beats a thrown render. */
  it('treats a payload that is not a row array as an empty page', () => {
    expect(toProjectListPage(undefined)).toEqual({ rows: [], recordsTotal: 0 });
  });

  it('maps the row fields the table renders', () => {
    const row: IProjectListingRow = toProjectListPage([API_ROW]).rows[0];

    expect(row.buildingId).toBe(41);
    expect(row.buildingName).toBe('Office Renovation Plan');
    expect(row.isFavorite).toBe(true);
    expect(row.contactPerson).toBe('Jane Dijkstra');
  });

  it('maps each phase cell, keeping the folder id the columns are matched on', () => {
    const cells: IDashboardCell[] = toProjectListPage([API_ROW]).rows[0].cells;

    expect(cells.map((cell: IDashboardCell) => cell.folderId)).toEqual([1, 2, 6]);
    expect(cells[0].colorCode).toBe('#22b14c');
    expect(cells[0].borderColorCode).toBe('2.0px solid#d9d9d9');
  });

  /** `lstDashboard` is absent on a row whose project has no folders at all. */
  it('survives a row with no dashboard cells', () => {
    const page: IProjectListPage = toProjectListPage([{ ...API_ROW, lstDashboard: undefined }]);

    expect(page.rows[0].cells).toEqual([]);
  });
});

describe('resolveCell - what a phase cell renders as', () => {
  /** The mapped row, built per test so a failure here cannot abort the whole suite. */
  function row(overrides?: Partial<typeof API_ROW>): IProjectListingRow {
    return toProjectListPage([{ ...API_ROW, ...overrides }]).rows[0];
  }

  it('renders the colour the server computed for a folder the user may see', () => {
    const cell: ICellRender = resolveCell(row(), 1);

    expect(cell).toEqual({
      kind: 'dot',
      colorCode: '#22b14c',
      borderColorCode: '2.0px solid#d9d9d9'
    });
  });

  /** `else if (row.lstDashboard[i].isAccessRights == true)` in the reference view. */
  it('renders a dash for a folder this user has no access rights to', () => {
    expect(resolveCell(row(), 2)).toEqual({ kind: 'dash' });
  });

  /**
   * The reference's first branch: a maintenance column on a project that has no
   * maintenance folder is a dash, however the colour rules coloured it.
   */
  it('renders a dash in the maintenance column of a project that has none', () => {
    expect(resolveCell(row(), 6)).toEqual({ kind: 'dash' });
  });

  it('renders the maintenance column normally once the project has that folder', () => {
    expect(resolveCell(row({ isMaintenanceFolderAvailable: true }), 6)).toEqual({
      kind: 'dot',
      colorCode: '#22b14c',
      borderColorCode: '2.0px solid#d9d9d9'
    });
  });

  /**
   * The column set comes from `Folder/GetFolderWithoutBIM`, which is every active folder
   * in the system; a project whose template omits one has no cell for that column.
   */
  it('renders a dash for a column this project has no cell for', () => {
    expect(resolveCell(row(), 99)).toEqual({ kind: 'dash' });
  });
});

describe('toListingColumns - the phase columns the table shows', () => {
  /**
   * `Folder/GetFolderWithoutBIM` answers `List<LookupItem>`, which `toLookupOptions` has
   * already stringified for `<select>` use. The cells are matched on a numeric folder id,
   * so it has to come back.
   */
  it('turns the folder lookup into columns carrying a numeric folder id', () => {
    const columns: IListingColumn[] = toListingColumns([
      { id: '1', name: 'Administratie' },
      { id: '6', name: 'Maintenance' }
    ]);

    expect(columns).toEqual([
      { folderId: 1, label: 'Administratie' },
      { folderId: 6, label: 'Maintenance' }
    ]);
  });

  /**
   * A column whose id is not a number could never match a cell, so every row would show a
   * dash under it. Dropping it loses one column; keeping it silently widens the table with
   * a permanently empty one.
   */
  it('drops a folder whose id is not a number', () => {
    expect(toListingColumns([{ id: 'BIM', name: 'BIM' }])).toEqual([]);
  });

  it('treats a failed column load as no columns rather than throwing', () => {
    expect(toListingColumns(undefined)).toEqual([]);
  });
});
