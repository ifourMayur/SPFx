import * as React from 'react';
import styles from './ProjectList.module.scss';
import type { IProjectListProps } from './IProjectListProps';
import { ILookupOption } from '../../../../models/Building';
import {
  ICellRender,
  IListingColumn,
  IProjectListPage,
  IProjectListQuery,
  IProjectListingRow,
  resolveCell,
  toListingColumns
} from '../../../../models/ProjectListing';

const PAGE_SIZE_OPTIONS: number[] = [10, 25, 50, 100];

/** How long the toast stays on screen, matching the reference's `autoHideDelay`. */
const NOTIFICATION_TIMEOUT_MS: number = 3000;

/**
 * How long the search box waits for the typing to stop.
 *
 * The reference's DataTable debounces its own search for the same reason: every keystroke
 * would otherwise be a round trip through a query that reads the whole Buildings table.
 */
const SEARCH_DEBOUNCE_MS: number = 400;

interface IProjectListState {
  /** Phase columns, from `Folder/GetFolderWithoutBIM`. */
  columns: IListingColumn[];
  /** The page currently on screen. */
  rows: IProjectListingRow[];
  /** How many projects there are in total - not how many are on screen. */
  recordsTotal: number;
  isLoading: boolean;
  error?: string;
  /** What is in the search box right now. */
  searchInput: string;
  /** What was last actually searched for, i.e. `searchInput` after the debounce. */
  search: string;
  pageSize: number;
  /** 1-based. */
  page: number;
  /** `''` is the server's favourites-first default, not "unsorted". */
  sortDirection: 'asc' | 'desc' | '';
  showFavoritesOnly: boolean;
  isListView: boolean;
  openMenuRowId?: number;
  /** Stars the user has clicked this session. See `_onToggleFavorite`. */
  favorites: { [buildingId: number]: boolean };
}

/**
 * Project overview table opened from the top-level "Project" item in the shared `Menu`:
 * the port of `Views/Building/_ProjectListingScreen.cshtml`.
 *
 * **The table is server-paged, as the reference's DataTable is** (`serverSide: true`).
 * Search, sort, page and the favourites filter are all round trips to
 * `Dashboard/GetListAllWithPagination` - none of them touches the rows already on screen.
 * That is why this component fetches its own data instead of being handed a `rows` prop:
 * a page is only meaningful together with the query that asked for it, so whoever owns the
 * query has to own the fetch.
 *
 * Two things follow from that and are easy to get wrong:
 *
 * - The pager is driven by `recordsTotal`, which the API stamps onto every row rather than
 *   sending beside them. `rows.length` is the size of *this page*.
 * - Only the project name can be sorted. `GetDashboardListWithPagination` branches on
 *   `SortColumnDirection` and orders by `BuildingName` either way, so the phase columns
 *   are headings, not sort buttons - offering to sort by them would do nothing.
 *
 * The phase columns are dynamic, built from `Folder/GetFolderWithoutBIM` exactly as
 * `BuildingController.ProjectLoadData` builds the reference's `<thead>`, and each cell is
 * matched to its column by folder id.
 *
 * A save reports itself as a single self-dismissing toast, nothing more: what the Web API
 * answered - the project payload and the SharePoint ids bound onto its folder tree - is a
 * development detail, and `BuildingService` already logs it.
 */
export default class ProjectList extends React.Component<IProjectListProps, IProjectListState> {
  /** Pending auto-dismiss of the toast. */
  private _notificationTimer: number | undefined;

  /** Pending debounce of the search box. */
  private _searchTimer: number | undefined;

  /**
   * Sequence number of the most recent page request.
   *
   * Typing quickly, or paging twice, leaves more than one request in flight; without this
   * the slower one can land last and put the wrong page on screen.
   */
  private _requestId: number = 0;

  /** Guards `setState` after the router has unmounted this page mid-request. */
  private _isMounted: boolean = false;

  public constructor(props: IProjectListProps) {
    super(props);

    this.state = {
      columns: [],
      rows: [],
      recordsTotal: 0,
      isLoading: true,
      searchInput: '',
      search: '',
      pageSize: PAGE_SIZE_OPTIONS[0],
      page: 1,
      sortDirection: '',
      showFavoritesOnly: false,
      isListView: true,
      favorites: {}
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    document.addEventListener('mousedown', this._onDocumentMouseDown);
    this._scheduleNotificationDismiss();

    this._loadColumns();
    this._loadPage();
  }

  public componentDidUpdate(previousProps: IProjectListProps, previousState: IProjectListState): void {
    if (previousProps.notification !== this.props.notification) {
      this._scheduleNotificationDismiss();
    }

    // Every part of the query is the server's business, so any change to one is a refetch.
    if (
      previousState.search !== this.state.search ||
      previousState.page !== this.state.page ||
      previousState.pageSize !== this.state.pageSize ||
      previousState.sortDirection !== this.state.sortDirection ||
      previousState.showFavoritesOnly !== this.state.showFavoritesOnly
    ) {
      this._loadPage();
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    document.removeEventListener('mousedown', this._onDocumentMouseDown);
    this._clearNotificationTimer();
    this._clearSearchTimer();
  }

  public render(): React.ReactElement<IProjectListProps> {
    const { canAddEdit, notification, siteUrl } = this.props;
    const { searchInput, pageSize, showFavoritesOnly, isListView, rows, recordsTotal, page } = this.state;
    const totalPages: number = Math.max(1, Math.ceil(recordsTotal / pageSize));
    const startEntry: number = recordsTotal === 0 ? 0 : (page - 1) * pageSize + 1;
    const endEntry: number = (page - 1) * pageSize + rows.length;

    return (
      <section className={styles.projectList}>
        {notification && (
          <div className={styles.toastHost}>
            <div className={styles.toast} role="status" aria-live="polite">
              <span className={styles.toastIcon} aria-hidden="true">&#10003;</span>
              <span className={styles.toastMessage}>{notification}</span>
              <button
                type="button"
                className={styles.toastClose}
                aria-label="Dismiss"
                onClick={this._onDismissNotification}
              >
                &times;
              </button>
            </div>
          </div>
        )}

        {/* Which site collection this session is working in. The shell's site bar names it
            by title; the URL is what identifies it, and this is the screen a project - and
            therefore its SharePoint folder tree - is created from. */}
        {siteUrl && (
          <div className={styles.siteUrlBar}>
            <span className={styles.siteUrlLabel}>Site URL:</span>{' '}
            <a
              className={styles.siteUrlValue}
              href={siteUrl}
              target="_blank"
              rel="noreferrer"
            >
              {siteUrl}
            </a>
          </div>
        )}

        {/* Gated exactly as the reference gates it: `@if (ViewBag.AddEditAccessRights)`. */}
        {canAddEdit && (
          <div className={styles.commandBar}>
            <button type="button" className={styles.addButton} onClick={this.props.onAddProject}>
              + Add Project
            </button>
          </div>
        )}

        <div className={styles.toolbar}>
          <label className={styles.toggleField}>
            <span className={styles.toggle}>
              <input type="checkbox" checked={showFavoritesOnly} onChange={this._onToggleFavoritesOnly} />
              <span className={styles.toggleTrack} aria-hidden="true" />
            </span>
            <span>Alleen favorieten weergeven</span>
          </label>

          <label className={`${styles.toggleField} ${styles.toggleFieldEnd}`}>
            <span>Overschakelen naar lijst</span>
            <span className={styles.toggle}>
              <input type="checkbox" checked={isListView} onChange={this._onToggleListView} />
              <span className={styles.toggleTrack} aria-hidden="true" />
            </span>
          </label>
        </div>

        <div className={styles.controls}>
          <label className={styles.pageSizeField}>
            Show{' '}
            <select value={pageSize} onChange={this._onPageSizeChange}>
              {PAGE_SIZE_OPTIONS.map((size: number) => <option key={size} value={size}>{size}</option>)}
            </select>{' '}
            entries
          </label>

          <label className={styles.searchField}>
            Search:{' '}
            <input type="search" value={searchInput} onChange={this._onSearchChange} />
          </label>
        </div>

        {isListView ? this._renderTable() : this._renderCompactList()}

        <div className={styles.footer}>
          <span>
            {recordsTotal === 0
              ? 'No entries'
              : `Showing ${startEntry} to ${endEntry} of ${recordsTotal} entries`}
          </span>
          <div className={styles.pagination}>
            <button type="button" disabled={page <= 1} onClick={this._onPreviousPage}>Previous</button>
            <span className={styles.pageIndicator}>{page}</span>
            <button type="button" disabled={page >= totalPages} onClick={this._onNextPage}>Next</button>
          </div>
        </div>
      </section>
    );
  }

  private _renderTable(): React.ReactElement {
    const { columns, rows } = this.state;

    return (
      <div className={styles.tableWrapper}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.favoriteHeader} aria-hidden="true" />
              {this._renderNameHeader()}
              {/* Headings, not sort buttons: the API can only order by project name. */}
              {columns.map((column: IListingColumn) => (
                <th key={column.folderId} className={styles.statusCell}>{column.label}</th>
              ))}
              <th className={styles.optionHeader}>Optie</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row: IProjectListingRow) => this._renderRow(row))}
            {rows.length === 0 && (
              <tr>
                <td className={styles.empty} colSpan={columns.length + 3}>
                  {this._renderEmptyMessage()}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    );
  }

  /** What stands in for the rows: a failure, a load in progress, or genuinely nothing. */
  private _renderEmptyMessage(): string {
    const { isLoading, error } = this.state;

    if (error) {
      return error;
    }

    return isLoading ? 'Loading projects…' : 'No projects found.';
  }

  private _renderCompactList(): React.ReactElement {
    const { rows } = this.state;

    return (
      <ul className={styles.compactList}>
        {rows.map((row: IProjectListingRow) => (
          <li key={row.buildingId} className={styles.compactItem}>
            {this._renderStar(row)}
            <span className={styles.compactName}>{row.buildingName}</span>
          </li>
        ))}
        {rows.length === 0 && <li className={styles.empty}>{this._renderEmptyMessage()}</li>}
      </ul>
    );
  }

  /** The only sortable column: `SortColumnDirection` orders by `BuildingName` or nothing. */
  private _renderNameHeader(): React.ReactElement {
    const { sortDirection } = this.state;

    return (
      <th>
        <button type="button" className={styles.sortButton} onClick={this._onSortByName}>
          Naam van het project
          <span className={styles.sortIcon} aria-hidden="true">
            {sortDirection === 'asc' ? '▲' : sortDirection === 'desc' ? '▼' : '⇅'}
          </span>
        </button>
      </th>
    );
  }

  private _renderRow(row: IProjectListingRow): React.ReactElement {
    const { openMenuRowId, columns } = this.state;

    return (
      <tr key={row.buildingId}>
        <td>{this._renderStar(row)}</td>
        <td className={styles.nameCell}>
          <a href="#" className={styles.nameLink} onClick={this._onNameClick}>{row.buildingName}</a>
        </td>
        {columns.map((column: IListingColumn) => (
          <td key={column.folderId} className={styles.statusCell}>
            {ProjectList._renderCell(row, column)}
          </td>
        ))}
        <td className={styles.optionCell}>
          <div className={styles.menuWrapper}>
            <button
              type="button"
              className={styles.optionButton}
              data-id={row.buildingId}
              aria-haspopup="menu"
              aria-expanded={openMenuRowId === row.buildingId}
              onClick={this._onToggleRowMenu}
            >
              &hellip;
            </button>
            {openMenuRowId === row.buildingId && (
              <ul className={styles.rowMenu} role="menu">
                <li role="none"><button type="button" role="menuitem" onClick={this._onCloseRowMenu}>Openen</button></li>
                {/* Same gate as the Add button: the reference wraps its Edit link in
                    `@if (ViewBag.AddEditAccessRights)` too. */}
                {this.props.canAddEdit && (
                  <li role="none">
                    <button type="button" role="menuitem" data-id={row.buildingId} onClick={this._onEditRow}>
                      Bewerken
                    </button>
                  </li>
                )}
                <li role="none"><button type="button" role="menuitem" onClick={this._onCloseRowMenu}>Verwijderen</button></li>
              </ul>
            )}
          </div>
        </td>
      </tr>
    );
  }

  /**
   * One phase cell, coloured by the server.
   *
   * `colorCode` is a document-completion percentage turned into a colour by
   * `DashboardRepository.GetColorCodes`, and `borderColorCode` is a whole CSS `border`
   * shorthand rather than a colour - both are applied inline exactly as the reference's
   * `style='background-color: …;border: …'` does. A local traffic-light enum would have to
   * throw every intermediate percentage away.
   *
   * The dot carries the column's name so it is not colour-only to a screen reader.
   */
  private static _renderCell(row: IProjectListingRow, column: IListingColumn): React.ReactElement {
    const cell: ICellRender = resolveCell(row, column.folderId);

    if (cell.kind === 'dash') {
      return <span className={styles.dash}>-</span>;
    }

    return (
      <span
        className={styles.circle}
        role="img"
        aria-label={column.label}
        style={{ backgroundColor: cell.colorCode, border: cell.borderColorCode }}
      />
    );
  }

  /**
   * The favourite star.
   *
   * Clicking it moves the star and nothing else: persisting a favourite is
   * `Building/IsFavoriteBuilding`, which this screen does not call yet. The flag therefore
   * comes back as the server still has it on the next page load, and "Alleen favorieten
   * weergeven" filters on the server's value rather than on what is on screen.
   */
  private _renderStar(row: IProjectListingRow): React.ReactElement {
    const isFavorite: boolean = this._isFavorite(row);

    return (
      <button
        type="button"
        className={styles.star}
        aria-pressed={isFavorite}
        aria-label={isFavorite ? 'Remove from favorites' : 'Add to favorites'}
        data-id={row.buildingId}
        onClick={this._onToggleFavorite}
      >
        {isFavorite ? '★' : '☆'}
      </button>
    );
  }

  private _isFavorite(row: IProjectListingRow): boolean {
    const override: boolean | undefined = this.state.favorites[row.buildingId];
    return override === undefined ? row.isFavorite : override;
  }

  /**
   * Loads the phase columns.
   *
   * A failure here is deliberately not surfaced as an error: the projects themselves still
   * load and are still worth showing, just without their status columns. Turning a missing
   * `Folder/GetFolderWithoutBIM` into a blank screen would be a worse trade.
   */
  private _loadColumns(): void {
    this.props.lookupService
      .getFolderColumns()
      .then((options: ILookupOption[]): void => {
        if (this._isMounted) {
          this.setState({ columns: toListingColumns(options) });
        }
      })
      .catch((): void => {
        if (this._isMounted) {
          this.setState({ columns: [] });
        }
      });
  }

  /** Loads the page the current state describes, ignoring anything that lands out of order. */
  private _loadPage(): void {
    const requestId: number = ++this._requestId;

    this.setState({ isLoading: true, error: undefined });

    this.props.listingService
      .getProjectPage(this._toQuery())
      .then((result: IProjectListPage): void => {
        // A slower earlier request must not overwrite a faster later one.
        if (!this._isMounted || requestId !== this._requestId) {
          return;
        }

        this.setState({ rows: result.rows, recordsTotal: result.recordsTotal, isLoading: false });
        this.props.onProjectsLoaded(result.rows);
      })
      .catch((error: unknown): void => {
        if (!this._isMounted || requestId !== this._requestId) {
          return;
        }

        // Named rather than swallowed: an empty table would be indistinguishable from a
        // tenant that genuinely has no projects.
        this.setState({
          rows: [],
          recordsTotal: 0,
          isLoading: false,
          error: error instanceof Error ? error.message : 'The project list could not be loaded.'
        });
      });
  }

  private _toQuery(): IProjectListQuery {
    const { page, pageSize, search, sortDirection, showFavoritesOnly } = this.state;

    return {
      page,
      pageSize,
      search,
      sortDirection,
      showFavoritesOnly,
      domainGroup: this.props.domainGroup
    };
  }

  /** Restarts the auto-dismiss countdown for the current toast. */
  private _scheduleNotificationDismiss(): void {
    this._clearNotificationTimer();

    if (!this.props.notification) {
      return;
    }

    this._notificationTimer = setTimeout((): void => {
      this._notificationTimer = undefined;
      this.props.onDismissNotification();
    }, NOTIFICATION_TIMEOUT_MS) as unknown as number;
  }

  private _clearNotificationTimer(): void {
    if (this._notificationTimer !== undefined) {
      clearTimeout(this._notificationTimer);
      this._notificationTimer = undefined;
    }
  }

  private _clearSearchTimer(): void {
    if (this._searchTimer !== undefined) {
      clearTimeout(this._searchTimer);
      this._searchTimer = undefined;
    }
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login`/`Menu` use for their own click handlers.

  /**
   * Types into the search box without searching yet.
   *
   * The box stays responsive because `searchInput` updates immediately; only `search` -
   * the value the query is built from - waits for the typing to stop.
   */
  private _onSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const value: string = event.target.value;

    this.setState({ searchInput: value });
    this._clearSearchTimer();

    this._searchTimer = setTimeout((): void => {
      this._searchTimer = undefined;
      // Back to page one: the old page number belongs to the old result set.
      this.setState({ search: value, page: 1 });
    }, SEARCH_DEBOUNCE_MS) as unknown as number;
  };

  private _onPageSizeChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ pageSize: Number(event.target.value), page: 1 });
  };

  private _onToggleFavoritesOnly = (): void => {
    this.setState((state: IProjectListState) => ({
      showFavoritesOnly: !state.showFavoritesOnly,
      page: 1
    }));
  };

  private _onToggleListView = (): void => {
    this.setState((state: IProjectListState) => ({ isListView: !state.isListView }));
  };

  /** Cycles the one sort the API offers: favourites-first, then name ascending, then descending. */
  private _onSortByName = (): void => {
    this.setState((state: IProjectListState) => ({
      sortDirection: state.sortDirection === 'asc' ? 'desc' : 'asc',
      page: 1
    }));
  };

  private _onToggleFavorite = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const id: number = Number(event.currentTarget.dataset.id);

    this.setState((state: IProjectListState) => {
      const current: boolean | undefined = state.favorites[id];
      const row: IProjectListingRow | undefined = state.rows.filter(
        (candidate: IProjectListingRow) => candidate.buildingId === id
      )[0];

      return {
        favorites: {
          ...state.favorites,
          [id]: !(current === undefined ? row?.isFavorite : current)
        }
      };
    });
  };

  private _onEditRow = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const id: number = Number(event.currentTarget.dataset.id);

    this.setState({ openMenuRowId: undefined });
    this.props.onEditProject(id);
  };

  private _onDismissNotification = (): void => {
    this._clearNotificationTimer();
    this.props.onDismissNotification();
  };

  private _onPreviousPage = (): void => {
    this.setState((state: IProjectListState) => ({ page: Math.max(1, state.page - 1) }));
  };

  private _onNextPage = (): void => {
    this.setState((state: IProjectListState) => ({ page: state.page + 1 }));
  };

  private _onToggleRowMenu = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const id: number = Number(event.currentTarget.dataset.id);
    this.setState((state: IProjectListState) => ({
      openMenuRowId: state.openMenuRowId === id ? undefined : id
    }));
  };

  private _onCloseRowMenu = (): void => {
    this.setState({ openMenuRowId: undefined });
  };

  private _onNameClick = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  };

  private _onDocumentMouseDown = (event: MouseEvent): void => {
    if (this.state.openMenuRowId === undefined) {
      return;
    }

    const target: HTMLElement = event.target as HTMLElement;
    if (!target.closest(`.${styles.menuWrapper}`)) {
      this.setState({ openMenuRowId: undefined });
    }
  };
}
