import * as React from 'react';
import styles from './ProjectList.module.scss';

/** Traffic-light status shown as a circle for a single project phase. */
type PhaseStatus = 'none' | 'success' | 'danger';

/** The five phase columns every project always shows a circle for. */
type PhaseKey = 'administratie' | 'voorbereiding' | 'ontwerp' | 'uitvoering' | 'revisie';

/** Columns the table can be sorted by. */
type SortKey = 'name' | PhaseKey | 'maintenance';

interface IProjectListRow {
  id: number;
  name: string;
  isFavorite: boolean;
  administratie: PhaseStatus;
  voorbereiding: PhaseStatus;
  ontwerp: PhaseStatus;
  uitvoering: PhaseStatus;
  revisie: PhaseStatus;
  /** Undefined renders as "-": most projects have not reached their maintenance phase yet. */
  maintenance?: PhaseStatus;
}

const PHASE_COLUMNS: Array<{ key: PhaseKey; label: string }> = [
  { key: 'administratie', label: 'Administratie' },
  { key: 'voorbereiding', label: 'Voorbereiding' },
  { key: 'ontwerp', label: 'Ontwerp' },
  { key: 'uitvoering', label: 'Uitvoering' },
  { key: 'revisie', label: 'Revisie' }
];

const PAGE_SIZE_OPTIONS: number[] = [10, 25, 50, 100];

const SORT_RANK: Record<PhaseStatus, number> = { none: 0, success: 1, danger: 2 };

/**
 * Sample rows for the project overview table, matching the reference design.
 *
 * `IProjectService` does not expose per-phase status yet, so this component renders static
 * sample data rather than a real project list - replace with a service call (following the
 * `Project`/`ProjectService` pattern) once the Web API exposes that data.
 */
const SAMPLE_ROWS: IProjectListRow[] = [
  { id: 1, name: 'Website Redesign', isFavorite: false, administratie: 'success', voorbereiding: 'success', ontwerp: 'success', uitvoering: 'none', revisie: 'none' },
  { id: 2, name: 'Office Renovation Plan', isFavorite: false, administratie: 'success', voorbereiding: 'danger', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 3, name: 'Tender Response A12', isFavorite: true, administratie: 'none', voorbereiding: 'none', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 4, name: 'Client Onboarding Portal', isFavorite: false, administratie: 'success', voorbereiding: 'success', ontwerp: 'success', uitvoering: 'success', revisie: 'none' },
  { id: 5, name: 'Infrastructure Upgrade', isFavorite: false, administratie: 'success', voorbereiding: 'success', ontwerp: 'danger', uitvoering: 'none', revisie: 'none', maintenance: 'danger' },
  { id: 6, name: 'Marketing Campaign Q3', isFavorite: false, administratie: 'none', voorbereiding: 'none', ontwerp: 'success', uitvoering: 'none', revisie: 'none' },
  { id: 7, name: 'Data Migration Project', isFavorite: false, administratie: 'success', voorbereiding: 'success', ontwerp: 'success', uitvoering: 'success', revisie: 'success', maintenance: 'success' },
  { id: 8, name: 'Vendor Contract Review', isFavorite: true, administratie: 'danger', voorbereiding: 'none', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 9, name: 'Product Launch Roadmap', isFavorite: false, administratie: 'success', voorbereiding: 'none', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 10, name: 'Annual Compliance Audit', isFavorite: false, administratie: 'success', voorbereiding: 'success', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 11, name: 'New project', isFavorite: false, administratie: 'none', voorbereiding: 'none', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 12, name: 'new projects', isFavorite: false, administratie: 'none', voorbereiding: 'none', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 13, name: 'NewProject create folder', isFavorite: false, administratie: 'none', voorbereiding: 'none', ontwerp: 'none', uitvoering: 'none', revisie: 'none' },
  { id: 14, name: 'Project by dev test', isFavorite: false, administratie: 'success', voorbereiding: 'none', ontwerp: 'danger', uitvoering: 'success', revisie: 'success', maintenance: 'danger' },
  { id: 15, name: 'SharePointSite Project', isFavorite: false, administratie: 'danger', voorbereiding: 'danger', ontwerp: 'danger', uitvoering: 'danger', revisie: 'danger' },
  { id: 16, name: 'Tender project', isFavorite: false, administratie: 'none', voorbereiding: 'none', ontwerp: 'danger', uitvoering: 'none', revisie: 'none' },
  { id: 17, name: 'Test document', isFavorite: false, administratie: 'danger', voorbereiding: 'none', ontwerp: 'success', uitvoering: 'danger', revisie: 'danger' }
];

const CIRCLE_CLASS_NAME: Record<PhaseStatus, string> = {
  none: styles.circleNone,
  success: styles.circleSuccess,
  danger: styles.circleDanger
};

interface IProjectListState {
  rows: IProjectListRow[];
  search: string;
  pageSize: number;
  currentPage: number;
  sortKey?: SortKey;
  sortDirection: 'asc' | 'desc';
  showFavoritesOnly: boolean;
  isListView: boolean;
  openMenuRowId?: number;
}

/**
 * Project overview table opened from the top-level "Project" item in the shared `Menu`,
 * matching the reference DataTables-style design: favorites/list-view toggles, a
 * show-N-entries and search control, a sortable phase-status table, and pagination.
 */
export default class ProjectList extends React.Component<Record<string, never>, IProjectListState> {
  public constructor(props: Record<string, never>) {
    super(props);

    this.state = {
      rows: SAMPLE_ROWS,
      search: '',
      pageSize: PAGE_SIZE_OPTIONS[0],
      currentPage: 1,
      sortDirection: 'asc',
      showFavoritesOnly: false,
      isListView: true
    };
  }

  public componentDidMount(): void {
    document.addEventListener('mousedown', this._onDocumentMouseDown);
  }

  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this._onDocumentMouseDown);
  }

  public render(): React.ReactElement {
    const { search, pageSize, showFavoritesOnly, isListView } = this.state;
    const filteredRows: IProjectListRow[] = this._getFilteredRows();
    const sortedRows: IProjectListRow[] = this._getSortedRows(filteredRows);
    const totalEntries: number = sortedRows.length;
    const totalPages: number = Math.max(1, Math.ceil(totalEntries / pageSize));
    const currentPage: number = Math.min(this.state.currentPage, totalPages);
    const startIndex: number = totalEntries === 0 ? 0 : (currentPage - 1) * pageSize;
    const pagedRows: IProjectListRow[] = sortedRows.slice(startIndex, startIndex + pageSize);
    const startEntry: number = totalEntries === 0 ? 0 : startIndex + 1;
    const endEntry: number = Math.min(startIndex + pageSize, totalEntries);

    return (
      <section className={styles.projectList}>
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
            <input type="search" value={search} onChange={this._onSearchChange} />
          </label>
        </div>

        {isListView ? this._renderTable(pagedRows) : this._renderCompactList(pagedRows)}

        <div className={styles.footer}>
          <span>
            {totalEntries === 0 ? 'No entries' : `Showing ${startEntry} to ${endEntry} of ${totalEntries} entries`}
          </span>
          <div className={styles.pagination}>
            <button type="button" disabled={currentPage <= 1} onClick={this._onPreviousPage}>Previous</button>
            <span className={styles.pageIndicator}>{currentPage}</span>
            <button type="button" disabled={currentPage >= totalPages} onClick={this._onNextPage}>Next</button>
          </div>
        </div>
      </section>
    );
  }

  private _renderTable(pagedRows: IProjectListRow[]): React.ReactElement {
    return (
      <div className={styles.tableWrapper}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.favoriteHeader} aria-hidden="true" />
              {this._renderSortableHeader('name', 'Naam van het project')}
              {PHASE_COLUMNS.map((column) => this._renderSortableHeader(column.key, column.label))}
              {this._renderSortableHeader('maintenance', 'Maintenance')}
              <th className={styles.optionHeader}>Optie</th>
            </tr>
          </thead>
          <tbody>
            {pagedRows.map((row: IProjectListRow) => this._renderRow(row))}
            {pagedRows.length === 0 && (
              <tr>
                <td className={styles.empty} colSpan={PHASE_COLUMNS.length + 4}>No projects found.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    );
  }

  private _renderCompactList(pagedRows: IProjectListRow[]): React.ReactElement {
    return (
      <ul className={styles.compactList}>
        {pagedRows.map((row: IProjectListRow) => (
          <li key={row.id} className={styles.compactItem}>
            <button
              type="button"
              className={styles.star}
              aria-pressed={row.isFavorite}
              aria-label={row.isFavorite ? 'Remove from favorites' : 'Add to favorites'}
              data-id={row.id}
              onClick={this._onToggleFavorite}
            >
              {row.isFavorite ? '★' : '☆'}
            </button>
            <span className={styles.compactName}>{row.name}</span>
          </li>
        ))}
        {pagedRows.length === 0 && <li className={styles.empty}>No projects found.</li>}
      </ul>
    );
  }

  private _renderSortableHeader(key: SortKey, label: string): React.ReactElement {
    const { sortKey, sortDirection } = this.state;
    const isActive: boolean = sortKey === key;

    return (
      <th key={key}>
        <button type="button" className={styles.sortButton} data-key={key} onClick={this._onSortClick}>
          {label}
          <span className={styles.sortIcon} aria-hidden="true">
            {isActive ? (sortDirection === 'asc' ? '▲' : '▼') : '⇅'}
          </span>
        </button>
      </th>
    );
  }

  private _renderRow(row: IProjectListRow): React.ReactElement {
    const { openMenuRowId } = this.state;

    return (
      <tr key={row.id}>
        <td>
          <button
            type="button"
            className={styles.star}
            aria-pressed={row.isFavorite}
            aria-label={row.isFavorite ? 'Remove from favorites' : 'Add to favorites'}
            data-id={row.id}
            onClick={this._onToggleFavorite}
          >
            {row.isFavorite ? '★' : '☆'}
          </button>
        </td>
        <td className={styles.nameCell}>
          <a href="#" className={styles.nameLink} onClick={this._onNameClick}>{row.name}</a>
        </td>
        {PHASE_COLUMNS.map((column) => (
          <td key={column.key} className={styles.statusCell}>
            <span
              className={`${styles.circle} ${CIRCLE_CLASS_NAME[row[column.key]]}`}
              aria-label={row[column.key]}
            />
          </td>
        ))}
        <td className={styles.statusCell}>
          {row.maintenance
            ? <span className={`${styles.circle} ${CIRCLE_CLASS_NAME[row.maintenance]}`} aria-label={row.maintenance} />
            : <span className={styles.dash}>-</span>}
        </td>
        <td className={styles.optionCell}>
          <div className={styles.menuWrapper}>
            <button
              type="button"
              className={styles.optionButton}
              data-id={row.id}
              aria-haspopup="menu"
              aria-expanded={openMenuRowId === row.id}
              onClick={this._onToggleRowMenu}
            >
              &hellip;
            </button>
            {openMenuRowId === row.id && (
              <ul className={styles.rowMenu} role="menu">
                <li role="none"><button type="button" role="menuitem" onClick={this._onCloseRowMenu}>Openen</button></li>
                <li role="none"><button type="button" role="menuitem" onClick={this._onCloseRowMenu}>Bewerken</button></li>
                <li role="none"><button type="button" role="menuitem" onClick={this._onCloseRowMenu}>Verwijderen</button></li>
              </ul>
            )}
          </div>
        </td>
      </tr>
    );
  }

  private _getFilteredRows(): IProjectListRow[] {
    const { rows, search, showFavoritesOnly } = this.state;
    const term: string = search.trim().toLowerCase();

    return rows.filter((row: IProjectListRow) => {
      if (showFavoritesOnly && !row.isFavorite) {
        return false;
      }

      if (term && row.name.toLowerCase().indexOf(term) === -1) {
        return false;
      }

      return true;
    });
  }

  private _getSortedRows(rows: IProjectListRow[]): IProjectListRow[] {
    const { sortKey, sortDirection } = this.state;
    if (!sortKey) {
      return rows;
    }

    const direction: number = sortDirection === 'asc' ? 1 : -1;

    return rows.slice().sort((a: IProjectListRow, b: IProjectListRow) => {
      const valueA: string | number = ProjectList._getSortValue(a, sortKey);
      const valueB: string | number = ProjectList._getSortValue(b, sortKey);

      if (valueA < valueB) {
        return -1 * direction;
      }
      if (valueA > valueB) {
        return 1 * direction;
      }
      return 0;
    });
  }

  private static _getSortValue(row: IProjectListRow, key: SortKey): string | number {
    if (key === 'name') {
      return row.name.toLowerCase();
    }
    if (key === 'maintenance') {
      return row.maintenance ? SORT_RANK[row.maintenance] : -1;
    }
    return SORT_RANK[row[key]];
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login`/`Menu` use for their own click handlers.
  private _onSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ search: event.target.value, currentPage: 1 });
  };

  private _onPageSizeChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ pageSize: Number(event.target.value), currentPage: 1 });
  };

  private _onToggleFavoritesOnly = (): void => {
    this.setState((state: IProjectListState) => ({ showFavoritesOnly: !state.showFavoritesOnly, currentPage: 1 }));
  };

  private _onToggleListView = (): void => {
    this.setState((state: IProjectListState) => ({ isListView: !state.isListView, currentPage: 1 }));
  };

  private _onSortClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const key: SortKey = event.currentTarget.dataset.key as SortKey;

    this.setState((state: IProjectListState) => {
      if (state.sortKey === key) {
        return { sortDirection: state.sortDirection === 'asc' ? 'desc' : 'asc' };
      }
      return { sortKey: key, sortDirection: 'asc' };
    });
  };

  private _onToggleFavorite = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const id: number = Number(event.currentTarget.dataset.id);

    this.setState((state: IProjectListState) => ({
      rows: state.rows.map((row: IProjectListRow) => row.id === id ? { ...row, isFavorite: !row.isFavorite } : row)
    }));
  };

  private _onPreviousPage = (): void => {
    this.setState((state: IProjectListState) => ({ currentPage: Math.max(1, state.currentPage - 1) }));
  };

  private _onNextPage = (): void => {
    this.setState((state: IProjectListState) => ({ currentPage: state.currentPage + 1 }));
  };

  private _onToggleRowMenu = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const id: number = Number(event.currentTarget.dataset.id);
    this.setState((state: IProjectListState) => ({ openMenuRowId: state.openMenuRowId === id ? undefined : id }));
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
