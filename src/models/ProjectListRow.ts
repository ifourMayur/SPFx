/**
 * The project overview table's row shape, its sample data, and the pure logic that folds a
 * saved form back into the list.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any
 * SPFx package. It lives here rather than inside `ProjectList` because three components
 * now share it - `ProjectList` renders the rows, `XsProject` owns them so they survive
 * navigating to the form and back, and `ProjectAddEdit` produces one on save.
 */

import { IBuildingForm } from './Building';

/** Traffic-light status shown as a circle for a single project phase. */
export type PhaseStatus = 'none' | 'success' | 'danger';

/** The five phase columns every project always shows a circle for. */
export type PhaseKey = 'administratie' | 'voorbereiding' | 'ontwerp' | 'uitvoering' | 'revisie';

/** One row of the project overview table. */
export interface IProjectListRow {
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

/**
 * Sample rows for the project overview table, matching the reference design.
 *
 * The Web API's `GET api/Building/GetProjectList` does not expose per-phase status, so
 * this component tree renders static sample data - replace it with a service call
 * (following the `Project`/`ProjectService` pattern) once that data is available.
 */
export const SAMPLE_PROJECT_ROWS: IProjectListRow[] = [
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

/**
 * The lowest unused row id.
 *
 * Derived from the highest existing id rather than the row count, so deleting a row can
 * never hand out an id that is already taken.
 */
export function nextProjectId(rows: IProjectListRow[]): number {
  return rows.reduce((highest: number, current: IProjectListRow): number => {
    return current.id > highest ? current.id : highest;
  }, 0) + 1;
}

/** Result of {@link upsertProjectRow}: the new list, and the id the project now has. */
export interface IUpsertResult {
  rows: IProjectListRow[];
  id: number;
}

/**
 * Folds a saved form into the row list, returning a new list.
 *
 * A form still in add mode (`id === 0`) becomes a new row under a freshly allocated id; an
 * existing project is renamed in place, keeping its position, its favorite star and its
 * phase circles - none of which the form edits. An id that is not in the list is added
 * rather than dropped, so a project can never be saved into nowhere.
 *
 * A new row goes to the **front** of the list. The listing pages at ten rows and the user
 * is returned to the first page after saving, so appending would hide the project they
 * just created on page two. Ordering is this component tree's own concern here - the rows
 * are local sample data, not a server-ordered page - so there is no server ordering to
 * contradict.
 */
export function upsertProjectRow(rows: IProjectListRow[], form: IBuildingForm): IUpsertResult {
  const name: string = form.buildingName.trim();
  const isExisting: boolean = form.id > 0 && rows.some((row: IProjectListRow) => row.id === form.id);

  if (isExisting) {
    return {
      id: form.id,
      rows: rows.map((row: IProjectListRow) => (row.id === form.id ? { ...row, name } : row))
    };
  }

  const id: number = form.id > 0 ? form.id : nextProjectId(rows);

  return {
    id,
    rows: [
      {
        id,
        name,
        isFavorite: false,
        administratie: 'none',
        voorbereiding: 'none',
        ontwerp: 'none',
        uitvoering: 'none',
        revisie: 'none'
      },
      ...rows
    ]
  };
}
