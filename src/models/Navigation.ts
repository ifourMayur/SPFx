/**
 * The web part's route table and the shape of a shared-menu entry.
 *
 * Part of the models layer: no imports from `config`, `services` or `components`, so the
 * shared `Menu`, the router and every routed page agree on the same navigation contract
 * without depending on each other directly.
 *
 * Routing uses `HashRouter` (see `components/XsProject/AppShell.tsx`). A SPFx web part is
 * a guest on a SharePoint page whose path SharePoint itself owns and navigates, so the
 * app's own routes live in the fragment - `.../workbench.aspx#/projects/add` - where they
 * are real, linkable URLs that never trigger a page load.
 */

/**
 * Every path the app routes, declared once.
 *
 * `projectEdit` is a pattern for `<Route path>`; build a concrete one with
 * {@link projectEditPath} rather than interpolating the `:id` by hand.
 */
export const ROUTE_PATHS = {
  projectList: '/projects',
  projectAdd: '/projects/add',
  projectEdit: '/projects/edit/:id',
  projectDashboard: '/dashboard',
  projectDocument: '/document',
  search: '/search',
  suppliers: '/suppliers'
} as const;

/** A path this app routes. */
export type RoutePath = (typeof ROUTE_PATHS)[keyof typeof ROUTE_PATHS];

/**
 * Where `/` lands. Matches the view the web part opened on before it had routes, so
 * signing in still shows the dashboard.
 */
export const DEFAULT_ROUTE_PATH: string = ROUTE_PATHS.projectDashboard;

/** The edit route for one project, e.g. `projectEditPath(12)` -> `/projects/edit/12`. */
export function projectEditPath(projectId: number): string {
  return `${ROUTE_PATHS.projectList}/edit/${projectId}`;
}

/**
 * A single entry in the shared `Menu`.
 *
 * A leaf item sets `path` and navigates directly. A parent item (SharePoint-style, e.g.
 * "Project") sets `children` to open a submenu, and may also set its own `path` so
 * clicking the parent itself navigates (in addition to opening the submenu) rather than
 * only toggling it.
 */
export interface IMenuItem {
  /** Identifies the item within the menu; usually the last segment of `path`. */
  key: string;
  label: string;
  path?: string;
  children?: IMenuItem[];
}
