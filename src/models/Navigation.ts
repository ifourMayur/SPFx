/**
 * View-state shapes for in-page navigation between the web part's authenticated views.
 *
 * Part of the models layer: no imports from `config`, `services` or `components`, so the
 * shared `Menu` component and whatever renders alongside it agree on the same navigation
 * contract without depending on each other directly.
 */

/** A view the shared `Menu` component can navigate the app to. Add a member here to add a view. */
export type AppView = 'projectList' | 'projectDashboard' | 'projectDocument' | 'search' | 'suppliers';

/**
 * A single entry in the shared `Menu`.
 *
 * A leaf item sets `view` and navigates directly. A parent item (SharePoint-style, e.g.
 * "Project") sets `children` to open a submenu, and may also set its own `view` so clicking
 * the parent itself navigates (in addition to opening the submenu) rather than only toggling it.
 */
export interface IMenuItem {
  /** Identifies the item within the menu; equal to `view` for a leaf item. */
  key: string;
  label: string;
  view?: AppView;
  children?: IMenuItem[];
}
