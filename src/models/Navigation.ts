/**
 * View-state shapes for in-page navigation between the web part's authenticated views.
 *
 * Part of the models layer: no imports from `config`, `services` or `components`, so the
 * shared `Menu` component and whatever renders alongside it agree on the same navigation
 * contract without depending on each other directly.
 */

/** A view the shared `Menu` component can navigate the app to. Add a member here to add a view. */
export type AppView = 'projectDashboard' | 'projectDocument' | 'search' | 'suppliers';

/**
 * A single entry in the shared `Menu`.
 *
 * A leaf item sets `view` and navigates directly. A parent item (SharePoint-style, e.g.
 * "Project") omits `view` and sets `children` instead - selecting it opens a submenu rather
 * than navigating on its own.
 */
export interface IMenuItem {
  /** Identifies the item within the menu; equal to `view` for a leaf item. */
  key: string;
  label: string;
  view?: AppView;
  children?: IMenuItem[];
}
