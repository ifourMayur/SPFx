import * as React from 'react';

import styles from './Menu.module.scss';
import type { IMenuProps } from './IMenuProps';
import type { IMenuState } from './IMenuState';
import { IMenuItem, ROUTE_PATHS } from '../../../../models/Navigation';

/**
 * Items shown in the shared menu. Add a leaf entry (with `path`) for a new top-level page,
 * or add to `children` for a new submenu entry - and a matching entry in `ROUTE_PATHS` plus
 * a `<Route>` in `AppShell` - to extend the menu; nothing else in this component changes.
 */
const MENU_ITEMS: IMenuItem[] = [
  {
    key: 'project',
    label: 'Project',
    path: ROUTE_PATHS.projectList,
    children: [
      // No "Add Project" entry: creating a project is permission-dependent, and the
      // gated button on the listing is the reference's only way in. A menu entry here
      // would offer restricted roles a link that only leads to a refusal.
      { key: 'projectDashboard', label: 'Dashboard', path: ROUTE_PATHS.projectDashboard },
      { key: 'projectDocument', label: 'Document', path: ROUTE_PATHS.projectDocument }
    ]
  },
  { key: 'search', label: 'Search', path: ROUTE_PATHS.search },
  { key: 'suppliers', label: 'Suppliers', path: ROUTE_PATHS.suppliers }
];

/**
 * Shared, SharePoint-style left navigation for the web part's authenticated views.
 *
 * This is the one menu every page uses. `AppShell` renders it once, outside the router's
 * `<Routes>`, so no page contains menu markup of its own and none can drift from the rest.
 * Its entries live in a single `MENU_ITEMS` tree, and it is driven entirely by two props
 * (`activePath` in, `onNavigate` out) so it depends on no particular router or page.
 *
 * Carries no authentication logic of its own - it is only ever rendered once `Login`
 * reports a successful sign-in - so it stays reusable wherever an authenticated view needs
 * a menu, independent of how that authentication happened.
 *
 * Collapses to a single toggle icon, matching SharePoint's own left nav. Whether it is
 * collapsed and which submenu is expanded are both local UI state - they never affect
 * which path the parent has navigated to.
 */
export default class Menu extends React.Component<IMenuProps, IMenuState> {
  public constructor(props: IMenuProps) {
    super(props);

    this.state = { openKey: undefined, isCollapsed: false };
  }

  public render(): React.ReactElement<IMenuProps> {
    const { isCollapsed } = this.state;

    return (
      <aside className={`${styles.sidebar} ${isCollapsed ? styles.collapsed : ''}`} aria-label="XSProject">
        <div className={styles.header}>
          <button
            type="button"
            className={styles.toggle}
            aria-label={isCollapsed ? 'Expand menu' : 'Collapse menu'}
            aria-expanded={!isCollapsed}
            onClick={this._onToggleClick}
          >
            <span className={styles.menuIcon} aria-hidden="true">&#9776;</span>
          </button>

          {!isCollapsed && (
            <div className={styles.brand}>
              <span className={styles.logo} aria-hidden="true">XS</span>
              <span className={styles.title}>XSProject</span>
            </div>
          )}
        </div>

        {!isCollapsed && (
          <ul className={styles.list}>
            {MENU_ITEMS.map((item: IMenuItem) => this._renderItem(item))}
          </ul>
        )}
      </aside>
    );
  }

  private _renderItem(item: IMenuItem): React.ReactElement {
    const { activePath } = this.props;
    const hasChildren: boolean = !!item.children && item.children.length > 0;
    const isActive: boolean = Menu._isActive(item, activePath)
      || (!!item.children && item.children.some((child: IMenuItem) => Menu._isActive(child, activePath)));
    const isOpen: boolean = this.state.openKey === item.key;

    return (
      <li key={item.key}>
        <button
          type="button"
          className={`${styles.item} ${isActive ? styles.active : ''}`}
          aria-haspopup={hasChildren ? 'menu' : undefined}
          aria-expanded={hasChildren ? isOpen : undefined}
          aria-current={Menu._isActive(item, activePath) ? 'page' : undefined}
          data-key={item.key}
          data-path={item.path}
          data-has-children={hasChildren ? 'true' : undefined}
          onClick={this._onItemClick}
        >
          <span className={styles.itemLabel}>{item.label}</span>
          {hasChildren && <span className={styles.chevron} aria-hidden="true">{isOpen ? '▴' : '▾'}</span>}
        </button>

        {hasChildren && isOpen && (
          <ul className={styles.submenu} role="menu">
            {item.children && item.children.map((child: IMenuItem) => (
              <li key={child.key} role="none">
                <button
                  type="button"
                  role="menuitem"
                  className={`${styles.subItem} ${Menu._isActive(child, activePath) ? styles.subItemActive : ''}`}
                  aria-current={Menu._isActive(child, activePath) ? 'page' : undefined}
                  data-path={child.path}
                  onClick={this._onChildClick}
                >
                  {child.label}
                </button>
              </li>
            ))}
          </ul>
        )}
      </li>
    );
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` uses for its own click handler.
  private _onToggleClick = (): void => {
    this.setState((state: IMenuState) => ({ isCollapsed: !state.isCollapsed }));
  };

  // Handles every top-level item: navigates when the item (leaf or parent) has its own
  // `view`, and toggles the submenu open/closed when it has `children` - both can apply to
  // the same click, e.g. "Project" navigates to its own page and opens its submenu.
  private _onItemClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const { key, path, hasChildren } = event.currentTarget.dataset;

    if (path) {
      this.props.onNavigate(path);
    }

    if (hasChildren) {
      this.setState((state: IMenuState) => ({ openKey: state.openKey === key ? undefined : key }));
    }
  };

  private _onChildClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    // Deliberately does not close the submenu: the parent group (e.g. "Project") stays
    // expanded after navigating so the newly active child item remains visible/highlighted
    // within it, rather than the whole group collapsing out from under the selection.
    this.props.onNavigate(event.currentTarget.dataset.path as string);
  };

  /**
   * Whether an item is the one currently shown.
   *
   * An item matches its own path and anything nested below it, so "Project"
   * (`/projects`) stays highlighted while the user is on `/projects/add` or
   * `/projects/edit/12`. Matching on a `/` boundary keeps a sibling such as
   * `/projects-archive` from counting as nested.
   */
  private static _isActive(item: IMenuItem, activePath: string): boolean {
    if (!item.path) {
      return false;
    }

    return activePath === item.path || activePath.indexOf(`${item.path}/`) === 0;
  }
}
