import * as React from 'react';

import styles from './Menu.module.scss';
import type { IMenuProps } from './IMenuProps';
import type { IMenuState } from './IMenuState';
import { AppView, IMenuItem } from '../../../../models/Navigation';

/**
 * Items shown in the shared menu. Add a leaf entry (with `view`) for a new top-level page,
 * or add to `children` for a new submenu entry - and a matching `AppView` member in
 * `src/models/Navigation.ts` - to extend the menu; nothing else in this component changes.
 */
const MENU_ITEMS: IMenuItem[] = [
  {
    key: 'project',
    label: 'Project',
    view: 'projectList',
    children: [
      { key: 'projectDashboard', label: 'Dashboard', view: 'projectDashboard' },
      { key: 'projectDocument', label: 'Document', view: 'projectDocument' }
    ]
  },
  { key: 'search', label: 'Search', view: 'search' },
  { key: 'suppliers', label: 'Suppliers', view: 'suppliers' }
];

/**
 * Shared, SharePoint-style left navigation for the web part's authenticated views.
 *
 * Carries no authentication logic of its own - it is only ever rendered by `XsProject`
 * once `Login` reports a successful sign-in - so it stays reusable wherever an
 * authenticated view needs a menu, independent of how that authentication happened.
 *
 * Collapses to a single toggle icon, matching SharePoint's own left nav. Whether it is
 * collapsed and which submenu is expanded are both local UI state - they never affect what
 * `activeView` the parent has selected.
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
    const { activeView } = this.props;
    const hasChildren: boolean = !!item.children && item.children.length > 0;
    const isActive: boolean = item.view === activeView
      || (!!item.children && item.children.some((child: IMenuItem) => child.view === activeView));
    const isOpen: boolean = this.state.openKey === item.key;

    return (
      <li key={item.key}>
        <button
          type="button"
          className={`${styles.item} ${isActive ? styles.active : ''}`}
          aria-haspopup={hasChildren ? 'menu' : undefined}
          aria-expanded={hasChildren ? isOpen : undefined}
          aria-current={item.view === activeView ? 'page' : undefined}
          data-key={item.key}
          data-view={item.view}
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
                  className={`${styles.subItem} ${child.view === activeView ? styles.subItemActive : ''}`}
                  aria-current={child.view === activeView ? 'page' : undefined}
                  data-view={child.view}
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
    const { key, view, hasChildren } = event.currentTarget.dataset;

    if (view) {
      this.props.onNavigate(view as AppView);
    }

    if (hasChildren) {
      this.setState((state: IMenuState) => ({ openKey: state.openKey === key ? undefined : key }));
    }
  };

  private _onChildClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    // Deliberately does not close the submenu: the parent group (e.g. "Project") stays
    // expanded after navigating so the newly active child item remains visible/highlighted
    // within it, rather than the whole group collapsing out from under the selection.
    this.props.onNavigate(event.currentTarget.dataset.view as AppView);
  };
}
