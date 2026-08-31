export interface IMenuState {
  /** Key of the top-level item whose submenu is expanded, if any. */
  openKey?: string;
  /** True when the sidebar is collapsed down to just its toggle icon. */
  isCollapsed: boolean;
}
