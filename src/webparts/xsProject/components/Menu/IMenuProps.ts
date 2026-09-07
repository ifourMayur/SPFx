export interface IMenuProps {
  /** Path currently shown, used to highlight the matching menu item. */
  activePath: string;
  /** Called with the path the user selected. */
  onNavigate: (path: string) => void;
}
