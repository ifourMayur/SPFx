import { AppView } from '../../../../models/Navigation';

export interface IMenuProps {
  /** View currently shown, used to highlight the matching menu item. */
  activeView: AppView;
  /** Called with the view the user selected. */
  onNavigate: (view: AppView) => void;
}
