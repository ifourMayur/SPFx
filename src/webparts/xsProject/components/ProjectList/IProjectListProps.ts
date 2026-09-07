import { IProjectListRow } from '../../../../models/ProjectListRow';

export interface IProjectListProps {
  /**
   * Rows to show.
   *
   * Owned by `XsProject` rather than this component, so a project saved on the add/edit
   * form is still in the list when the router brings the user back here - navigating away
   * unmounts this component and would discard its own state.
   */
  rows: IProjectListRow[];
  /**
   * Whether this user may add or edit projects, from `canAddEditProject`. Hides the Add
   * button and the Edit row action when `false`, matching the reference's
   * `ViewBag.AddEditAccessRights`.
   */
  canAddEdit: boolean;
  /** Success message to show above the table, e.g. after a save. */
  notification?: string;
  /** Called when the user asks to create a project. */
  onAddProject: () => void;
  /** Called with the project to edit. */
  onEditProject: (projectId: number) => void;
  /** Called with the project whose favorite star was clicked. */
  onToggleFavorite: (projectId: number) => void;
  /** Clears {@link notification}, whether dismissed by the user or by its own timer. */
  onDismissNotification: () => void;
}
