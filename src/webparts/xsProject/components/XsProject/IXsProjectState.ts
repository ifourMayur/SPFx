import { ILoginResponse } from '../../../../models/Auth';
import { IBuildingForm } from '../../../../models/Building';
import { IProjectListRow } from '../../../../models/ProjectListRow';

export interface IXsProjectState {
  /** Set from `Login`'s `onLoginSucceeded`/`onLoginFailed` callbacks; gates the whole app. */
  isAuthenticated: boolean;
  /**
   * The sign-in response, kept for the identity and entitlement facts the app needs -
   * `userRoleId` for the permission gates, `domainGroup` for the tenant feature gates,
   * and the client id/name shown on the project form.
   *
   * The application token it also carries is not used anywhere yet.
   */
  login?: ILoginResponse;
  /**
   * Project rows shown by `ProjectList`.
   *
   * Owned here rather than by `ProjectList` because routing to the add/edit form unmounts
   * that component - a project saved on the form has to survive the round trip.
   */
  rows: IProjectListRow[];
  /**
   * Forms saved during this session, keyed by project id, so a project can be reopened
   * and edited with its values intact.
   *
   * Only covers projects saved here: the sample rows have no stored form, which is why
   * `ProjectAddEdit` falls back to just the name. A `GET api/Building?id=` call would
   * replace this store.
   */
  savedForms: { [projectId: number]: IBuildingForm };
  /** Success message to show on the listing after a save. */
  notification?: string;
}
