import { ILoginResponse } from '../../../../models/Auth';
import { IBuildingForm } from '../../../../models/Building';
import { ISharePointSite } from '../../../../models/SharePointSite';

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
   * The site chosen in `SharePointSites`; gates the app the way `isAuthenticated` does.
   *
   * Owned here rather than by the picker because it outlives it: the picker unmounts as
   * soon as the app renders.
   */
  selectedSite?: ISharePointSite;
  /**
   * True while the user is back in the picker to change site.
   *
   * A flag rather than clearing `selectedSite`, so the picker can show the site currently
   * in use as its selection instead of presenting itself as a first-time choice.
   */
  isChangingSite?: boolean;
  /**
   * Names of the projects seen on the listing this session, keyed by project id.
   *
   * `ProjectList` is server-paged and refetches on mount, so it needs nothing kept for it
   * here - but routing to the add/edit form unmounts it, and the form wants the name of
   * the project being edited. That name is all that has to survive the round trip.
   */
  projectNames: { [projectId: number]: string };
  /**
   * Forms saved during this session, keyed by project id, so a project can be reopened
   * and edited with its values intact.
   *
   * Only covers projects saved here: a project merely listed has no stored form, which is
   * why `ProjectAddEdit` falls back to just the name. A `GET api/Building?id=` call would
   * replace this store.
   */
  savedForms: { [projectId: number]: IBuildingForm };
  /**
   * Save confirmation, shown on the listing as a toast.
   *
   * Owned here for the same reason the rows are: the save happens on the form, which
   * routing unmounts before the listing renders the result.
   */
  notification?: string;
}
