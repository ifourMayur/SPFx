import { ISharePointSite } from '../../../../models/SharePointSite';

export interface ISharePointSitesState {
  /** True while the search query is in flight. */
  isLoading: boolean;
  /** Sites offered in the dropdown, already ordered by the service. */
  sites: ISharePointSite[];
  /** URL of the site currently selected in the dropdown, `''` for no selection. */
  selectedUrl: string;
  /**
   * Set when the query failed.
   *
   * Kept apart from an empty `sites` list: a tenant with no sites and a query that never
   * completed look identical in the data and mean entirely different things to the user.
   */
  hasFailed: boolean;
}
