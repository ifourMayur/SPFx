import { ISharePointSite } from '../../../../models/SharePointSite';
import { ISharePointSiteService } from '../../../../services/SharePointSiteService';

export interface ISharePointSitesProps {
  /**
   * Injected by the parent, which received it from the web part through `ServiceFactory`.
   * Keeps SharePoint REST access out of the component itself.
   */
  siteService: ISharePointSiteService;
  /**
   * URL of the site already in use, preselected in the dropdown.
   *
   * Set when the user came back here to change site, so the current choice is visible
   * rather than the picker looking as though nothing had been chosen.
   */
  selectedUrl?: string;
  /** Called with the site the user chose and confirmed. */
  onSiteSelected: (site: ISharePointSite) => void;
}
