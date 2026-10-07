import { IProjectListingRow } from '../../../../models/ProjectListing';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectListingService } from '../../../../services/ProjectListingService';

export interface IProjectListProps {
  /**
   * The paged read behind the table: `Dashboard/GetListAllWithPagination`.
   *
   * The rows are fetched here rather than handed down as a prop, because the table is
   * server-paged - a page belongs to the search, sort and page number that asked for it,
   * so whoever owns those has to own the fetch too.
   */
  listingService: IProjectListingService;
  /** Supplies the phase columns through `getFolderColumns`. */
  lookupService: ILookupService;
  /**
   * Whether this user may add or edit projects, from `canAddEditProject`. Hides the Add
   * button and the Edit row action when `false`, matching the reference's
   * `ViewBag.AddEditAccessRights`.
   */
  canAddEdit: boolean;
  /**
   * `ILoginResponse.domainGroup`, sent with every page request.
   *
   * The one field of the request the Web API does not derive from the caller's token - the
   * reference reads it from the MVC session - and it decides how every cell in the table
   * is coloured.
   */
  domainGroup?: number;
  /** Success message to show as a self-dismissing toast, e.g. after a save. */
  notification?: string;
  /**
   * Absolute URL of the SharePoint site chosen in `SharePointSites`, named at the top of
   * the listing.
   *
   * The shell's site bar names the site by its title, which does not say which tenant or
   * site collection it is - and this listing is where a project is created from, so it is
   * where knowing the site a project's folders will be created in matters. `''` when no
   * site URL is known, which hides the line rather than labelling nothing.
   */
  siteUrl: string;
  /** Called when the user asks to create a project. */
  onAddProject: () => void;
  /** Called with the project to edit. */
  onEditProject: (projectId: number) => void;
  /** Clears {@link notification}, whether dismissed by the user or by the toast's timer. */
  onDismissNotification: () => void;
  /**
   * Reports the rows of the page just loaded.
   *
   * `XsProject` keeps their names so the edit form can title itself for a project opened
   * straight from this listing - routing unmounts this component on the way there, so the
   * name has to be handed to something that outlives it.
   */
  onProjectsLoaded: (rows: IProjectListingRow[]) => void;
}
