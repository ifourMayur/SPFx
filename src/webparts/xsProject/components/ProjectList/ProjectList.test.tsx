/**
 * Render tests for the project overview table.
 *
 * The table is server-paged, exactly as the reference's DataTable is: search, sort and
 * paging are round trips, not array operations. So most of what is worth asserting here is
 * *what the component asked the server for* in response to a click - a client-side filter
 * would render identically on the first page and then be wrong on the second.
 *
 * `react-dom/test-utils` and the jsdom environment are already available, so this needs no
 * new dependency. Every `ReactDOM.render` here is unmounted in `afterEach`.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import ProjectList from './ProjectList';
import { IProjectListProps } from './IProjectListProps';
import { ILookupOption } from '../../../../models/Building';
import {
  IProjectListPage,
  IProjectListQuery,
  IProjectListingRow
} from '../../../../models/ProjectListing';
import type { ILookupService } from '../../../../services/LookupService';
import type { IProjectListingService } from '../../../../services/ProjectListingService';

const SITE_URL: string = 'https://ifourtechnolab.sharepoint.com/sites/Projects';

/** The phase columns `Folder/GetFolderWithoutBIM` answers with. */
const COLUMNS: ILookupOption[] = [
  { id: '1', name: 'Administratie' },
  { id: '2', name: 'Voorbereiding' }
];

/**
 * A page as the service hands it over: two rows out of a collection of 137. The gap
 * between those two numbers is what proves the paging is the server's.
 */
function page(overrides?: Partial<IProjectListPage>): IProjectListPage {
  return {
    recordsTotal: 137,
    rows: [
      {
        buildingId: 41,
        encryBuildingId: 'NDE=',
        buildingName: 'Office Renovation Plan',
        isFavorite: true,
        contactPerson: 'Jane Dijkstra',
        maintenanceFolderId: 6,
        isMaintenanceFolderAvailable: true,
        cells: [
          {
            folderId: 1,
            encryFolderId: 'MQ==',
            folderName: 'Administratie',
            colorCode: '#22b14c',
            borderColorCode: '2.0px solid #d9d9d9',
            isAccessRights: true
          },
          {
            folderId: 2,
            encryFolderId: 'Mg==',
            folderName: 'Voorbereiding',
            colorCode: '#ff0000',
            borderColorCode: '2.0px solid #d9d9d9',
            isAccessRights: false
          }
        ]
      },
      {
        buildingId: 42,
        encryBuildingId: 'NDI=',
        buildingName: 'Tender Response A12',
        isFavorite: false,
        contactPerson: '',
        maintenanceFolderId: 6,
        isMaintenanceFolderAvailable: false,
        cells: []
      }
    ],
    ...overrides
  };
}

/** Every query the component sent, in order. */
let queries: IProjectListQuery[];

/** A listing service that records what it was asked and answers with {@link page}. */
function stubListingService(result?: IProjectListPage | Error): IProjectListingService {
  return {
    getProjectPage: async (query: IProjectListQuery): Promise<IProjectListPage> => {
      queries.push(query);

      if (result instanceof Error) {
        throw result;
      }

      return result || page();
    }
  };
}

/** Only `getFolderColumns` is reached from this screen; the rest never run. */
function stubLookupService(columns: ILookupOption[] = COLUMNS): ILookupService {
  const unused = async (): Promise<ILookupOption[]> => [];

  return {
    getProjectTemplates: unused,
    getUsers: unused,
    getSpatialBreakdowns: unused,
    getTenderTemplates: unused,
    getSuppliers: unused,
    getProjectFolders: unused,
    getProjectSubFolders: unused,
    getProjectSubSubFolders: unused,
    getProjectSubSubSubFolders: unused,
    getFolderLevel: unused,
    getFolderColumns: async (): Promise<ILookupOption[]> => columns
  };
}

function props(overrides?: Partial<IProjectListProps>): IProjectListProps {
  return {
    listingService: stubListingService(),
    lookupService: stubLookupService(),
    canAddEdit: true,
    siteUrl: SITE_URL,
    domainGroup: 1,
    onAddProject: (): void => undefined,
    onEditProject: (): void => undefined,
    onDismissNotification: (): void => undefined,
    onProjectsLoaded: (): void => undefined,
    ...overrides
  };
}

describe('ProjectList', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    queries = [];
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  /** Renders and lets the columns and the first page settle. */
  async function render(overrides?: Partial<IProjectListProps>): Promise<void> {
    await act(async () => {
      ReactDOM.render(<ProjectList {...props(overrides)} />, host);
    });
  }

  /** Clicks an element and lets the refetch it triggers settle. */
  async function click(element: Element | null): Promise<void> {
    expect(element).not.toBeNull();

    await act(async () => {
      (element as HTMLElement).click();
    });
  }

  /**
   * Types into an input the way a user does.
   *
   * Through the prototype's own setter: React keeps a value tracker per input and skips
   * `onChange` when a plain `element.value = …` has already moved it.
   */
  function setValue(element: HTMLInputElement | HTMLSelectElement, value: string): void {
    const prototype = element instanceof HTMLSelectElement
      ? HTMLSelectElement.prototype
      : HTMLInputElement.prototype;

    Object.getOwnPropertyDescriptor(prototype, 'value')?.set?.call(element, value);
    element.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function findButton(label: string): HTMLButtonElement | null {
    const buttons: HTMLButtonElement[] = Array.from(host.querySelectorAll('button'));
    return buttons.filter((button) => (button.textContent || '').indexOf(label) === 0)[0] || null;
  }

  describe('binding the page', () => {
    it('loads a page of projects on mount', async () => {
      await render();

      expect(host.textContent).toContain('Office Renovation Plan');
      expect(host.textContent).toContain('Tender Response A12');
    });

    it('builds a column per folder the API reports', async () => {
      await render();

      const headers: string[] = Array.from(host.querySelectorAll('thead th')).map(
        (cell) => cell.textContent || ''
      );

      expect(headers.filter((header) => header.indexOf('Administratie') > -1)).toHaveLength(1);
      expect(headers.filter((header) => header.indexOf('Voorbereiding') > -1)).toHaveLength(1);
    });

    /**
     * The colour is a document-completion percentage the server computed; rendering it
     * from a local three-state enum would throw every intermediate value away.
     */
    it('paints a cell with the colour the server computed', async () => {
      await render();

      const dot: HTMLElement | null = host.querySelector('[aria-label="Administratie"]');

      expect(dot).not.toBeNull();
      expect(dot!.style.backgroundColor).toBe('rgb(34, 177, 76)');
    });

    /** `else if (cell.isAccessRights == true)` in the reference view. */
    it('shows a dash instead of a colour where the user has no access rights', async () => {
      await render();

      expect(host.querySelector('[aria-label="Voorbereiding"]')).toBeNull();
    });

    /**
     * The count that drives the pager is the collection's, stamped onto the rows - not the
     * length of the page. Two rows of 137 is the whole point of server-side paging.
     */
    it('reports the collection total rather than the number of rows on screen', async () => {
      await render();

      expect(host.textContent).toContain('Showing 1 to 2 of 137 entries');
    });

    /**
     * `XsProject` keeps these so the edit form can title itself for a project the user
     * opens straight from the listing.
     */
    it('reports the loaded projects to its owner', async () => {
      let reported: IProjectListingRow[] = [];

      await render({ onProjectsLoaded: (rows: IProjectListingRow[]): void => { reported = rows; } });

      expect(reported.map((row) => row.buildingName)).toEqual([
        'Office Renovation Plan',
        'Tender Response A12'
      ]);
    });
  });

  describe('asking the server for a different page', () => {
    it('requests the first page, unsorted, on mount', async () => {
      await render();

      expect(queries[0]).toMatchObject({
        page: 1,
        search: '',
        sortDirection: '',
        showFavoritesOnly: false,
        domainGroup: 1
      });
    });

    it('fetches the next page rather than slicing the rows it has', async () => {
      await render();
      await click(findButton('Next'));

      expect(queries).toHaveLength(2);
      expect(queries[1].page).toBe(2);
    });

    it('returns to the first page when the page size changes', async () => {
      await render();
      await click(findButton('Next'));

      await act(async () => {
        setValue(host.querySelector('select')!, '50');
      });

      expect(queries[queries.length - 1]).toMatchObject({ page: 1, pageSize: 50 });
    });

    /**
     * Only the project name can be sorted: `GetDashboardListWithPagination` branches on
     * `SortColumnDirection` and orders by `BuildingName` either way.
     */
    it('sorts by name through the server', async () => {
      await render();

      const header: HTMLButtonElement | null = findButton('Naam van het project');
      await click(header);

      expect(queries[1].sortDirection).toBe('asc');

      await click(findButton('Naam van het project'));

      expect(queries[2].sortDirection).toBe('desc');
    });

    it('asks the server for favourites only, rather than filtering the page it has', async () => {
      await render();

      const toggle: HTMLInputElement = host.querySelectorAll('input[type="checkbox"]')[0] as HTMLInputElement;
      await act(async () => {
        toggle.click();
      });

      expect(queries[1]).toMatchObject({ showFavoritesOnly: true, page: 1 });
    });

    /**
     * The reference's DataTable debounces its search box too. Without it every keystroke
     * is a round trip through a query that pages the whole Buildings table.
     */
    it('searches on the server, once the typing stops', async () => {
      jest.useFakeTimers();

      try {
        await render();

        await act(async () => {
          setValue(host.querySelector('input[type="search"]')!, 'renov');
        });

        expect(queries).toHaveLength(1);

        await act(async () => {
          jest.runOnlyPendingTimers();
        });

        expect(queries[1]).toMatchObject({ search: 'renov', page: 1 });
      } finally {
        jest.useRealTimers();
      }
    });
  });

  describe('when the API cannot be reached', () => {
    /**
     * An empty table would be indistinguishable from a tenant with no projects, which is
     * the failure this whole layer is built to make visible.
     */
    it('says so instead of rendering an empty table', async () => {
      await render({ listingService: stubListingService(new Error('Invalid data.')) });

      expect(host.textContent).toContain('Invalid data.');
    });
  });

  describe('naming the site the session works in', () => {
    it('names the URL of the site the session is working in', async () => {
      await render();

      expect(host.textContent).toContain(SITE_URL);
    });

    /**
     * The URL is shown for the user to read and to copy into a browser, so it is rendered
     * as a link to the site rather than as text that has to be retyped.
     */
    it('links to the site', async () => {
      await render();

      const link: HTMLAnchorElement | null = host.querySelector(`a[href="${SITE_URL}"]`);

      expect(link).not.toBeNull();
      expect(link!.textContent).toBe(SITE_URL);
    });

    /**
     * The listing is rendered by `AppShell` only once a site has been chosen, so an empty
     * URL should not happen - but an empty label reading "Site URL:" with nothing after it
     * would be worse than no label at all.
     */
    it('omits the line entirely when no site URL is known', async () => {
      await render({ siteUrl: '' });

      expect(host.textContent).not.toContain('Site URL:');
    });
  });
});
