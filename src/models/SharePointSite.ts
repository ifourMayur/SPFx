/**
 * One SharePoint site the user may choose, and the mapping of a SharePoint Search
 * response onto a list of them.
 *
 * Part of the models layer: pure, with no imports outside `models`, so the shape the
 * dropdown binds to is independent of how the sites were fetched and the flattening below
 * is unit-testable without a tenant.
 */

/** A site collection the signed-in user can open. */
export interface ISharePointSite {
  /** Display name, or the URL when the site has no title of its own. */
  title: string;
  /** Absolute URL of the site, without a trailing slash. */
  url: string;
  /** The site collection's GUID as search reported it, or `''` when it did not. */
  siteId: string;
}

/**
 * Managed properties requested from search, and therefore the cell keys read below.
 *
 * `Path` is the site's URL: search names the address of a result `Path`, not `Url`.
 */
export const SITE_SELECT_PROPERTIES: string[] = ['Title', 'Path', 'SiteId'];

/**
 * A site whose URL contains one of these is somebody's OneDrive.
 *
 * `contentclass:STS_Site` matches every site collection the user can see, and in most
 * tenants that includes their own personal site - not something anyone would pick as a
 * project's site.
 */
const PERSONAL_SITE_MARKERS: string[] = ['-my.sharepoint.com', '/personal/'];

/** Narrows an unknown value to an indexable object, so a missing branch reads as empty. */
function record(value: unknown): { [key: string]: unknown } {
  return typeof value === 'object' && value !== null ? (value as { [key: string]: unknown }) : {};
}

/**
 * Reads a search collection whichever way the response wrapped it.
 *
 * `odata=nometadata` gives a plain array; `odata=verbose` wraps every array in a
 * `results` property. Both are accepted so the mapper does not silently return nothing
 * when a transport sends a different `Accept` header than the one this solution asks for.
 */
function collection(value: unknown): unknown[] {
  if (Array.isArray(value)) {
    return value;
  }

  const results: unknown = record(value).results;

  return Array.isArray(results) ? results : [];
}

/** The result rows of a search response, or an empty list for any other payload. */
function rowsOf(response: unknown): unknown[] {
  const root: { [key: string]: unknown } = record(response);
  // The verbose envelope nests the whole answer under `d.query`.
  const query: { [key: string]: unknown } = root.PrimaryQueryResult
    ? root
    : record(record(root.d).query);

  const table: { [key: string]: unknown } = record(
    record(record(query.PrimaryQueryResult).RelevantResults).Table
  );

  return collection(table.Rows);
}

/** One row's `{ Key, Value }` cells, flattened into a plain lookup by managed property. */
function cellsOf(row: unknown): { [key: string]: string } {
  const values: { [key: string]: string } = {};

  collection(record(row).Cells).forEach((cell: unknown): void => {
    const entry: { [key: string]: unknown } = record(cell);
    const key: unknown = entry.Key;

    if (typeof key === 'string') {
      values[key] = typeof entry.Value === 'string' ? entry.Value.trim() : '';
    }
  });

  return values;
}

/**
 * Drops a single trailing slash, so the tenant root and a site read alike.
 *
 * Exported because the same normalization decides where a request is addressed:
 * `SharePointSiteService` appends `/_api/search/query` to a site URL, and `//_api` is a
 * 404.
 */
export function withoutTrailingSlash(url: string): string {
  return url.length > 1 && url.charAt(url.length - 1) === '/' ? url.substring(0, url.length - 1) : url;
}

function isPersonalSite(url: string): boolean {
  const lower: string = url.toLowerCase();

  return PERSONAL_SITE_MARKERS.some((marker: string): boolean => lower.indexOf(marker) >= 0);
}

/**
 * Maps a SharePoint Search response onto the sites the dropdown offers.
 *
 * Search answers a column store ordered by rank, so this both flattens the rows and puts
 * them in the order someone reading a dropdown expects. Anything unusable is dropped
 * rather than allowed to produce a broken option: a payload that is not a search response
 * yields an empty list (a failed load renders an empty dropdown instead of throwing), a
 * row with no `Path` is skipped because nothing could open it, and a row with no `Title`
 * is named by its URL.
 *
 * Duplicates are removed by URL, ignoring case and a trailing slash: the query passes
 * `trimduplicates=false` - deliberately, since search's own de-duplication discards whole
 * sites it judges near-identical - which leaves the same site free to come back twice.
 */
export function toSharePointSites(response: unknown): ISharePointSite[] {
  const sites: ISharePointSite[] = [];
  const seen: { [url: string]: boolean } = {};

  rowsOf(response).forEach((row: unknown): void => {
    const cells: { [key: string]: string } = cellsOf(row);
    const url: string = withoutTrailingSlash(cells.Path || '');

    if (!url || isPersonalSite(url)) {
      return;
    }

    const key: string = url.toLowerCase();

    if (seen[key]) {
      return;
    }

    seen[key] = true;
    sites.push({ title: cells.Title || url, url, siteId: cells.SiteId || '' });
  });

  return sites.sort((left: ISharePointSite, right: ISharePointSite): number =>
    left.title.toLowerCase().localeCompare(right.title.toLowerCase())
  );
}
