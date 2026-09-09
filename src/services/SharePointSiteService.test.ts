/**
 * Tests for the site list behind the `SharePointSites` dropdown: the real
 * `SharePointSiteService` over a fake SharePoint REST transport.
 *
 * The query itself is the thing worth pinning. Its parameters are not decoration -
 * `contentclass:STS_Site` is what makes the result site collections rather than documents,
 * `selectproperties` is what makes the cells the mapper reads come back at all, and
 * `trimduplicates=false` is what stops search discarding whole sites it judges
 * near-identical - and every one of them fails silently, as an empty or wrong dropdown,
 * rather than as an error.
 */

import { ISharePointSite } from '../models/SharePointSite';
// Type-only: the module also exports the SPFx-backed transport, whose `@microsoft/sp-http`
// import cannot load under jsdom. `import type` is erased, so nothing is required here.
import type { ISpHttpClient } from './SharePointHttpClient';
import { ISharePointSiteService, SharePointSiteService } from './SharePointSiteService';

/** One `Table.Rows` entry, as the `odata=nometadata` search response shapes it. */
function row(title: string, path: string): unknown {
  return {
    Cells: [
      { Key: 'Title', Value: title },
      { Key: 'Path', Value: path }
    ]
  };
}

/** A complete search response carrying the given rows. */
function searchResponse(rows: unknown[]): unknown {
  return {
    PrimaryQueryResult: {
      RelevantResults: { Table: { Rows: rows }, TotalRows: rows.length }
    }
  };
}

/**
 * Transport that answers every request with one body, recording the URLs it was asked for.
 *
 * `postJson` throws rather than being omitted: listing sites is a read, so a write here
 * would be a bug, and a stub that fails says so instead of quietly allowing it.
 */
function fakeClient(body: unknown, urls: string[]): ISpHttpClient {
  return {
    getJson: async <T>(url: string): Promise<T> => {
      urls.push(url);
      return body as T;
    },
    postJson: async <T>(): Promise<T> => {
      throw new Error('SharePointSiteService must not write.');
    }
  };
}

const WEB_URL: string = 'https://contoso.sharepoint.com/sites/intranet';

describe('SharePointSiteService', () => {
  it('queries the search endpoint of the current site', async () => {
    const urls: string[] = [];
    const service: ISharePointSiteService = new SharePointSiteService(
      fakeClient(searchResponse([]), urls),
      WEB_URL
    );

    await service.getSites();

    expect(urls).toHaveLength(1);
    expect(urls[0].indexOf(`${WEB_URL}/_api/search/query?`)).toBe(0);
  });

  it('asks for site collections, the three cells the mapper reads, and no duplicate trimming', async () => {
    const urls: string[] = [];
    const service: ISharePointSiteService = new SharePointSiteService(
      fakeClient(searchResponse([]), urls),
      WEB_URL
    );

    await service.getSites();

    // Decoded, so the assertion reads as the query search actually receives rather than as
    // percent escapes.
    const query: string = decodeURIComponent(urls[0]);

    expect(query).toContain("querytext='contentclass:STS_Site'");
    expect(query).toContain("selectproperties='Title,Path,SiteId'");
    expect(query).toContain('trimduplicates=false');
    expect(query).toContain('rowlimit=500');
  });

  // The site URL arrives from `pageContext.web.absoluteUrl`, which is normally unslashed,
  // but a configured override is exactly the sort of value that carries one - and
  // `//_api/search` is a 404.
  it('does not double the slash when the site URL ends in one', async () => {
    const urls: string[] = [];
    const service: ISharePointSiteService = new SharePointSiteService(
      fakeClient(searchResponse([]), urls),
      `${WEB_URL}/`
    );

    await service.getSites();

    expect(urls[0].indexOf(`${WEB_URL}/_api/search/query?`)).toBe(0);
  });

  it('answers the sites in the response, mapped for the dropdown', async () => {
    const service: ISharePointSiteService = new SharePointSiteService(
      fakeClient(
        searchResponse([
          row('Marketing', 'https://contoso.sharepoint.com/sites/marketing'),
          row('Construction', 'https://contoso.sharepoint.com/sites/construction')
        ]),
        []
      ),
      WEB_URL
    );

    const sites: ISharePointSite[] = await service.getSites();

    expect(sites).toEqual([
      { title: 'Construction', url: 'https://contoso.sharepoint.com/sites/construction', siteId: '' },
      { title: 'Marketing', url: 'https://contoso.sharepoint.com/sites/marketing', siteId: '' }
    ]);
  });

  it('answers an empty list when the query matched nothing', async () => {
    const service: ISharePointSiteService = new SharePointSiteService(
      fakeClient(searchResponse([]), []),
      WEB_URL
    );

    await expect(service.getSites()).resolves.toEqual([]);
  });

  // A tenant that has disabled search, or a user whose token was rejected, has to reach the
  // component as a failure it can report - not as an empty dropdown that looks like a
  // tenant with no sites.
  it('lets a transport failure through rather than reporting no sites', async () => {
    const service: ISharePointSiteService = new SharePointSiteService(
      {
        getJson: async <T>(): Promise<T> => {
          throw new Error('403 Forbidden');
        },
        postJson: async <T>(): Promise<T> => {
          throw new Error('SharePointSiteService must not write.');
        }
      },
      WEB_URL
    );

    await expect(service.getSites()).rejects.toThrow('403 Forbidden');
  });
});
