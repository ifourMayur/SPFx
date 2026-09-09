/**
 * Tests for the SharePoint Search payload mapper.
 *
 * The search endpoint answers a column store - rows of `{ Key, Value }` cells - and the
 * envelope around it differs with the `Accept` header the transport happened to send
 * (`odata=nometadata` gives plain arrays, `odata=verbose` nests everything under
 * `d.query` and wraps every array in `results`). The dropdown wants neither of those, so
 * the flattening lives in one pure function and its quirks are pinned here rather than
 * discovered against a live tenant.
 */

import { ISharePointSite, toSharePointSites } from './SharePointSite';

/** One `Table.Rows` entry as the `odata=nometadata` response shapes it. */
function row(cells: { [key: string]: string }): unknown {
  return {
    Cells: Object.keys(cells).map((key: string) => ({
      Key: key,
      Value: cells[key],
      ValueType: 'Edm.String'
    }))
  };
}

/** A complete `odata=nometadata` search response carrying the given rows. */
function response(rows: unknown[]): unknown {
  return {
    PrimaryQueryResult: {
      RelevantResults: {
        Table: { Rows: rows },
        TotalRows: rows.length
      }
    }
  };
}

describe('toSharePointSites', () => {
  it('maps Title, Path and SiteId onto a site', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([
        row({
          Title: 'Contoso Projects',
          Path: 'https://contoso.sharepoint.com/sites/projects',
          SiteId: '2f6d1c3e-0000-4000-8000-000000000001'
        })
      ])
    );

    expect(sites).toEqual([
      {
        title: 'Contoso Projects',
        url: 'https://contoso.sharepoint.com/sites/projects',
        siteId: '2f6d1c3e-0000-4000-8000-000000000001'
      }
    ]);
  });

  it('reads the verbose envelope, where the rows hide under d.query and every array is wrapped in results', () => {
    const sites: ISharePointSite[] = toSharePointSites({
      d: {
        query: {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: {
                  results: [
                    {
                      Cells: {
                        results: [
                          { Key: 'Title', Value: 'Verbose site' },
                          { Key: 'Path', Value: 'https://contoso.sharepoint.com/sites/verbose' }
                        ]
                      }
                    }
                  ]
                }
              }
            }
          }
        }
      }
    });

    expect(sites).toEqual([
      { title: 'Verbose site', url: 'https://contoso.sharepoint.com/sites/verbose', siteId: '' }
    ]);
  });

  // A site collection created without a title still has to be selectable, and its URL is
  // the only thing left to name it by.
  it('names a site by its URL when Title is blank', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([row({ Title: '   ', Path: 'https://contoso.sharepoint.com/sites/untitled' })])
    );

    expect(sites[0].title).toBe('https://contoso.sharepoint.com/sites/untitled');
  });

  // `contentclass:STS_Site` matches every site collection the user can see, and in most
  // tenants that includes their own OneDrive - which is not a site anyone would pick here.
  it('drops OneDrive personal sites', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([
        row({ Title: 'Mayur Maheta', Path: 'https://contoso-my.sharepoint.com/personal/mayur_contoso_com' }),
        row({ Title: 'Projects', Path: 'https://contoso.sharepoint.com/sites/projects' })
      ])
    );

    expect(sites.map((site: ISharePointSite) => site.title)).toEqual(['Projects']);
  });

  it('skips a row with no Path, which nothing could navigate to', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([row({ Title: 'Pathless' }), row({ Title: 'Real', Path: 'https://contoso.sharepoint.com/sites/real' })])
    );

    expect(sites.map((site: ISharePointSite) => site.title)).toEqual(['Real']);
  });

  // `trimduplicates=false` is deliberate - it stops search from trimming rows it thinks are
  // near-identical - so the same site can come back twice, differing only in a trailing
  // slash or in case.
  it('keeps one entry per site, ignoring trailing slashes and case in the URL', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([
        row({ Title: 'Projects', Path: 'https://contoso.sharepoint.com/sites/projects' }),
        row({ Title: 'Projects again', Path: 'https://contoso.sharepoint.com/sites/Projects/' })
      ])
    );

    expect(sites).toEqual([
      { title: 'Projects', url: 'https://contoso.sharepoint.com/sites/projects', siteId: '' }
    ]);
  });

  it('strips a trailing slash from the URL it reports', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([row({ Title: 'Root', Path: 'https://contoso.sharepoint.com/' })])
    );

    expect(sites[0].url).toBe('https://contoso.sharepoint.com');
  });

  // Search orders by rank, which puts the sites in an order that means nothing to someone
  // reading a dropdown.
  it('sorts by title, ignoring case', () => {
    const sites: ISharePointSite[] = toSharePointSites(
      response([
        row({ Title: 'zebra', Path: 'https://contoso.sharepoint.com/sites/z' }),
        row({ Title: 'Alpha', Path: 'https://contoso.sharepoint.com/sites/a' }),
        row({ Title: 'beta', Path: 'https://contoso.sharepoint.com/sites/b' })
      ])
    );

    expect(sites.map((site: ISharePointSite) => site.title)).toEqual(['Alpha', 'beta', 'zebra']);
  });

  it('returns an empty list for undefined, so a failed load renders an empty dropdown', () => {
    expect(toSharePointSites(undefined)).toEqual([]);
  });

  it('returns an empty list for null', () => {
    expect(toSharePointSites(null)).toEqual([]);
  });

  it('returns an empty list for a payload without the search envelope', () => {
    expect(toSharePointSites({ unexpected: true })).toEqual([]);
  });

  it('returns an empty list when the query matched nothing', () => {
    expect(toSharePointSites(response([]))).toEqual([]);
  });
});
