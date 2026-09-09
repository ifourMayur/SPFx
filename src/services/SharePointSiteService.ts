import {
  ISharePointSite,
  SITE_SELECT_PROPERTIES,
  toSharePointSites,
  withoutTrailingSlash
} from '../models/SharePointSite';
import { ISpHttpClient } from './SharePointHttpClient';

/**
 * KQL that matches site collections and nothing else. `STS_Site` is the content class of a
 * site collection; `STS_Web` would add every subsite, and no content class at all would
 * return documents.
 */
const SITE_CONTENT_CLASS_QUERY: string = 'contentclass:STS_Site';

/**
 * Rows asked for in one request, which is also the most search returns per query.
 *
 * A tenant with more site collections than this would need `startrow` paging; no tenant
 * this solution is deployed to is near the limit, so the dropdown reads one page.
 */
const MAX_SITES: number = 500;

/** The SharePoint sites a user may choose from. */
export interface ISharePointSiteService {
  /**
   * Every site collection the signed-in user can open, ordered by title.
   *
   * Rejects when the query fails, so a caller can tell a tenant with no sites from a
   * request that never succeeded.
   */
  getSites(): Promise<ISharePointSite[]>;
}

/**
 * Lists sites through the SharePoint Search API.
 *
 * Search is used rather than Microsoft Graph's `/sites` because it needs no app
 * registration and no tenant-admin API approval: the query runs as the signed-in user
 * through SPFx's own `SPHttpClient`, and its results are security-trimmed to the sites
 * that user may open. The cost is search's own indexing latency - a site created minutes
 * ago may not be listed yet.
 */
export class SharePointSiteService implements ISharePointSiteService {
  private readonly _client: ISpHttpClient;
  private readonly _webAbsoluteUrl: string;

  public constructor(client: ISpHttpClient, webAbsoluteUrl: string) {
    this._client = client;
    this._webAbsoluteUrl = webAbsoluteUrl;
  }

  public async getSites(): Promise<ISharePointSite[]> {
    const response: unknown = await this._client.getJson<unknown>(this._buildQueryUrl());

    return toSharePointSites(response);
  }

  /**
   * The search query, addressed to the site hosting the web part.
   *
   * Any site's `_api/search/query` searches the whole tenant, so the current web is simply
   * the nearest endpoint to ask. The KQL and the property list are quoted and then encoded
   * because search expects the quotes as part of the parameter value.
   */
  private _buildQueryUrl(): string {
    const site: string = withoutTrailingSlash(this._webAbsoluteUrl.trim());

    const parameters: string[] = [
      `querytext=${encodeURIComponent(`'${SITE_CONTENT_CLASS_QUERY}'`)}`,
      `selectproperties=${encodeURIComponent(`'${SITE_SELECT_PROPERTIES.join(',')}'`)}`,
      `rowlimit=${MAX_SITES}`,
      // Search's own de-duplication drops whole sites it judges near-identical, which would
      // hide real choices from the dropdown; `toSharePointSites` de-duplicates by URL
      // instead.
      'trimduplicates=false'
    ];

    return `${site}/_api/search/query?${parameters.join('&')}`;
  }
}
