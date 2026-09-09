import { Log } from '@microsoft/sp-core-library';

import { IProjectTemplateFolder } from '../models/ProjectTemplateFolder';
import {
  IFolderProvisionFailure,
  IFolderProvisionResult,
  planFolderTree
} from '../models/SharePointFolder';
import { ISharePointSite, withoutTrailingSlash } from '../models/SharePointSite';
import { ISpHttpClient } from './SharePointHttpClient';

/**
 * Creates a project's folder tree in a SharePoint site's default document library.
 *
 * The SPFx-native replacement for the reference application's
 * `SharePointHelper.CreateProjectFolders`, which used app-only Graph credentials against a
 * tenant-configured site. Here the site is the one the user chose in `SharePointSites`, and
 * every request runs as the signed-in user through `ISpHttpClient` - so this needs no app
 * registration, no client secret and no tenant-admin API approval, and a user can only
 * create folders where they already have permission to.
 */
export interface ISharePointFolderService {
  /**
   * Creates the tree the template defines, under a folder named after the project.
   *
   * Never rejects for a folder it could not create: the caller runs this after the project
   * has already been saved, so a rejection would report a failed save for a project that
   * exists. Individual failures come back in {@link IFolderProvisionResult.failed}. It does
   * reject when the library itself cannot be resolved, because then nothing was attempted.
   *
   * @param site - the site chosen by the user.
   * @param projectName - becomes the root folder name.
   * @param folders - `lstProjectFolderDetailsViewModel` from `ProjectTemplate/GetSubFolder`.
   */
  createProjectFolders(
    site: ISharePointSite,
    projectName: string,
    folders: IProjectTemplateFolder[] | undefined
  ): Promise<IFolderProvisionResult>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'SharePointFolderService';

/** Shape of `_api/web/defaultDocumentLibrary/rootFolder`; SharePoint REST stays PascalCase. */
interface IRootFolderResponse {
  ServerRelativeUrl?: string;
}

/**
 * Escapes a server-relative path for use inside an OData string literal in a URL.
 *
 * Two separate encodings are at work. A single quote inside an OData literal is escaped by
 * doubling it, and the result then has to survive as part of a query string - so spaces,
 * ampersands and percent signs are percent-encoded. The path separators are put back
 * afterwards: `DecodedUrl` expects a real path, and `%2F` would read as a folder name
 * containing a slash rather than as a level of nesting.
 */
function toPathLiteral(serverRelativePath: string): string {
  return encodeURIComponent(serverRelativePath.replace(/'/g, "''")).replace(/%2F/gi, '/');
}

/** The `https://host` part of a site URL, or `''` when there is not one. */
function toSiteOrigin(siteUrl: string): string {
  const schemeEnd: number = siteUrl.indexOf('://');

  if (schemeEnd < 0) {
    return '';
  }

  const hostEnd: number = siteUrl.indexOf('/', schemeEnd + 3);

  return hostEnd < 0 ? siteUrl : siteUrl.substring(0, hostEnd);
}

/** A server-relative path turned into a URL a browser can open, segment by segment. */
function toFolderUrl(origin: string, serverRelativePath: string): string {
  return (
    origin +
    serverRelativePath
      .split('/')
      .map((segment: string): string => encodeURIComponent(segment))
      .join('/')
  );
}

/** Default {@link ISharePointFolderService}, built on `ISpHttpClient`. */
export class SharePointFolderService implements ISharePointFolderService {
  private readonly _client: ISpHttpClient;

  public constructor(client: ISpHttpClient) {
    this._client = client;
  }

  public async createProjectFolders(
    site: ISharePointSite,
    projectName: string,
    folders: IProjectTemplateFolder[] | undefined
  ): Promise<IFolderProvisionResult> {
    const siteUrl: string = withoutTrailingSlash((site.url || '').trim());
    const paths: string[] = planFolderTree(projectName, folders);
    const libraryRoot: string = await this._resolveLibraryRoot(siteUrl);
    const origin: string = toSiteOrigin(siteUrl);

    const result: IFolderProvisionResult = {
      rootUrl: paths.length > 0 ? toFolderUrl(origin, `${libraryRoot}/${paths[0]}`) : '',
      created: [],
      skipped: [],
      failed: []
    };

    // Sequential on purpose: `planFolderTree` orders parents before children, and
    // SharePoint cannot create `A/B` until `A` exists. Running these in parallel would
    // fail every nested folder.
    for (let index: number = 0; index < paths.length; index++) {
      const path: string = paths[index];
      const parentFailure: string = SharePointFolderService._findParentFailure(result.failed, path);

      if (parentFailure) {
        // Recorded rather than silently dropped, so the counts add up to what was planned.
        result.failed.push({ path, message: `Its parent folder could not be created: ${parentFailure}` });
        continue;
      }

      try {
        const wasCreated: boolean = await this._createFolder(siteUrl, `${libraryRoot}/${path}`);
        (wasCreated ? result.created : result.skipped).push(path);
      } catch (error) {
        result.failed.push({
          path,
          message: error instanceof Error ? error.message : 'The folder could not be created.'
        });
      }
    }

    Log.info(
      LOG_SOURCE,
      `${site.url}: ${result.created.length} folder(s) created, ${result.skipped.length} already present, ${result.failed.length} failed.`
    );

    return result;
  }

  /**
   * The server-relative URL of the site's default document library root.
   *
   * Asked for rather than assumed: the library is titled "Documents" in the UI but lives at
   * `Shared Documents` in most sites, at `Documents` in others, and under a translated name
   * in a non-English tenant. `defaultDocumentLibrary` resolves all of those.
   */
  private async _resolveLibraryRoot(siteUrl: string): Promise<string> {
    const response: IRootFolderResponse = await this._client.getJson<IRootFolderResponse>(
      `${siteUrl}/_api/web/defaultDocumentLibrary/rootFolder?$select=ServerRelativeUrl`
    );
    const serverRelativeUrl: string = (response?.ServerRelativeUrl || '').trim();

    if (!serverRelativeUrl) {
      throw new Error(`${siteUrl} did not report a default document library to create folders in.`);
    }

    return withoutTrailingSlash(serverRelativeUrl);
  }

  /**
   * Creates one folder, reporting whether it had to.
   *
   * Attempt-then-verify rather than check-then-create: asking first would cost a request
   * per folder and still race another user, so the create is attempted and a failure is
   * only believed once the folder turns out not to be there. That is what makes an existing
   * folder a skip instead of an error - and it means a partly-failed run can be retried by
   * saving again.
   *
   * @returns `true` when this call created the folder, `false` when it already existed.
   */
  private async _createFolder(siteUrl: string, serverRelativePath: string): Promise<boolean> {
    try {
      await this._client.postJson<unknown>(
        `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl='${toPathLiteral(serverRelativePath)}')`
      );

      return true;
    } catch (error) {
      if (await this._exists(siteUrl, serverRelativePath)) {
        return false;
      }

      throw error;
    }
  }

  /** Whether a folder is there, used only to interpret a failed create. */
  private async _exists(siteUrl: string, serverRelativePath: string): Promise<boolean> {
    try {
      await this._client.getJson<unknown>(
        `${siteUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${toPathLiteral(serverRelativePath)}')?$select=Exists`
      );

      return true;
    } catch {
      return false;
    }
  }

  /** The message of the nearest failed ancestor of `path`, or `''` when it has none. */
  private static _findParentFailure(failed: IFolderProvisionFailure[], path: string): string {
    const ancestor: IFolderProvisionFailure | undefined = failed.filter(
      (failure: IFolderProvisionFailure): boolean => path.indexOf(`${failure.path}/`) === 0
    )[0];

    return ancestor ? ancestor.message : '';
  }
}
