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
 * The part of an `SP.Folder` this service reads.
 *
 * `UniqueId` is asked for by name on the existence check but not on the create: it is one
 * of `SP.Folder`'s default properties, so `AddUsingPath` answers with it already, and
 * adding a `$select` to a method call is a way to break the create for no gain.
 */
interface IFolderResponse {
  UniqueId?: string;
}

/**
 * The part of a `driveItem` this service reads, from the site's own drive API.
 *
 * Graph-shaped rather than SharePoint-shaped, so `id` here is camelCase where `UniqueId`
 * above is not - the two responses come from two different services.
 */
interface IDriveItemResponse {
  id?: string;
}

/** What creating (or finding) one folder produced. */
interface ICreatedFolder {
  /** `true` when this call created it, `false` when it was already there. */
  wasCreated: boolean;
  /** SharePoint's `UniqueId`, or `''` when it answered without one. */
  uniqueId: string;
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

/** A path with each segment percent-encoded and the separators left alone. */
function toEncodedPath(path: string): string {
  return path
    .split('/')
    .map((segment: string): string => encodeURIComponent(segment))
    .join('/');
}

/** A server-relative path turned into a URL a browser can open, segment by segment. */
function toFolderUrl(origin: string, serverRelativePath: string): string {
  return origin + toEncodedPath(serverRelativePath);
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
      rootId: '',
      folders: [],
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
        const created: ICreatedFolder = await this._createFolder(siteUrl, `${libraryRoot}/${path}`);
        (created.wasCreated ? result.created : result.skipped).push(path);

        // Asked for once the folder is known to exist, whether this run made it or found
        // it: a retry after a partial failure has to be able to report the ids of folders
        // an earlier run created, and those are skips the second time round.
        const driveItemId: string = await this._readDriveItemId(siteUrl, path);

        result.folders.push({
          name: path.split('/').slice(-1)[0],
          path,
          id: driveItemId,
          uniqueId: created.uniqueId,
          url: toFolderUrl(origin, `${libraryRoot}/${path}`),
          wasCreated: created.wasCreated
        });

        // `planFolderTree` always plans the project's own root folder first, so the first
        // path is the one the Web API knows the project's folder tree by.
        if (index === 0) {
          result.rootId = driveItemId;
        }
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
   * @returns whether this call created the folder, and the `UniqueId` SharePoint holds it
   *   under - which is **not** the id Graph addresses it by; see {@link _readDriveItemId}.
   */
  private async _createFolder(siteUrl: string, serverRelativePath: string): Promise<ICreatedFolder> {
    try {
      const folder: IFolderResponse = await this._client.postJson<IFolderResponse>(
        `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl='${toPathLiteral(serverRelativePath)}')`
      );

      return { wasCreated: true, uniqueId: (folder?.UniqueId || '').trim() };
    } catch (error) {
      const existing: IFolderResponse | undefined = await this._readFolder(siteUrl, serverRelativePath);

      if (existing) {
        return { wasCreated: false, uniqueId: (existing.UniqueId || '').trim() };
      }

      throw error;
    }
  }

  /**
   * Reads a folder, used to interpret a failed create.
   *
   * Its answer carries the folder's id as well as its existence, which is what lets a
   * folder that was already there still be reported to the Web API - the case a retry after
   * a partial failure runs into, where the ids that matter belong to folders this run did
   * not create.
   *
   * @returns the folder, or `undefined` when it is not there.
   */
  private async _readFolder(
    siteUrl: string,
    serverRelativePath: string
  ): Promise<IFolderResponse | undefined> {
    try {
      return await this._client.getJson<IFolderResponse>(
        `${siteUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${toPathLiteral(serverRelativePath)}')?$select=Exists,UniqueId`
      );
    } catch {
      return undefined;
    }
  }

  /**
   * The `driveItem.id` Graph addresses a folder by, read from the site's own drive API.
   *
   * `AddUsingPath` answers with SharePoint's `UniqueId`, and that GUID is useless to the
   * Web API: it reaches these folders through Graph with app-only credentials, which
   * addresses items by an opaque drive-scoped id (`01VL4HET...`). The two are separate
   * identifiers rather than two encodings of one value, so this has to be looked up.
   *
   * `_api/v2.0` is the same drive API Graph serves, hosted by the site itself, so
   * `ISpHttpClient` can call it as the signed-in user - no app registration, no Graph token
   * and no tenant-admin consent, which is the property that makes every other call in this
   * service deployable to a new tenant unchanged. `_api/v2.0/drive` is the site's default
   * document library, the same one {@link _resolveLibraryRoot} resolves, so the planned
   * paths address it untranslated.
   *
   * @returns the id, or `''` when it could not be read. Deliberately not an error: the
   *   folder itself exists by the time this runs, so failing here would report a folder
   *   that is genuinely there as one that is not. An empty id instead leaves the folder out
   *   of what is reported to the Web API.
   */
  private async _readDriveItemId(siteUrl: string, path: string): Promise<string> {
    try {
      const item: IDriveItemResponse = await this._client.getJson<IDriveItemResponse>(
        `${siteUrl}/_api/v2.0/drive/root:/${toEncodedPath(path)}`
      );
      const id: string = (item?.id || '').trim();

      if (!id) {
        Log.warn(LOG_SOURCE, `${path} was created, but the drive API reported no id for it.`);
      }

      return id;
    } catch (error) {
      Log.warn(
        LOG_SOURCE,
        `${path} was created, but its drive API id could not be read, so it will not be reported: ${
          error instanceof Error ? error.message : 'the request failed'
        }`
      );

      return '';
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
