import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import { unwrapResponseDetail } from '../models/ApiEnvelope';
import { IProjectTemplateFolders, countTemplateFolders } from '../models/ProjectTemplateFolder';
import { IApiService } from './ApiService';

/**
 * Reads what a project template defines.
 *
 * Separate from `ILookupService`, which also touches the `ProjectTemplate` controller: that
 * one binds the template *dropdown* (`GetTemplateList`, a flat list of options), while this
 * reads the folder tree a template is made of.
 */
export interface IProjectTemplateService {
  /**
   * The template's folder tree - `GET ProjectTemplate/GetSubFolder?id=&isBIM=`.
   *
   * @param templateId - the project template's id.
   * @param isBim - the project's "Enable BIM folder" setting. The API filters a folder
   *   literally named `BIM` out of every template, then appends a synthetic `BIM` folder
   *   with a single `Model` child when this is `true`. Both carry id `0`, because neither
   *   is a row in the folder table.
   * @returns the tree, or `undefined` when the API answers without one.
   */
  getSubFolders(templateId: number, isBim: boolean): Promise<IProjectTemplateFolders | undefined>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'ProjectTemplateService';

/**
 * The action on the `ProjectTemplate` controller, held here while the controller route
 * itself lives in `src/config/environment.ts` - the same split `LookupService` and
 * `LoginService` use.
 */
const SUB_FOLDER_ACTION: string = 'GetSubFolder';

/** Default {@link IProjectTemplateService}, built on `IApiService`. */
export class ProjectTemplateService implements IProjectTemplateService {
  private readonly _apiService: IApiService;
  private readonly _endpoint: string;

  /**
   * @param apiService - HTTP gateway used for every call.
   * @param endpoint - controller route; defaults to `api.endpoints.projectTemplates` from
   *   the environment configuration, and can be overridden in tests.
   */
  public constructor(
    apiService: IApiService,
    endpoint: string = getEnvironment().api.endpoints.projectTemplates
  ) {
    this._apiService = apiService;
    this._endpoint = endpoint;
  }

  public async getSubFolders(
    templateId: number,
    isBim: boolean
  ): Promise<IProjectTemplateFolders | undefined> {
    const relativeUrl: string = `${this._endpoint}/${SUB_FOLDER_ACTION}`;
    const query = { id: templateId, isBIM: isBim };
    const url: string = this._apiService.resolveUrl(relativeUrl, query);

    const payload: unknown = await this._apiService.get<unknown>(relativeUrl, query);

    // Printed before the envelope is unwrapped, for the same reason `BuildingService` does
    // it in that order: `unwrapResponseDetail` throws on `success: false`, so a dump placed
    // after it would stay silent on exactly the responses worth reading. And because the
    // caller deliberately swallows this failure - the project is already saved by then -
    // this console line is the only place a rejected read is visible in full.
    this._logResponse(url, payload);

    const folders: IProjectTemplateFolders | undefined = unwrapResponseDetail<IProjectTemplateFolders>(
      payload,
      `the folders of template ${templateId}`
    );

    // An empty tree and a tree whose shape did not parse both render as "no folders"
    // downstream; the count is what tells them apart.
    Log.info(LOG_SOURCE, `Template ${templateId} defines ${countTemplateFolders(folders)} folder(s).`);

    return folders;
  }

  /**
   * Writes the complete response body to the browser console, exactly as the API sent it -
   * the `ResponseDetail` envelope included, since that is where a failure is reported.
   *
   * The dev-only dump `LoginService` and `BuildingService` also make; remove it alongside
   * them before shipping to production.
   */
  private _logResponse(url: string, payload: unknown): void {
    Log.info(LOG_SOURCE, `GET ${url} answered.`);
    console.log(`[${LOG_SOURCE}] GET ${url} response:`, payload);
  }
}
