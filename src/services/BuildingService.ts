import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import { unwrapResponseDetail } from '../models/ApiEnvelope';
import { ApiError } from '../models/ApiError';
import { IBuildingForm } from '../models/Building';
import {
  IBuildingDetails,
  IBuildingSaveRequest,
  IBuildingSaveResponse,
  IPreservedProjectFields,
  IProjectSaveContext,
  IProjectSaveResult,
  toBuildingSaveRequest,
  toPreservedFields
} from '../models/BuildingSave';
import {
  IDocumentStorageRequest,
  IDocumentStorageResult,
  toDocumentStorageModel
} from '../models/DocumentStorage';
import { IProjectTemplateFolders, toTemplateFolders } from '../models/ProjectTemplateFolder';
import { IFolderProvisionResult } from '../models/SharePointFolder';
import { ISharePointSite } from '../models/SharePointSite';
import { IApiService } from './ApiService';
import { IDocumentStorageService } from './DocumentStorageService';
import { IProjectTemplateService } from './ProjectTemplateService';
import { ISharePointFolderService } from './SharePointFolderService';

/**
 * Persistence for the project add/edit screen.
 *
 * Two endpoints on the `ProjectController` that is aliased as `api/Building` - distinct
 * from `IProjectService`, which serves the REST-style `api/projects` the dashboard reads:
 *
 * Saving fans out to three calls whose order is a business rule, which is why it lives here
 * rather than in a component:
 *
 * | # | Call | Reference | When |
 * | --- | --- | --- | --- |
 * | 1 | `GET api/Building?id=` | step 7a, `DoActionForGet<BuildingModel>("?id=" + Id, "Building")` | update only |
 * | 2 | `POST api/Building` | step 8, `DoActionForPost<BuildingModel>(buildingModel, "Building")` | always |
 * | 3 | `GET ProjectTemplate/GetSubFolder` | step 5's template-only sibling | after a successful save |
 * | 4 | Create those folders in SharePoint | step 6, `SharePointHelper.CreateProjectFolders` | insert only, and only with a chosen site |
 * | 5 | `GET Building/GetProjectsubFolder` | step 9d.1 | insert only, once folders exist |
 * | 6 | `POST Document/AddDocumentStorageDetails` | step 9d.4 | insert only, once folders exist |
 *
 * The POST routes on `model.Id`, so the one call inserts (`id: 0`) and updates (`id > 0`).
 * The read-back before it exists because `ProjectController.Update` assigns the project's
 * image, its SharePoint folder id and the third-party integration fields from the request
 * **unconditionally** - see `IPreservedProjectFields`.
 *
 * ## Why the folder tree is read twice
 *
 * Steps 3 and 5 fetch the same shape from two different endpoints, and the difference is
 * not redundancy. Step 3 decides **what to create**: `ProjectTemplate/GetSubFolder` reports
 * the template as authored, and handles the `BIM` folder the API synthesizes. Step 5
 * decides **what to report**: `AddList` stamps the SharePoint ids onto `ProjectSubFolder`
 * rows matched by id, and only `Building/GetProjectsubFolder` knows those - the template
 * read answers with `SubFolder.ID` instead. Driving creation from the project-scoped read
 * as the reference does would create every active folder in the system rather than the
 * template's, so the two reads stay separate here. See `models/DocumentStorage.ts`.
 *
 * One step of the reference is still not implemented: the "rebind" second `POST
 * api/Building` that stores the tree's root id on the project itself, and the tender
 * folders. So a created project's own `sharepointFolderId` column stays empty even though
 * its folders and their ids are now recorded.
 */
export interface IBuildingService {
  /**
   * Reads a project back, as step 7a of the process document does.
   *
   * @returns the project, or `undefined` when the API answers without one.
   */
  getProject(id: number): Promise<IBuildingDetails | undefined>;

  /**
   * Reads a project's own folder tree - `GET Building/GetProjectsubFolder`.
   *
   * Step 1 of the reference's 9d. Unlike `IProjectTemplateService.getSubFolders`, the
   * sub-folder ids this reports are the project's `ProjectSubFolder` rows, which is what
   * `Document/AddDocumentStorageDetails` matches on.
   *
   * @param templateId - the template the project is on.
   * @param projectId - the project itself.
   * @returns the tree, or `undefined` when the API answers without one.
   */
  getProjectSubFolders(
    templateId: number,
    projectId: number
  ): Promise<IProjectTemplateFolders | undefined>;

  /**
   * Inserts or updates the project the form describes.
   *
   * An update reads the project back first ({@link getProject}) and echoes the fields it
   * must not lose. It is **refused** when that read comes back empty, rather than going
   * ahead with a POST that would blank the project's image and orphan its SharePoint
   * folders - the reference has no such guard because its Razor view round-trips those
   * fields through `@Html.HiddenFor`.
   *
   * Once the save succeeds, the template's folder tree is read through
   * `IProjectTemplateService` and, for a newly inserted project, created in the chosen
   * SharePoint site - and the ids SharePoint gave those folders are then reported back to
   * the Web API. None of that is allowed to fail the save: the project genuinely exists by
   * then, so reporting a failure would invite the user to save again and create a
   * duplicate. What could not be done comes back in the result instead - an absent
   * `templateFolders`, a `folderProvision` carrying failures, or a `documentStorage` saying
   * the ids were not recorded.
   *
   * @param form - the validated form; `form.id === 0` inserts.
   * @param context - session-derived values the form does not hold.
   * @returns the id the project now has, the response body, and the template's folders.
   * @throws ApiError - when either request fails, when a project cannot be read back
   *   before an update, when the API reports a failure inside an HTTP 200 (a `ModelState`
   *   rejection, or the subscription cap), or when it answers without a usable project id.
   *   Every one of those is surfaced to the user by `ProjectAddEdit`, which keeps them on
   *   the form with their input intact.
   */
  saveProject(form: IBuildingForm, context: IProjectSaveContext): Promise<IProjectSaveResult>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'BuildingService';

/**
 * The project-scoped folder-tree action on this controller. Note the API's own casing -
 * `subFolder` with a lower-case `s`, unlike `ProjectTemplate/GetSubFolder`.
 */
const PROJECT_SUB_FOLDER_ACTION: string = 'GetProjectsubFolder';

/** Default {@link IBuildingService}, built on `IApiService`. */
export class BuildingService implements IBuildingService {
  private readonly _apiService: IApiService;
  private readonly _projectTemplateService: IProjectTemplateService;
  private readonly _folderService: ISharePointFolderService;
  private readonly _documentStorageService: IDocumentStorageService;
  private readonly _endpoint: string;

  /**
   * @param apiService - HTTP gateway used for the calls on this controller. It supplies
   *   `baseUrl` and the `Authorization: Bearer` header carrying the application JWT that
   *   `LoginService` published on sign-in, so these `[Authorize]`d endpoints accept the
   *   requests.
   * @param projectTemplateService - reads the template's folder tree once a save succeeds.
   *   Injected as an interface rather than reached for, so the save can be tested without
   *   a second transport.
   * @param folderService - creates that tree in the chosen SharePoint site.
   * @param documentStorageService - reports the ids SharePoint gave those folders back to
   *   the Web API.
   * @param endpoint - controller route; defaults to `api.endpoints.building` from the
   *   environment configuration, and can be overridden in tests.
   */
  public constructor(
    apiService: IApiService,
    projectTemplateService: IProjectTemplateService,
    folderService: ISharePointFolderService,
    documentStorageService: IDocumentStorageService,
    endpoint: string = getEnvironment().api.endpoints.building
  ) {
    this._apiService = apiService;
    this._projectTemplateService = projectTemplateService;
    this._folderService = folderService;
    this._documentStorageService = documentStorageService;
    this._endpoint = endpoint;
  }

  public async getProject(id: number): Promise<IBuildingDetails | undefined> {
    const payload: unknown = await this._apiService.get<unknown>(this._endpoint, { id });

    return unwrapResponseDetail<IBuildingDetails>(payload, `project ${id}`);
  }

  public async getProjectSubFolders(
    templateId: number,
    projectId: number
  ): Promise<IProjectTemplateFolders | undefined> {
    const relativeUrl: string = `${this._endpoint}/${PROJECT_SUB_FOLDER_ACTION}`;
    // `isCreateProject` makes the API report the folders of a project that has just been
    // created, which is exactly the moment this is called from.
    const payload: unknown = await this._apiService.get<unknown>(relativeUrl, {
      id: templateId,
      ProjectId: projectId,
      isCreateProject: true
    });

    return unwrapResponseDetail<IProjectTemplateFolders>(
      payload,
      `the folders of project ${projectId}`
    );
  }

  public async saveProject(
    form: IBuildingForm,
    context: IProjectSaveContext
  ): Promise<IProjectSaveResult> {
    const isCreated: boolean = form.id <= 0;
    const details: IBuildingDetails | undefined = isCreated ? undefined : await this._readProject(form.id);
    const preserved: IPreservedProjectFields | undefined = details
      ? toPreservedFields(details)
      : undefined;
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form, context, preserved);
    const url: string = this._apiService.resolveUrl(this._endpoint);

    const payload: unknown = await this._apiService.post<unknown, IBuildingSaveRequest>(
      this._endpoint,
      request
    );

    // Printed before anything is read out of it, so a refused save is exactly as visible as
    // an accepted one. That ordering is the whole point: `unwrapResponseDetail` throws on
    // `success: false` - which is how the API reports a `ModelState` rejection, since it
    // answers HTTP 200 either way - so a dump placed after it would never run for the
    // responses most worth seeing.
    this._logResponse(url, isCreated, payload);

    const response: IBuildingSaveResponse | undefined = unwrapResponseDetail<IBuildingSaveResponse>(
      payload,
      'the saved project'
    );

    if (!response) {
      throw new ApiError('The Web API returned an empty response to the project save.', {
        status: 200,
        statusText: 'OK',
        url
      });
    }

    // The subscription cap arrives here rather than in the envelope, and stops the save:
    // the reference carries on and inserts orphan rows against project `0` instead.
    const message: string = (response.message || '').trim();
    if (message) {
      throw new ApiError(message, { status: 200, statusText: 'OK', url });
    }

    const projectId: number = Number(response.id);
    if (!isFinite(projectId) || projectId <= 0) {
      throw new ApiError('The Web API saved the project without returning its id.', {
        status: 200,
        statusText: 'OK',
        url
      });
    }

    // The template the project ended up on. The response is the first source of truth: an
    // insert that sent no template has the base one filled in server-side, so the request's
    // own value is not the resolved one. The other two only stand in when the API echoes
    // nothing back.
    const projectTemplateId: number =
      Number(response.projectTemplateId) ||
      Number(request.projectTemplateId) ||
      Number(details?.projectTemplateId) ||
      0;

    Log.info(
      LOG_SOURCE,
      `Project ${projectId} ${isCreated ? 'inserted' : 'updated'} on template ${projectTemplateId}.`
    );

    const templateFolders: IProjectTemplateFolders | undefined = await this._readTemplateFolders(
      projectTemplateId,
      request.isEnableBIMFolder
    );
    const folderProvision: IFolderProvisionResult | undefined = isCreated
      ? await this._createFolders(context.site, request.projectName, templateFolders)
      : undefined;

    return {
      projectId,
      isCreated,
      response,
      projectTemplateId: projectTemplateId || undefined,
      templateFolders,
      folderProvision,
      documentStorage: await this._recordFolderIds(
        projectId,
        projectTemplateId,
        request.projectName,
        folderProvision
      )
    };
  }

  /**
   * Tells the Web API which SharePoint folder each of the project's folders became.
   *
   * Steps 5 and 6 of the table above, run together because neither is worth doing without
   * the other: the read exists only to give the ids somewhere to be bound.
   *
   * Reports rather than throws, like the two steps before it. The project and its folders
   * both exist by the time this runs, so a failure here has to leave the save successful -
   * what it could not do comes back on the result instead.
   *
   * @returns the outcome, or `undefined` when there was nothing to record: no folders were
   *   provisioned (an update, or no chosen site), the root folder has no id, the
   *   project-scoped tree could not be read, or nothing in it matched a folder that exists.
   */
  private async _recordFolderIds(
    projectId: number,
    projectTemplateId: number,
    projectName: string,
    provision: IFolderProvisionResult | undefined
  ): Promise<IDocumentStorageResult | undefined> {
    if (!provision || !provision.rootId) {
      return undefined;
    }

    const projectFolders: IProjectTemplateFolders | undefined = await this._readProjectFolders(
      projectTemplateId,
      projectId
    );
    const model: IDocumentStorageRequest | undefined = toDocumentStorageModel(
      projectName,
      projectFolders,
      provision,
      projectId
    );

    if (!model) {
      Log.warn(
        LOG_SOURCE,
        `No folder of project ${projectId} could be matched to one created in SharePoint, so no ids were recorded.`
      );

      return undefined;
    }

    // Named from what was created rather than from the request, so a name SharePoint had
    // to sanitize is reported as the folder the user will actually find.
    const rootName: string = provision.folders.length > 0 ? provision.folders[0].name : projectName;

    try {
      await this._documentStorageService.addDocumentStorageDetails(model);

      return { isRecorded: true, rootName, message: '', model };
    } catch (error) {
      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(`The folder ids of project ${projectId} could not be recorded.`)
      );

      return {
        isRecorded: false,
        rootName,
        message:
          error instanceof Error
            ? error.message
            : 'The project folder ids could not be recorded on the Web API.',
        model
      };
    }
  }

  /**
   * Reads the project's own folder tree, reporting rather than throwing.
   *
   * Swallows its failure for the same reason {@link _readTemplateFolders} does: the project
   * and its SharePoint folders both exist by now, so a rejection here would show a failed
   * save for work that succeeded. The caller sees an absent tree and records no ids.
   */
  private async _readProjectFolders(
    projectTemplateId: number,
    projectId: number
  ): Promise<IProjectTemplateFolders | undefined> {
    if (projectTemplateId <= 0) {
      Log.warn(
        LOG_SOURCE,
        `Project ${projectId} reported no template, so its own folders were not read.`
      );

      return undefined;
    }

    try {
      return await this.getProjectSubFolders(projectTemplateId, projectId);
    } catch (error) {
      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(`The folders of project ${projectId} could not be read.`)
      );

      return undefined;
    }
  }

  /**
   * Creates the template's folders in the chosen site, reporting rather than throwing.
   *
   * **Inserts only**, which is what the reference does too: an update would create a second
   * root folder whenever the project had been renamed, leaving the original tree orphaned
   * and its documents where nobody would look for them. Repairing an existing project's
   * folders needs the stored `sharepointFolderId` to know which root it already has, and
   * that is not read back yet.
   *
   * Like the folder read before it, a failure here does not fail the save - the project
   * exists by now.
   */
  private async _createFolders(
    site: ISharePointSite | undefined,
    projectName: string,
    templateFolders: IProjectTemplateFolders | undefined
  ): Promise<IFolderProvisionResult | undefined> {
    if (!site) {
      Log.warn(LOG_SOURCE, 'No SharePoint site was chosen, so no project folders were created.');

      return undefined;
    }

    try {
      return await this._folderService.createProjectFolders(
        site,
        projectName,
        toTemplateFolders(templateFolders)
      );
    } catch (error) {
      // Reached when the site's document library could not be resolved at all, so nothing
      // was attempted. Individual folder failures come back inside the result instead.
      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(`The project folders could not be created in ${site.url}.`)
      );

      return undefined;
    }
  }

  /**
   * Reads back the project an update is about, refusing the save if it cannot.
   *
   * A project that will not read back is the one case where failing is safer than saving:
   * the POST would go ahead and set the project's image, storage type and SharePoint folder
   * id to null, and unlike a failed insert there is no risk of the user creating a
   * duplicate by trying again.
   */
  private async _readProject(id: number): Promise<IBuildingDetails> {
    const details: IBuildingDetails | undefined = await this.getProject(id);

    if (!details || !(Number(details.id) > 0)) {
      throw new ApiError(
        `Project ${id} could not be read back, so it was not saved. Reopen it and try again.`,
        { status: 200, statusText: 'OK', url: this._apiService.resolveUrl(this._endpoint, { id }) }
      );
    }

    return details;
  }

  /**
   * Reads the saved project's template folder tree, reporting rather than throwing.
   *
   * Deliberately swallows its failure: the project is already saved by the time this runs,
   * so rejecting here would show the user a failed save for a project that exists and
   * invite them to create a duplicate. The failure is logged instead, and the caller sees
   * an absent tree.
   */
  private async _readTemplateFolders(
    projectTemplateId: number,
    isBim: boolean
  ): Promise<IProjectTemplateFolders | undefined> {
    if (projectTemplateId <= 0) {
      Log.warn(LOG_SOURCE, 'The saved project reported no template, so its folders were not read.');

      return undefined;
    }

    try {
      return await this._projectTemplateService.getSubFolders(projectTemplateId, isBim);
    } catch (error) {
      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(`The folders of template ${projectTemplateId} could not be read.`)
      );

      return undefined;
    }
  }

  /**
   * Writes the complete response body to the browser console, exactly as the API sent it -
   * the `ResponseDetail` envelope included, not just the `data` inside it, because whether
   * the save was accepted is decided by `success` and `data.message` rather than by the
   * status code.
   *
   * This is how the save is verified during development: the same dev-only dump
   * `LoginService` makes for the sign-in handshake, and it should be removed alongside it
   * before shipping to production.
   */
  private _logResponse(url: string, isCreated: boolean, payload: unknown): void {
    const operation: string = isCreated ? 'insert' : 'update';

    Log.info(LOG_SOURCE, `POST ${url} (${operation}) answered.`);
    console.log(`[${LOG_SOURCE}] POST ${url} (${operation}) response:`, payload);
  }
}
