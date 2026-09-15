import { Log } from '@microsoft/sp-core-library';

import { getEnvironment } from '../config/environment';
import { MESSAGE_TYPE, unwrapResponseDetail } from '../models/ApiEnvelope';
import { ApiError } from '../models/ApiError';
import { IDocumentStorageRequest } from '../models/DocumentStorage';
import { IApiService } from './ApiService';

/**
 * Records which SharePoint folder each of a project's folders ended up as.
 *
 * Step 4 of the reference application's SharePoint branch - step 9d of
 * `D:\Projects\TaskXS\BMDeskV2\docs\BuildingController-AddEdit-POST-Process.md`:
 * `DoActionForPost<GetProjectTemplateDetailsViewModel>(responseDetail.Result.Data,
 * "Document/AddDocumentStorageDetails")`.
 *
 * Server-side this writes a `DocumentStorageDetails` row per top-level folder and stamps
 * the SharePoint id onto the project's existing `ProjectSubFolder` rows, which is what lets
 * the rest of the application find a project's documents in SharePoint. Until it runs, the
 * folders exist in the site but the Web API does not know they do.
 *
 * The reference's sibling call in the same branch - the second `POST api/Building` that
 * stores the tree's root id on the project itself - runs just before this one, so the two
 * together leave the project and every one of its folders addressable. See
 * `BuildingService._rebindFolderId`.
 */
export interface IDocumentStorageService {
  /**
   * Posts the bound folder tree, as `toDocumentStorageModel` built it.
   *
   * @param model - the project-scoped tree with a SharePoint id on every node. Its
   *   `projectId` is what every row is written against, so a model without one records
   *   nothing - which is why `toDocumentStorageModel` refuses to build one.
   * @throws ApiError - when the API refuses the call, **including the refusals it reports
   *   inside an HTTP 200**. The caller is expected to catch: by the time this runs the
   *   project and its folders both exist, so a rejection here must not read as a failed
   *   save.
   */
  addDocumentStorageDetails(model: IDocumentStorageRequest): Promise<void>;
}

/** Source name used for SPFx log entries emitted by this service. */
const LOG_SOURCE: string = 'DocumentStorageService';

/**
 * The action on the `Document` controller, held here while the controller route itself
 * lives in `src/config/environment.ts` - the same split `LookupService`,
 * `ProjectTemplateService` and `LoginService` use.
 */
const ADD_ACTION: string = 'AddDocumentStorageDetails';

/** Default {@link IDocumentStorageService}, built on `IApiService`. */
export class DocumentStorageService implements IDocumentStorageService {
  private readonly _apiService: IApiService;
  private readonly _endpoint: string;

  /**
   * @param apiService - HTTP gateway used for the call. It supplies `baseUrl` and the
   *   `Authorization: Bearer` header carrying the application JWT.
   * @param endpoint - controller route; defaults to `api.endpoints.document` from the
   *   environment configuration, and can be overridden in tests.
   */
  public constructor(
    apiService: IApiService,
    endpoint: string = getEnvironment().api.endpoints.document
  ) {
    this._apiService = apiService;
    this._endpoint = endpoint;
  }

  public async addDocumentStorageDetails(model: IDocumentStorageRequest): Promise<void> {
    const relativeUrl: string = `${this._endpoint}/${ADD_ACTION}`;
    const url: string = this._apiService.resolveUrl(relativeUrl);

    const payload: unknown = await this._apiService.post<unknown, IDocumentStorageRequest>(
      relativeUrl,
      model
    );

    // Printed before anything is read out of it, for the same reason `BuildingService` does
    // it in that order: the checks below throw, and the caller swallows what they throw, so
    // a dump placed after them would stay silent on exactly the responses worth reading.
    this._logResponse(url, model, payload);

    // Throws on `success: false` - a request the API rejected outright.
    unwrapResponseDetail<IDocumentStorageRequest>(payload, "the project's folder ids");

    const faulted: string = DocumentStorageService._toFaultMessage(payload);
    if (faulted) {
      throw new ApiError(faulted, { status: 200, statusText: 'OK', url });
    }

    Log.info(
      LOG_SOURCE,
      `Recorded the SharePoint folder ids of project ${model.projectId} against ${
        (model.lstProjectFolderDetailsViewModel || []).length
      } folder(s).`
    );
  }

  /**
   * The message of a failure the API reported while still claiming success.
   *
   * `DocumentController.AddDocumentStorageDetails` catches its own exception and answers
   * `GetDataWithMessage(..., "Something went wrong, please try after sometime", **true**,
   * DropMessageType.Error)` - so `success` is `true` on the one response that means nothing
   * at all was written, and `messageType` is the only field that still says so.
   *
   * @returns the message, or `''` when the response does not report a fault.
   */
  private static _toFaultMessage(payload: unknown): string {
    if (typeof payload !== 'object' || payload === null) {
      return '';
    }

    const envelope: { message?: string; messageType?: number } = payload as {
      message?: string;
      messageType?: number;
    };

    if (envelope.messageType !== MESSAGE_TYPE.error) {
      return '';
    }

    return (
      (envelope.message || '').trim() ||
      'The Web API could not record the project folder ids, without saying why.'
    );
  }

  /**
   * Writes the request and the complete response body to the browser console, exactly as
   * the API sent it - the `ResponseDetail` envelope included, because whether the folders
   * were recorded is decided by `success` and `messageType` rather than by the status code.
   *
   * The same dev-only dump `BuildingService` and `LoginService` make, and it should be
   * removed alongside them before shipping to production.
   */
  private _logResponse(url: string, model: IDocumentStorageRequest, payload: unknown): void {
    Log.info(LOG_SOURCE, `POST ${url} answered.`);
    console.log(`[${LOG_SOURCE}] POST ${url} request:`, model);
    console.log(`[${LOG_SOURCE}] POST ${url} response:`, payload);
  }
}
