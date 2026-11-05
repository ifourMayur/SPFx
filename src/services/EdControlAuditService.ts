import { getEnvironment } from '../config/environment';
import { ApiError } from '../models/ApiError';
import { IResponseDetail, unwrapResponseDetail } from '../models/ApiEnvelope';
import { ILookupOption } from '../models/Building';
import {
  IAudit,
  IAuditSaveRequest,
  IAuditSaveResult,
  toAudits,
  toContactOptions,
  toSaveResults,
  toTemplateOptions
} from '../models/EdControlsAudit';
import { IApiEndpoints } from '../models/Environment';
import { IApiResponse } from '../models/ApiResponse';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/**
 * The project Audits page, ported from the reference's `_AddEdControlAudit.cshtml`.
 *
 * | Method | Web API | Reference (MVC proxy) |
 * | --- | --- | --- |
 * | `getAudits` | `GET EDControlAudit/GetByProject` | `Building/GetEdControlAudits` |
 * | `getTemplates` | `GET EDControls/GetAuditTemplates` | `Building/GetEdControlsAuditTemplates` (called ED Controls itself) |
 * | `getSupplierContacts` | `GET Contact/GetAllContact` | `Building/GetSupplierContactList` |
 * | `createAudit` | `POST EDControlAudit/post` | `Building/AddEdControlAudit` |
 * | `createAudits` | `POST EDControlAudit/postBulk` | `Building/AddEdControlAuditBulk` |
 * | `deleteAudit` | `DELETE EDControlAudit/delete` | `Building/DeleteEdControlAudit` |
 *
 * Creating an audit creates it in ED Controls and a placeholder document; removing one deletes
 * that document (ED Controls itself keeps the audit). Every route is behind the API's
 * `EDCONTROL.CORE` gate.
 */
export interface IEdControlAuditService {
  getAudits(projectId: number): Promise<IAudit[]>;

  /** Audit templates of the project's ED Controls project. */
  getTemplates(ticketProjectId: string): Promise<ILookupOption[]>;

  /** Contacts of one supplier; the chosen contact's e-mail becomes the audit's responsible person. */
  getSupplierContacts(supplierId: number): Promise<ILookupOption[]>;

  /** Creates one audit. Rejects with the API's message when it was refused. */
  createAudit(request: IAuditSaveRequest): Promise<string>;

  /**
   * Creates the duplicate rows. Rows are created independently, so this resolves with the
   * outcome per row (and the API's summary) even when some or all of them failed.
   */
  createAudits(requests: IAuditSaveRequest[]): Promise<{ message: string; results: IAuditSaveResult[] }>;

  /** Removes an audit and permanently deletes its document from the project. */
  deleteAudit(id: number): Promise<void>;
}

const LOG_SOURCE: string = 'EdControlAuditService';

const ACTIONS = {
  byProject: 'GetByProject',
  post: 'post',
  postBulk: 'postBulk',
  delete: 'delete',
  templates: 'GetAuditTemplates',
  contacts: 'GetAllContact'
};

/** Default {@link IEdControlAuditService}. */
export class EdControlAuditService extends BaseApiService implements IEdControlAuditService {
  private readonly _endpoints: IApiEndpoints;

  public constructor(apiService: IApiService, endpoints: IApiEndpoints = getEnvironment().api.endpoints) {
    super(apiService, LOG_SOURCE);
    this._endpoints = endpoints;
  }

  public async getAudits(projectId: number): Promise<IAudit[]> {
    const relativeUrl: string = this.route(this._endpoints.edControlAudit, ACTIONS.byProject);

    return toAudits(await this.getData<unknown>(relativeUrl, { projectId }, { operation: 'the audits' }));
  }

  public async getTemplates(ticketProjectId: string): Promise<ILookupOption[]> {
    const relativeUrl: string = this.route(this._endpoints.edControls, ACTIONS.templates);

    return toTemplateOptions(
      await this.getData<unknown>(relativeUrl, { ticketProjectId }, { operation: 'the ED Controls templates' })
    );
  }

  public async getSupplierContacts(supplierId: number): Promise<ILookupOption[]> {
    const relativeUrl: string = this.route(this._endpoints.contact, ACTIONS.contacts);

    return toContactOptions(await this.getData<unknown>(relativeUrl, { issuerId: supplierId }, { operation: 'the contacts' }));
  }

  public async createAudit(request: IAuditSaveRequest): Promise<string> {
    const relativeUrl: string = this.route(this._endpoints.edControlAudit, ACTIONS.post);
    const payload: unknown = await this.postRaw<IAuditSaveRequest>(relativeUrl, request, { operation: 'the audit' });

    unwrapResponseDetail<unknown>(payload, 'the audit');

    return EdControlAuditService._message(payload);
  }

  public async createAudits(requests: IAuditSaveRequest[]): Promise<{ message: string; results: IAuditSaveResult[] }> {
    const relativeUrl: string = this.route(this._endpoints.edControlAudit, ACTIONS.postBulk);
    // Read without unwrapping: the envelope says `success: false` when every row failed, and
    // the per-row outcome is exactly what the page needs then.
    const payload: unknown = await this.postRaw<{ audits: IAuditSaveRequest[] }>(
      relativeUrl,
      { audits: requests },
      { operation: 'the audits' }
    );
    const results: IAuditSaveResult[] = toSaveResults(
      typeof payload === 'object' && payload !== null ? (payload as IResponseDetail<unknown>).data : undefined
    );

    if (results.length === 0) {
      // A refusal of the whole batch (nothing supplied, too many rows) carries no rows.
      unwrapResponseDetail<unknown>(payload, 'the audits');
      throw this.refusal('The Web API did not report how the audits were saved.', relativeUrl);
    }

    return { message: EdControlAuditService._message(payload), results };
  }

  public async deleteAudit(id: number): Promise<void> {
    const relativeUrl: string = this.route(this._endpoints.edControlAudit, ACTIONS.delete);
    // `IApiService.delete` discards the body, and the envelope's `success` is the answer here.
    const response: IApiResponse<unknown> = await this.apiService.request<unknown>({
      method: 'DELETE',
      relativeUrl,
      query: { id }
    });
    const result: unknown = unwrapResponseDetail<unknown>(response.data, 'the removal of the audit');

    if (typeof result === 'object' && result !== null && (result as { value?: unknown }).value === false) {
      throw new ApiError('The audit could not be removed.', { status: 200 });
    }
  }

  private static _message(payload: unknown): string {
    const message: unknown = typeof payload === 'object' && payload !== null ? (payload as IResponseDetail<unknown>).message : '';

    return typeof message === 'string' ? message : '';
  }
}
