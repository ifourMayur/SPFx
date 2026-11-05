import { getEnvironment } from '../config/environment';
import { ILookupOption } from '../models/Building';
import { IApiEndpoints } from '../models/Environment';
import { IPlanningGroup, ITicketGroup, toPlanningGroups, toTicketGroups } from '../models/IntegrationWork';
import { toLookupOptions } from '../models/Lookup';
import {
  IEdControlsSettings,
  IKypProjectsRequest,
  IKypSettings,
  toEdControlsSettings,
  toKypSettings,
  toTagList
} from '../models/ProjectIntegration';
import { IApiService } from './ApiService';
import { BaseApiService } from './BaseApiService';

/**
 * The project form's ED Controls and KYP lookups.
 *
 * The reference's MVC controller calls ED Controls and KYP straight from its own server
 * with credentials it reads from the Web API. A browser must never hold those, so the Web
 * API makes the calls instead (`api/EDControls`, `api/KYP`) and answers only lists and
 * "is it configured" - no credential or token comes back.
 *
 * | Method | Web API | Reference |
 * | --- | --- | --- |
 * | `getEdControlsSettings` | `GET EDControls/GetProjectSettings` | `Setting/GetEDControlsConfiguration` + the Razor model |
 * | `getEdControlsProjects` | `GET EDControls/GetProjects` | `Building/GetEdControlsProjects` |
 * | `getEdControlsTags` | `GET EDControls/GetTags` | `Building/GetTagList` |
 * | `getKypSettings` | `GET KYP/GetProjectSettings` | `Setting/GetKYPProjectConfiguration` + the Razor model |
 * | `getKypProjects` | `POST KYP/GetProjects` | `Building/GetPlanningProjects` |
 * | `getMyTickets` | `GET EDControls/GetMyTickets` | `Task/TicketPartial`, `Building/TicketPartial` |
 * | `getMyPlannings` | `GET KYP/GetMyPlannings` | `Task/GetPlanning`, `Building/PlanningPartial` |
 *
 * Every call is behind the API's SaaS gate (`EDCONTROL.CORE` / `KYP.CORE`); the form only
 * renders the sections through a `FeatureGate`, so a locked tenant never calls them.
 */
export interface IProjectIntegrationService {
  /** ED Controls configuration and the project's stored link and tag rows (`projectId` 0: configuration only). */
  getEdControlsSettings(projectId: number): Promise<IEdControlsSettings>;

  /** ED Controls projects a project can be linked to: `id` is the ED Controls "database". */
  getEdControlsProjects(): Promise<ILookupOption[]>;

  /** Audit tags of one ED Controls project. */
  getEdControlsTags(ticketProjectId: string): Promise<string[]>;

  /** KYP configuration and the project's stored planning link (`projectId` 0: configuration only). */
  getKypSettings(projectId: number): Promise<IKypSettings>;

  /** KYP projects the chosen token can see - see `IKypProjectsRequest`. */
  getKypProjects(request: IKypProjectsRequest): Promise<ILookupOption[]>;

  /** The user's open ED Controls tickets per project (`projectId` 0: every project they may see). */
  getMyTickets(projectId?: number): Promise<ITicketGroup[]>;

  /** The KYP planning items the user is responsible for, per project (`projectId` 0: every project they may see). */
  getMyPlannings(projectId?: number): Promise<IPlanningGroup[]>;
}

const LOG_SOURCE: string = 'ProjectIntegrationService';

const ACTIONS = {
  projectSettings: 'GetProjectSettings',
  projects: 'GetProjects',
  tags: 'GetTags',
  myTickets: 'GetMyTickets',
  myPlannings: 'GetMyPlannings'
};

/** Default {@link IProjectIntegrationService}: `ResponseDetail`-enveloped calls through `BaseApiService`. */
export class ProjectIntegrationService extends BaseApiService implements IProjectIntegrationService {
  private readonly _endpoints: IApiEndpoints;

  public constructor(apiService: IApiService, endpoints: IApiEndpoints = getEnvironment().api.endpoints) {
    super(apiService, LOG_SOURCE);
    this._endpoints = endpoints;
  }

  public async getEdControlsSettings(projectId: number): Promise<IEdControlsSettings> {
    const relativeUrl: string = this.route(this._endpoints.edControls, ACTIONS.projectSettings);
    const query = { projectId: Math.max(0, projectId) };
    const settings: IEdControlsSettings | undefined = toEdControlsSettings(
      await this.getData<unknown>(relativeUrl, query, { operation: 'the ED Controls settings' })
    );

    if (!settings) {
      throw this.refusal('The Web API answered the ED Controls settings in an unexpected format.', relativeUrl, query);
    }

    return settings;
  }

  public async getEdControlsProjects(): Promise<ILookupOption[]> {
    const relativeUrl: string = this.route(this._endpoints.edControls, ACTIONS.projects);

    return toLookupOptions(await this.getData<unknown>(relativeUrl, undefined, { operation: 'the ED Controls projects' }));
  }

  public async getEdControlsTags(ticketProjectId: string): Promise<string[]> {
    const relativeUrl: string = this.route(this._endpoints.edControls, ACTIONS.tags);

    return toTagList(
      await this.getData<unknown>(relativeUrl, { ticketProjectId }, { operation: 'the ED Controls tags' })
    );
  }

  public async getKypSettings(projectId: number): Promise<IKypSettings> {
    const relativeUrl: string = this.route(this._endpoints.kyp, ACTIONS.projectSettings);
    const query = { projectId: Math.max(0, projectId) };
    const settings: IKypSettings | undefined = toKypSettings(
      await this.getData<unknown>(relativeUrl, query, { operation: 'the KYP settings' })
    );

    if (!settings) {
      throw this.refusal('The Web API answered the KYP settings in an unexpected format.', relativeUrl, query);
    }

    return settings;
  }

  public async getMyTickets(projectId: number = 0): Promise<ITicketGroup[]> {
    const relativeUrl: string = this.route(this._endpoints.edControls, ACTIONS.myTickets);

    return toTicketGroups(
      await this.getData<unknown>(relativeUrl, { projectId: Math.max(0, projectId) }, { operation: 'your ED Controls tickets' })
    );
  }

  public async getMyPlannings(projectId: number = 0): Promise<IPlanningGroup[]> {
    const relativeUrl: string = this.route(this._endpoints.kyp, ACTIONS.myPlannings);

    return toPlanningGroups(
      await this.getData<unknown>(relativeUrl, { projectId: Math.max(0, projectId) }, { operation: 'your KYP planning' })
    );
  }

  public async getKypProjects(request: IKypProjectsRequest): Promise<ILookupOption[]> {
    const relativeUrl: string = this.route(this._endpoints.kyp, ACTIONS.projects);

    return toLookupOptions(
      await this.postData<unknown, IKypProjectsRequest>(relativeUrl, request, { operation: 'the KYP projects' })
    );
  }
}
