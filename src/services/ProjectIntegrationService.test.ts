// `BaseApiService` pulls in `@microsoft/sp-core-library` only for `Log`.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: { info: jest.fn(), warn: jest.fn(), error: jest.fn(), verbose: jest.fn() }
}));

import { ApiError } from '../models/ApiError';
import { IApiService } from './ApiService';
import { ProjectIntegrationService } from './ProjectIntegrationService';

interface ICall {
  method: 'GET' | 'POST';
  url: string;
  query?: unknown;
  body?: unknown;
}

/** An `IApiService` answering every call with `payload`, recording what was asked. */
function fakeApi(payload: unknown, calls: ICall[]): IApiService {
  return {
    baseUrl: 'https://api.example.com/api',
    resolveUrl: (relativeUrl: string) => `https://api.example.com/api/${relativeUrl}`,
    get: async (url: string, query?: unknown) => {
      calls.push({ method: 'GET', url, query });
      return payload;
    },
    post: async (url: string, body: unknown) => {
      calls.push({ method: 'POST', url, body });
      return payload;
    }
  } as unknown as IApiService;
}

function envelope(data: unknown, success: boolean = true, message: string = ''): unknown {
  return { data, success, message };
}

describe('ProjectIntegrationService', () => {
  it('reads the ED Controls settings of a project', async () => {
    const calls: ICall[] = [];
    const service = new ProjectIntegrationService(
      fakeApi(envelope({ isConfigured: true, isTicketSource: true, ticketProjectId: 'db', tagRows: [] }), calls)
    );

    const settings = await service.getEdControlsSettings(7);

    expect(calls).toEqual([{ method: 'GET', url: 'EDControls/GetProjectSettings', query: { projectId: 7 } }]);
    expect(settings.ticketProjectId).toBe('db');
  });

  it('maps the ED Controls projects, keyed by database', async () => {
    const calls: ICall[] = [];
    const service = new ProjectIntegrationService(fakeApi(envelope([{ id: 0, strId: 'db-1', name: 'One' }]), calls));

    expect(await service.getEdControlsProjects()).toEqual([{ id: 'db-1', name: 'One' }]);
    expect(calls[0].url).toBe('EDControls/GetProjects');
  });

  it('asks the tags of one ED Controls project', async () => {
    const calls: ICall[] = [];
    const service = new ProjectIntegrationService(fakeApi(envelope(['a', 'b']), calls));

    expect(await service.getEdControlsTags('db-1')).toEqual(['a', 'b']);
    expect(calls[0]).toEqual({ method: 'GET', url: 'EDControls/GetTags', query: { ticketProjectId: 'db-1' } });
  });

  it('posts the KYP token choice in the body, never in the URL', async () => {
    const calls: ICall[] = [];
    const service = new ProjectIntegrationService(fakeApi(envelope([{ strId: 'k-1', name: 'K' }]), calls));

    await service.getKypProjects({ projectId: 7, useProjectToken: true, authToken: 'secret' });

    expect(calls[0]).toEqual({
      method: 'POST',
      url: 'KYP/GetProjects',
      body: { projectId: 7, useProjectToken: true, authToken: 'secret' }
    });
  });

  it("reads the user's tickets and planning for every project, or one", async () => {
    const calls: ICall[] = [];
    const service = new ProjectIntegrationService(fakeApi(envelope([{ projectId: 3, projectName: 'P', tickets: [], items: [] }]), calls));

    expect((await service.getMyTickets())[0].projectName).toBe('P');
    await service.getMyPlannings(3);

    expect(calls).toEqual([
      { method: 'GET', url: 'EDControls/GetMyTickets', query: { projectId: 0 } },
      { method: 'GET', url: 'KYP/GetMyPlannings', query: { projectId: 3 } }
    ]);
  });

  it("surfaces the API's own refusal message", async () => {
    const service = new ProjectIntegrationService(fakeApi(envelope(null, false, 'KYP did not accept the token.'), []));

    await expect(service.getKypProjects({ projectId: 0, useProjectToken: true, authToken: 'x' })).rejects.toThrow(
      'KYP did not accept the token.'
    );
  });

  it('refuses a settings answer in an unexpected format', async () => {
    const service = new ProjectIntegrationService(fakeApi(envelope('nope'), []));

    await expect(service.getKypSettings(1)).rejects.toBeInstanceOf(ApiError);
  });
});
