// `BaseApiService` pulls in `@microsoft/sp-core-library` only for `Log`.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: { info: jest.fn(), warn: jest.fn(), error: jest.fn(), verbose: jest.fn() }
}));

import { IAuditSaveRequest } from '../models/EdControlsAudit';
import { IApiService } from './ApiService';
import { EdControlAuditService } from './EdControlAuditService';

interface ICall {
  method: string;
  url: string;
  query?: unknown;
  body?: unknown;
}

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
    },
    request: async (request: { method: string; relativeUrl: string; query?: unknown }) => {
      calls.push({ method: request.method, url: request.relativeUrl, query: request.query });
      return { ok: true, status: 200, statusText: 'OK', data: payload };
    }
  } as unknown as IApiService;
}

const REQUEST: IAuditSaveRequest = {
  buildingId: 9,
  auditName: 'A',
  edControlsDatabase: 'db',
  edControlsTemplateId: 't',
  edControlsTemplateName: 'T',
  supplierId: 7,
  dueDate: '2026-12-31',
  folderId: 10
};

describe('EdControlAuditService', () => {
  it('lists a project\'s audits', async () => {
    const calls: ICall[] = [];
    const service = new EdControlAuditService(fakeApi({ success: true, data: [{ id: 1, auditName: 'A' }] }, calls));

    expect((await service.getAudits(9))[0].auditName).toBe('A');
    expect(calls[0]).toEqual({ method: 'GET', url: 'EDControlAudit/GetByProject', query: { projectId: 9 } });
  });

  it('reads templates through the Web API and contacts by supplier', async () => {
    const calls: ICall[] = [];
    const service = new EdControlAuditService(fakeApi({ success: true, data: [] }, calls));

    await service.getTemplates('db-1');
    await service.getSupplierContacts(7);

    expect(calls).toEqual([
      { method: 'GET', url: 'EDControls/GetAuditTemplates', query: { ticketProjectId: 'db-1' } },
      { method: 'GET', url: 'Contact/GetAllContact', query: { issuerId: 7 } }
    ]);
  });

  it('creates one audit and rejects with the API\'s refusal', async () => {
    const calls: ICall[] = [];
    const ok = new EdControlAuditService(fakeApi({ success: true, message: 'Audit added successfully.' }, calls));
    const refused = new EdControlAuditService(fakeApi({ success: false, message: 'ED Controls refused the audit.' }, []));

    expect(await ok.createAudit(REQUEST)).toBe('Audit added successfully.');
    expect(calls[0]).toEqual({ method: 'POST', url: 'EDControlAudit/post', body: REQUEST });
    await expect(refused.createAudit(REQUEST)).rejects.toThrow('ED Controls refused the audit.');
  });

  it('keeps the per-row outcome even when every row failed', async () => {
    const calls: ICall[] = [];
    const service = new EdControlAuditService(fakeApi({
      success: false,
      message: '0 audit(s) added, 1 failed.',
      data: [{ index: 0, auditName: 'A', success: false, message: 'No template', id: 0 }]
    }, calls));

    const outcome = await service.createAudits([REQUEST]);

    expect(calls[0]).toEqual({ method: 'POST', url: 'EDControlAudit/postBulk', body: { audits: [REQUEST] } });
    expect(outcome.message).toBe('0 audit(s) added, 1 failed.');
    expect(outcome.results[0]).toMatchObject({ index: 0, success: false, message: 'No template' });
  });

  it('rejects a bulk refusal that carries no rows with its message', async () => {
    const service = new EdControlAuditService(fakeApi({ success: false, message: 'A maximum of 20 audits can be saved at once.', data: null }, []));

    await expect(service.createAudits([REQUEST])).rejects.toThrow('A maximum of 20 audits can be saved at once.');
  });

  it('removes through DELETE and reports a refusal', async () => {
    const calls: ICall[] = [];
    const ok = new EdControlAuditService(fakeApi({ success: true, data: { value: true } }, calls));
    const refused = new EdControlAuditService(fakeApi({ success: false, message: 'Opps! Something went wrong.', data: { value: false } }, []));

    await ok.deleteAudit(4);

    expect(calls[0]).toEqual({ method: 'DELETE', url: 'EDControlAudit/delete', query: { id: 4 } });
    await expect(refused.deleteAudit(4)).rejects.toThrow('Opps! Something went wrong.');
  });
});
