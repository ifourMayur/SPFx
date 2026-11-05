/**
 * Render tests for the Audits page: the list, adding (validation, then the create call),
 * duplicating (count, drafts, the bulk save keeping failed rows) and removing (with a
 * confirmation), plus the permission refusal.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import Audits from './Audits';
import AuditManager from './AuditManager';
import { ILookupOption } from '../../../../models/Building';
import { IAudit, IAuditSaveRequest, IAuditSaveResult } from '../../../../models/EdControlsAudit';
import { IEdControlsSettings } from '../../../../models/ProjectIntegration';
import { IEdControlAuditService } from '../../../../services/EdControlAuditService';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';
import { ISaasService } from '../../../../services/SaasService';

const AUDIT: IAudit = {
  id: 4,
  buildingId: 9,
  auditName: 'Fire safety',
  edControlsDatabase: 'db-1',
  edControlsTemplateId: 't-1',
  edControlsTemplateName: 'Template 1',
  supplierId: 7,
  supplierName: 'Van Dijk',
  contactName: '',
  dueDate: '2026-12-31T00:00:00',
  folderId: 10,
  folderPath: 'Docs',
  status: 2
};

interface IAuditCalls {
  created: IAuditSaveRequest[];
  bulk: IAuditSaveRequest[][];
  deleted: number[];
}

function auditService(calls: IAuditCalls, bulkResults?: (requests: IAuditSaveRequest[]) => IAuditSaveResult[]): IEdControlAuditService {
  return {
    getAudits: async (): Promise<IAudit[]> => [AUDIT, { ...AUDIT, id: 5, auditName: 'Open one', status: 1 }],
    getTemplates: async (): Promise<ILookupOption[]> => [{ id: 't-1', name: 'Template 1' }],
    getSupplierContacts: async (): Promise<ILookupOption[]> => [{ id: '3', name: 'Anna' }],
    createAudit: async (request: IAuditSaveRequest): Promise<string> => {
      calls.created.push(request);
      return 'Audit added successfully.';
    },
    createAudits: async (requests: IAuditSaveRequest[]) => {
      calls.bulk.push(requests);
      const results: IAuditSaveResult[] = bulkResults
        ? bulkResults(requests)
        : requests.map((request: IAuditSaveRequest, index: number) => ({ index, auditName: request.auditName, success: true, message: '', id: 10 + index }));
      return { message: `${results.filter((r) => r.success).length} audit(s) added.`, results };
    },
    deleteAudit: async (id: number): Promise<void> => {
      calls.deleted.push(id);
    }
  };
}

function integrationService(settings: Partial<IEdControlsSettings> = {}): IProjectIntegrationService {
  return {
    getEdControlsSettings: async (): Promise<IEdControlsSettings> => ({
      isConfigured: true,
      projectName: 'Damrak 1',
      isTicketSource: true,
      ticketProjectId: 'db-1',
      tagRows: [],
      ...settings
    })
  } as unknown as IProjectIntegrationService;
}

function lookupService(): ILookupService {
  return {
    getSuppliers: async (): Promise<ILookupOption[]> => [{ id: '7', name: 'Van Dijk' }],
    getFolderLevel: async (projectId: number, level: string): Promise<ILookupOption[]> =>
      level === 'folderId' ? [{ id: '10', name: 'Docs' }] : []
  } as unknown as ILookupService;
}

function newCalls(): IAuditCalls {
  return { created: [], bulk: [], deleted: [] };
}

describe('Audits', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  async function render(element: React.ReactElement): Promise<void> {
    await act(async () => {
      ReactDOM.render(element, host);
    });
  }

  async function renderManager(calls: IAuditCalls, settings?: Partial<IEdControlsSettings>, bulk?: (r: IAuditSaveRequest[]) => IAuditSaveResult[]): Promise<void> {
    await render(
      <AuditManager
        projectId={9}
        integrationService={integrationService(settings)}
        auditService={auditService(calls, bulk)}
        lookupService={lookupService()}
        isGrandChildFolderEnabled={false}
      />
    );
  }

  async function clickText(text: string): Promise<void> {
    const buttons: HTMLButtonElement[] = Array.prototype.slice.call(host.querySelectorAll('button'));
    const target: HTMLButtonElement = buttons.filter((button: HTMLButtonElement) => (button.textContent || '').trim() === text)[0];
    await act(async () => {
      target.click();
    });
  }

  async function type(selector: string, value: string): Promise<void> {
    const input: HTMLInputElement = host.querySelector(selector) as HTMLInputElement;
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      setter?.call(input, value);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }

  async function choose(selector: string, value: string): Promise<void> {
    const select: HTMLSelectElement = host.querySelector(selector) as HTMLSelectElement;
    await act(async () => {
      select.value = value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  async function submit(): Promise<void> {
    const form: HTMLFormElement = host.querySelector('form') as HTMLFormElement;
    await act(async () => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });
  }

  it('lists the audits with their status', async () => {
    await renderManager(newCalls());

    expect(host.textContent).toContain('Fire safety');
    expect(host.textContent).toContain('31/12/2026');
    expect(host.textContent).toContain('Completed');
    expect(host.textContent).toContain('Waiting');
  });

  it('disables Add Audits and says why when the project has no saved ED Controls project', async () => {
    await renderManager(newCalls(), { isTicketSource: false, ticketProjectId: '' });

    const add: HTMLButtonElement = Array.prototype.slice.call(host.querySelectorAll('button'))
      .filter((button: HTMLButtonElement) => button.textContent === 'Add Audits')[0];
    expect(add.disabled).toBe(true);
    expect(host.textContent).toContain('save the project before adding audits');
  });

  it('validates the add form, then creates the audit in the saved ED Controls project', async () => {
    const calls: IAuditCalls = newCalls();
    await renderManager(calls);

    await clickText('Add Audits');
    await submit();
    expect(host.textContent).toContain('Audit name is required.');
    expect(host.textContent).toContain('Folder is required.');
    expect(calls.created).toHaveLength(0);

    await type('#audit-name', 'Roof check');
    await choose('#audit-templateId', 't-1');
    await choose('#audit-supplierId', '7');
    await type('#audit-dueDate', '15/01/2027');
    await choose('#audit-folderId', '10');
    await submit();

    expect(calls.created).toEqual([expect.objectContaining({
      buildingId: 9,
      auditName: 'Roof check',
      edControlsDatabase: 'db-1',
      edControlsTemplateId: 't-1',
      edControlsTemplateName: 'Template 1',
      supplierId: 7,
      dueDate: '2027-01-15',
      folderId: 10
    })]);
    expect(host.textContent).toContain('Audit added successfully.');
  });

  it('removes an audit only after the confirmation', async () => {
    const calls: IAuditCalls = newCalls();
    await renderManager(calls);

    const remove: HTMLButtonElement = Array.prototype.slice.call(host.querySelectorAll('button'))
      .filter((button: HTMLButtonElement) => button.textContent === 'Delete')[0];
    await act(async () => {
      remove.click();
    });
    expect(host.textContent).toContain('permanently deleted');
    expect(calls.deleted).toHaveLength(0);

    await clickText('Remove');
    expect(calls.deleted).toEqual([4]);
  });

  it('duplicates into numbered drafts and keeps a row that failed, with its reason', async () => {
    const calls: IAuditCalls = newCalls();
    await renderManager(calls, {}, (requests: IAuditSaveRequest[]) => requests.map((request, index) => ({
      index,
      auditName: request.auditName,
      success: index === 0,
      message: index === 0 ? '' : 'ED Controls refused it.',
      id: index === 0 ? 20 : 0
    })));

    const duplicate: HTMLButtonElement = Array.prototype.slice.call(host.querySelectorAll('button'))
      .filter((button: HTMLButtonElement) => button.textContent === 'Duplicate')[0];
    await act(async () => {
      duplicate.click();
    });
    await type('#audit-duplicateCount', '21');
    await submit();
    expect(host.textContent).toContain('Enter a whole number between 1 and 20.');

    await type('#audit-duplicateCount', '2');
    await submit();
    const names: string[] = Array.prototype.slice.call(host.querySelectorAll('input[aria-label="Audit name"]'))
      .map((input: HTMLInputElement) => input.value);
    expect(names).toEqual(['Fire safety - 1', 'Fire safety - 2']);

    await clickText('Save');

    expect(calls.bulk[0].map((request) => request.auditName)).toEqual(['Fire safety - 1', 'Fire safety - 2']);
    expect(calls.bulk[0][0]).toMatchObject({ edControlsDatabase: 'db-1', supplierId: 7, dueDate: '2026-12-31', folderId: 10 });
    const left: string[] = Array.prototype.slice.call(host.querySelectorAll('input[aria-label="Audit name"]'))
      .map((input: HTMLInputElement) => input.value);
    expect(left).toEqual(['Fire safety - 2']);
    expect(host.textContent).toContain('ED Controls refused it.');
  });

  it('refuses the page to a role that may not manage audits', async () => {
    await render(
      <Audits
        projectId={9}
        integrationService={integrationService()}
        auditService={auditService(newCalls())}
        lookupService={lookupService()}
        saasService={{} as ISaasService}
        isGrandChildFolderEnabled={false}
        canManage={false}
        onBack={() => undefined}
      />
    );

    expect(host.textContent).toContain('does not have permission to manage audits');
  });
});
