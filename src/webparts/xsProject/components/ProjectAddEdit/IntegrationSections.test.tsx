/**
 * Render tests for the project form's Audits (ED Controls) and Planning (KYP) sections.
 *
 * Each section is mounted inside a small host that holds the form value the way
 * `ProjectAddEdit` does, so `onLoaded` / `onChange` round-trip through props exactly as in
 * the form.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import EdControlsSection from './EdControlsSection';
import KypPlanningSection from './KypPlanningSection';
import { ILookupOption } from '../../../../models/Building';
import { ApiError } from '../../../../models/ApiError';
import {
  IEdControlsForm,
  IEdControlsSettings,
  IKypForm,
  IKypProjectsRequest,
  IKypSettings,
  ITagRow,
  createTagRow
} from '../../../../models/ProjectIntegration';
import { ILookupService } from '../../../../services/LookupService';
import { IProjectIntegrationService } from '../../../../services/ProjectIntegrationService';

const FOLDERS: ILookupOption[] = [{ id: '10', name: '01 Administratie' }];

function lookupService(): ILookupService {
  return {
    getFolderLevel: async (projectId: number, level: string): Promise<ILookupOption[]> =>
      level === 'folderId' ? FOLDERS : []
  } as unknown as ILookupService;
}

interface IIntegrationCalls {
  kypRequests: IKypProjectsRequest[];
  tagRequests: string[];
  edProjectRequests: number;
}

function integrationService(
  calls: IIntegrationCalls,
  ed: Partial<IEdControlsSettings> = {},
  kyp: Partial<IKypSettings> = {},
  failure?: Error
): IProjectIntegrationService {
  return {
    getEdControlsSettings: async (): Promise<IEdControlsSettings> => {
      if (failure) {
        throw failure;
      }
      return { isConfigured: true, isTicketSource: false, ticketProjectId: '', tagRows: [], ...ed };
    },
    getEdControlsProjects: async (): Promise<ILookupOption[]> => {
      calls.edProjectRequests += 1;
      return [{ id: 'db-1', name: 'ED project 1' }];
    },
    getEdControlsTags: async (ticketProjectId: string): Promise<string[]> => {
      calls.tagRequests.push(ticketProjectId);
      return ['Oplevering', 'Keuring'];
    },
    getKypSettings: async (): Promise<IKypSettings> => ({
      isConfigured: true,
      isPlanningSource: false,
      planningProjectId: '',
      hasProjectToken: false,
      ...kyp
    }),
    getKypProjects: async (request: IKypProjectsRequest): Promise<ILookupOption[]> => {
      calls.kypRequests.push(request);
      return [{ id: 'k-1', name: 'KYP project 1' }];
    },
    getMyTickets: async () => [],
    getMyPlannings: async () => []
  };
}

function newCalls(): IIntegrationCalls {
  return { kypRequests: [], tagRequests: [], edProjectRequests: 0 };
}

/** Holds the section's value like the form does, and remembers the last one. */
class EdHost extends React.Component<
  { projectId: number; service: IProjectIntegrationService; onValue: (value: IEdControlsForm) => void },
  { value?: IEdControlsForm }
> {
  public state: { value?: IEdControlsForm } = {};

  public render(): React.ReactElement {
    return (
      <EdControlsSection
        projectId={this.props.projectId}
        integrationService={this.props.service}
        lookupService={lookupService()}
        suppliers={[{ id: '7', name: 'Van Dijk' }]}
        isGrandChildFolderEnabled={false}
        value={this.state.value}
        isDisabled={false}
        onLoaded={this._set}
        onChange={this._set}
      />
    );
  }

  private _set = (value: IEdControlsForm): void => {
    this.setState({ value });
    this.props.onValue(value);
  };
}

class KypHost extends React.Component<
  { projectId: number; service: IProjectIntegrationService; onValue: (value: IKypForm) => void },
  { value?: IKypForm }
> {
  public state: { value?: IKypForm } = {};

  public render(): React.ReactElement {
    return (
      <KypPlanningSection
        projectId={this.props.projectId}
        integrationService={this.props.service}
        value={this.state.value}
        isDisabled={false}
        onLoaded={this._set}
        onChange={this._set}
      />
    );
  }

  private _set = (value: IKypForm): void => {
    this.setState({ value });
    this.props.onValue(value);
  };
}

describe('Integration sections', () => {
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

  async function clickText(text: string): Promise<void> {
    const buttons: HTMLButtonElement[] = Array.prototype.slice.call(host.querySelectorAll('button'));
    const target: HTMLButtonElement = buttons.filter(
      (button: HTMLButtonElement) => (button.textContent || '').trim() === text
    )[0];
    await act(async () => {
      target.click();
    });
  }

  async function clickCheckbox(label: string): Promise<void> {
    const boxes: HTMLInputElement[] = Array.prototype.slice.call(host.querySelectorAll('input[type="checkbox"]'));
    const target: HTMLInputElement = boxes.filter(
      (input: HTMLInputElement) => (input.closest('label')?.textContent || '').trim() === label
    )[0];
    await act(async () => {
      target.click();
    });
  }

  async function choose(selector: string, value: string): Promise<void> {
    const select: HTMLSelectElement = host.querySelector(selector) as HTMLSelectElement;
    await act(async () => {
      select.value = value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  describe('EdControlsSection', () => {
    it('says ED Controls is not set up and reports no value, so the save keeps what is stored', async () => {
      const values: IEdControlsForm[] = [];
      await render(<EdHost projectId={7} service={integrationService(newCalls(), { isConfigured: false })} onValue={(v) => values.push(v)} />);

      expect(host.textContent).toContain('ED Controls is not set up for your organisation.');
      expect(values).toHaveLength(0);
    });

    it('shows the stored link and tag rows, with the project, its tags and the row folders loaded', async () => {
      const calls: IIntegrationCalls = newCalls();
      const stored: ITagRow = { ...createTagRow(), tags: ['Keuring'], folder: { folderId: '10', subFolderId: '', subSubFolderId: '', subSubSubFolderId: '', supplierId: '7' } };
      await render(
        <EdHost
          projectId={7}
          service={integrationService(calls, { isTicketSource: true, ticketProjectId: 'db-1', tagRows: [stored] })}
          onValue={() => undefined}
        />
      );

      expect((host.querySelector('#project-edControls-ticketProjectId') as HTMLSelectElement).value).toBe('db-1');
      expect(host.textContent).toContain('ED project 1');
      expect(calls.tagRequests).toEqual(['db-1']);
      expect(host.querySelector('[aria-label="Remove Keuring"]')).not.toBeNull();
      expect(host.textContent).toContain('01 Administratie');
    });

    it('loads the ED Controls projects when switched on while adding, without tag rows', async () => {
      const calls: IIntegrationCalls = newCalls();
      const values: IEdControlsForm[] = [];
      await render(<EdHost projectId={0} service={integrationService(calls)} onValue={(v) => values.push(v)} />);

      await clickCheckbox('ED controls');

      expect(calls.edProjectRequests).toBe(1);
      expect(values[values.length - 1]).toEqual({ isTicketSource: true, ticketProjectId: '', tagRows: [] });
      expect(host.textContent).toContain('Tag rows can be added once the project has been saved.');
    });

    it('reloads the tags when another ED Controls project is chosen', async () => {
      const calls: IIntegrationCalls = newCalls();
      await render(<EdHost projectId={7} service={integrationService(calls, { isTicketSource: true })} onValue={() => undefined} />);

      await choose('#project-edControls-ticketProjectId', 'db-1');

      expect(calls.tagRequests).toEqual(['db-1']);
    });

    it('adds a tag row only once every row has tags and a folder', async () => {
      const values: IEdControlsForm[] = [];
      await render(
        <EdHost projectId={7} service={integrationService(newCalls(), { isTicketSource: true, ticketProjectId: 'db-1' })} onValue={(v) => values.push(v)} />
      );

      await clickText('+ Add tag row');
      expect(host.textContent).toContain('Please select at least one tag');
      expect(host.textContent).toContain('Please select a folder');

      const latest: IEdControlsForm = values[values.length - 1];
      expect(latest.tagRows).toHaveLength(1);
    });

    it('reports a failed load and retries', async () => {
      await render(
        <EdHost projectId={7} service={integrationService(newCalls(), {}, {}, new ApiError('ED Controls could not be reached.', { status: 500 }))} onValue={() => undefined} />
      );

      expect(host.textContent).toContain('ED Controls could not be reached.');
      expect(host.textContent).toContain('Try again');
    });
  });

  describe('KypPlanningSection', () => {
    it('says KYP is not set up when neither the account nor the project has a token', async () => {
      const values: IKypForm[] = [];
      await render(<KypHost projectId={7} service={integrationService(newCalls(), {}, { isConfigured: false })} onValue={(v) => values.push(v)} />);

      expect(host.textContent).toContain('KYP is not set up for your organisation.');
      expect(values).toHaveLength(0);
    });

    it('lists the projects of the account token when switched on', async () => {
      const calls: IIntegrationCalls = newCalls();
      await render(<KypHost projectId={7} service={integrationService(calls)} onValue={() => undefined} />);

      await clickCheckbox('KYP Project');

      expect(calls.kypRequests).toEqual([{ projectId: 7, useProjectToken: false, authToken: '' }]);
      expect(host.textContent).toContain('KYP project 1');
    });

    it('authorizes a typed alternative token through the Web API', async () => {
      const calls: IIntegrationCalls = newCalls();
      const values: IKypForm[] = [];
      await render(
        <KypHost projectId={7} service={integrationService(calls, {}, { isPlanningSource: true })} onValue={(v) => values.push(v)} />
      );

      await clickCheckbox('Use alternative token');
      const input: HTMLInputElement = host.querySelector('#project-kyp-token') as HTMLInputElement;
      await act(async () => {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
        setter?.call(input, 'typed-token');
        input.dispatchEvent(new Event('input', { bubbles: true }));
      });
      await clickText('Authorize');

      expect(calls.kypRequests[calls.kypRequests.length - 1]).toEqual({ projectId: 7, useProjectToken: true, authToken: 'typed-token' });
      expect(values[values.length - 1].projectToken).toBe('typed-token');
    });

    it('keeps a stored alternative token hidden, and says leaving the field empty keeps it', async () => {
      await render(
        <KypHost projectId={7} service={integrationService(newCalls(), {}, { isPlanningSource: true, hasProjectToken: true })} onValue={() => undefined} />
      );

      const input: HTMLInputElement = host.querySelector('#project-kyp-token') as HTMLInputElement;
      expect(input.value).toBe('');
      expect(input.placeholder).toContain('leave empty to keep it');
    });
  });
});
