/**
 * Render tests for the project form.
 *
 * These cover what a successful TypeScript build cannot: that the form actually mounts,
 * that its dropdowns are filled from the lookup service rather than from anything
 * hard-coded, that validation messages reach the DOM, that the excluded integration
 * sections really are absent, and that the permission refusal and the load-failure state
 * render instead of the form.
 *
 * `react-dom/test-utils` and the jsdom environment are already available, so this needs
 * no new dependency. Every `ReactDOM.render` here is unmounted in `afterEach`.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import { HashRouter } from 'react-router-dom';

import ProjectAddEditRoute from './ProjectAddEditRoute';
import { IBuildingForm, ILookupOption } from '../../../../models/Building';
import { ILookupService } from '../../../../services/LookupService';

/** Options the stub service hands back, so assertions can look for real names. */
const STUB = {
  projectTemplates: [{ id: '1', name: 'Base template' }],
  users: [{ id: '101', name: 'Jan Jansen' }],
  spatialBreakdowns: [{ id: '1', name: 'Ground floor (GF)' }],
  tenderTemplates: [{ id: '1', name: 'Standard tender' }],
  suppliers: [{ id: '201', name: 'Van Dijk Installaties' }],
  folders: [{ id: '10', name: '01 Administratie' }]
};

/** A lookup service that resolves immediately with {@link STUB}. */
function stubLookupService(overrides?: Partial<ILookupService>): ILookupService {
  return {
    getProjectTemplates: async (): Promise<ILookupOption[]> => STUB.projectTemplates,
    getUsers: async (): Promise<ILookupOption[]> => STUB.users,
    getSpatialBreakdowns: async (): Promise<ILookupOption[]> => STUB.spatialBreakdowns,
    getTenderTemplates: async (): Promise<ILookupOption[]> => STUB.tenderTemplates,
    getSuppliers: async (): Promise<ILookupOption[]> => STUB.suppliers,
    getProjectFolders: async (): Promise<ILookupOption[]> => STUB.folders,
    getProjectSubFolders: async (): Promise<ILookupOption[]> => [],
    getProjectSubSubFolders: async (): Promise<ILookupOption[]> => [],
    getProjectSubSubSubFolders: async (): Promise<ILookupOption[]> => [],
    ...overrides
  };
}

describe('ProjectAddEdit', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  /** Renders the add form and waits for its dropdown load to settle. */
  async function renderAddForm(
    onSave: (form: IBuildingForm) => Promise<void>,
    lookupService: ILookupService = stubLookupService(),
    canSave: boolean = true
  ): Promise<void> {
    await act(async () => {
      ReactDOM.render(
        <HashRouter>
          <ProjectAddEditRoute
            lookupService={lookupService}
            mode="add"
            rows={[]}
            savedForms={{}}
            features={{ domainGroup: 'taskXs', isGrandChildFolderEnabled: true }}
            clientId="42"
            clientName="Mayur"
            canSave={canSave}
            onSave={onSave}
          />
        </HashRouter>,
        host
      );
    });
  }

  /** Sets a text input the way React's value tracker notices. */
  function setValue(selector: string, value: string): void {
    const input: HTMLInputElement = host.querySelector(selector) as HTMLInputElement;

    act(() => {
      // React 17 tracks the value setter, so assign through the prototype descriptor.
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      setter?.call(input, value);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }

  /** Clicks the first button whose visible text matches exactly. */
  async function clickText(text: string): Promise<void> {
    const buttons: HTMLButtonElement[] = Array.prototype.slice.call(host.querySelectorAll('button'));
    const target: HTMLButtonElement = buttons.filter(
      (button: HTMLButtonElement) => (button.textContent || '').trim() === text
    )[0];

    await act(async () => {
      target.click();
    });
  }

  it('mounts the add form with its core fields', async () => {
    await renderAddForm(async () => undefined);

    expect(host.querySelector('#project-buildingName')).not.toBeNull();
    expect(host.querySelector('#project-postcode')).not.toBeNull();
    expect(host.querySelector('#project-address')).not.toBeNull();
    // Project Template must be enabled in add mode.
    expect((host.querySelector('#project-projectTemplateId') as HTMLSelectElement).disabled).toBe(false);
    // Reset is offered in add mode only.
    expect(host.textContent).toContain('Reset');
  });

  it('fills the single-select dropdowns from the lookup service', async () => {
    await renderAddForm(async () => undefined);

    // A native <select> renders its options whether or not it is focused.
    expect(host.textContent).toContain('Base template');
  });

  it('fills the multi-selects from the lookup service', async () => {
    await renderAddForm(async () => undefined);

    // A closed multi-select renders no options, so it has to be opened first.
    await act(async () => {
      (host.querySelector('#project-userIds') as HTMLButtonElement).click();
    });
    expect(host.textContent).toContain('Jan Jansen');

    await act(async () => {
      (host.querySelector('#project-spatialBreakdownIds') as HTMLButtonElement).click();
    });
    expect(host.textContent).toContain('Ground floor (GF)');
  });

  it('lets several users be selected and shows each as a chip', async () => {
    const saved: IBuildingForm[] = [];
    const twoUsers: ILookupService = stubLookupService({
      getUsers: async (): Promise<ILookupOption[]> => [
        { id: '101', name: 'Jan Jansen' },
        { id: '102', name: 'Sanne de Vries' }
      ]
    });

    await renderAddForm(async (form: IBuildingForm): Promise<void> => {
      saved.push(form);
    }, twoUsers);

    await act(async () => {
      (host.querySelector('#project-userIds') as HTMLButtonElement).click();
    });

    const boxes: HTMLInputElement[] = Array.prototype.slice.call(
      host.querySelectorAll('#project-userIds ~ div input[type="checkbox"]')
    );
    await act(async () => {
      boxes[0].click();
    });
    await act(async () => {
      boxes[1].click();
    });

    // Both selections survive, which a native <select multiple> loses on a plain click,
    // and each shows as its own chip.
    const chips: Element[] = Array.prototype.slice.call(
      host.querySelectorAll('[aria-label^="Remove "]')
    );
    expect(chips.map((chip: Element) => chip.getAttribute('aria-label'))).toEqual([
      'Remove Jan Jansen',
      'Remove Sanne de Vries'
    ]);

    setValue('#project-buildingName', 'Two users');
    setValue('#project-address', 'Damrak 2');
    setValue('#project-postcode', '1012 LP');
    const form: HTMLFormElement = host.querySelector('form') as HTMLFormElement;
    await act(async () => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });

    expect(saved[0].userIds).toEqual(['101', '102']);
  });

  it('removes a selected user through its chip', async () => {
    await renderAddForm(async () => undefined);

    await act(async () => {
      (host.querySelector('#project-userIds') as HTMLButtonElement).click();
    });
    await act(async () => {
      (host.querySelector('#project-userIds ~ div input[type="checkbox"]') as HTMLInputElement).click();
    });
    expect(host.querySelector('[aria-label="Remove Jan Jansen"]')).not.toBeNull();

    await act(async () => {
      (host.querySelector('[aria-label="Remove Jan Jansen"]') as HTMLElement).click();
    });

    expect(host.querySelector('[aria-label="Remove Jan Jansen"]')).toBeNull();
  });

  // An empty dropdown and a broken one look identical, so the empty case says so.
  it('says so when the API returns no options for a dropdown', async () => {
    const empty: ILookupService = stubLookupService({
      getUsers: async (): Promise<ILookupOption[]> => [],
      getProjectTemplates: async (): Promise<ILookupOption[]> => []
    });

    await renderAddForm(async () => undefined, empty);

    expect(host.textContent).toContain('The Web API returned no users for your account.');
    expect(host.textContent).toContain('The Web API returned no project templates for your account.');
  });

  it('requests every dropdown the reference binds, and none of the excluded ones', async () => {
    const called: string[] = [];
    const recording: ILookupService = stubLookupService({
      getProjectTemplates: async (): Promise<ILookupOption[]> => {
        called.push('projectTemplates');
        return STUB.projectTemplates;
      },
      getUsers: async (): Promise<ILookupOption[]> => {
        called.push('users');
        return STUB.users;
      },
      getSpatialBreakdowns: async (): Promise<ILookupOption[]> => {
        called.push('spatialBreakdowns');
        return STUB.spatialBreakdowns;
      },
      getTenderTemplates: async (): Promise<ILookupOption[]> => {
        called.push('tenderTemplates');
        return STUB.tenderTemplates;
      },
      getSuppliers: async (): Promise<ILookupOption[]> => {
        called.push('suppliers');
        return STUB.suppliers;
      }
    });

    await renderAddForm(async () => undefined, recording);

    expect(called.sort()).toEqual([
      'projectTemplates',
      'spatialBreakdowns',
      'suppliers',
      'tenderTemplates',
      'users'
    ]);
  });

  it('does not render the excluded integration sections', async () => {
    await renderAddForm(async () => undefined);

    expect(host.textContent).not.toContain('EDcontrols');
    expect(host.textContent).not.toContain('Audits');
    expect(host.textContent).not.toContain('KYP');
    expect(host.textContent).not.toContain('SharePoint');
  });

  it('blocks an empty save and shows the required-field messages', async () => {
    const onSave = jest.fn(async () => undefined);
    await renderAddForm(onSave);

    const form: HTMLFormElement = host.querySelector('form') as HTMLFormElement;
    await act(async () => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });

    expect(onSave).not.toHaveBeenCalled();
    expect(host.textContent).toContain('Project  Name is required');
    expect(host.textContent).toContain('Address is required');
    expect(host.textContent).toContain('Postcode is required');
  });

  it('saves a valid form and hands over the typed values', async () => {
    const saved: IBuildingForm[] = [];
    await renderAddForm(async (form: IBuildingForm): Promise<void> => {
      saved.push(form);
    });

    setValue('#project-buildingName', 'Roof replacement');
    setValue('#project-address', 'Keizersgracht 1');
    setValue('#project-postcode', '1015 CJ');

    const form: HTMLFormElement = host.querySelector('form') as HTMLFormElement;
    await act(async () => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });

    expect(saved).toHaveLength(1);
    expect(saved[0].buildingName).toBe('Roof replacement');
    expect(saved[0].address).toBe('Keizersgracht 1');
    expect(saved[0].postcode).toBe('1015 CJ');
    expect(saved[0].clientId).toBe('42');
    // Arrays must never be undefined: the Web API dereferences them without a null check.
    expect(saved[0].userIds).toEqual([]);
    expect(saved[0].spatialBreakdownIds).toEqual([]);
  });

  it('reveals the MOM folder cascade only once its toggle is on', async () => {
    await renderAddForm(async () => undefined);

    expect(host.querySelector('#project-mom-folderId')).toBeNull();

    const toggle: HTMLInputElement = host.querySelector('input[data-field="isAddMomPdf"]') as HTMLInputElement;
    await act(async () => {
      // `.click()` runs the real activation behaviour, which flips `checked` and lets
      // React's value tracker see an actual change. Pre-assigning `checked` instead makes
      // the tracker treat it as unchanged and skip `onChange`.
      toggle.click();
    });

    expect(host.querySelector('#project-mom-folderId')).not.toBeNull();
    // Deeper levels start disabled: nothing is selected above them.
    expect((host.querySelector('#project-mom-subFolderId') as HTMLSelectElement).disabled).toBe(true);
  });

  // Every folder endpoint takes a projectId, so in add mode there is nothing to ask for -
  // the same gap the reference has, where `projectId: $('#Id').val()` is 0 while adding.
  it('does not request folders while adding, since none exist yet', async () => {
    let folderCalls: number = 0;
    const recording: ILookupService = stubLookupService({
      getProjectFolders: async (): Promise<ILookupOption[]> => {
        folderCalls++;
        return STUB.folders;
      }
    });

    await renderAddForm(async () => undefined, recording);

    const toggle: HTMLInputElement = host.querySelector('input[data-field="isAddMomPdf"]') as HTMLInputElement;
    await act(async () => {
      toggle.click();
    });

    expect(folderCalls).toBe(0);
    expect((host.querySelector('#project-mom-folderId') as HTMLSelectElement).disabled).toBe(true);
  });

  it('withholds the form and offers a retry when the dropdown data cannot be loaded', async () => {
    const failing: ILookupService = stubLookupService({
      getUsers: async (): Promise<ILookupOption[]> => {
        throw new Error('The Web API is unreachable.');
      }
    });

    await renderAddForm(async () => undefined, failing);

    expect(host.textContent).toContain('The Web API is unreachable.');
    expect(host.textContent).toContain('Try again');
    // No half-populated form the user could submit.
    expect(host.querySelector('#project-buildingName')).toBeNull();
  });

  it('recovers when a retry succeeds', async () => {
    let shouldFail: boolean = true;
    const flaky: ILookupService = stubLookupService({
      getUsers: async (): Promise<ILookupOption[]> => {
        if (shouldFail) {
          throw new Error('Temporary failure.');
        }
        return STUB.users;
      }
    });

    await renderAddForm(async () => undefined, flaky);
    expect(host.querySelector('#project-buildingName')).toBeNull();

    shouldFail = false;
    await clickText('Try again');

    expect(host.querySelector('#project-buildingName')).not.toBeNull();
  });

  it('refuses the form to a role that may not open it', async () => {
    await renderAddForm(async () => undefined, stubLookupService(), false);

    expect(host.textContent).toContain('does not have permission');
    expect(host.querySelector('form')).toBeNull();
  });
});
