/**
 * End-to-end flow tests for the web part's routed shell.
 *
 * Drives the real `XsProject` - stubbing only the two injected services - through
 * sign-in -> listing -> Add -> save -> listing, and then reopens the created project for
 * editing, which is the flow this screen exists to serve. They also pin the permission
 * gates and the McD tenant behaviour to actual rendered output rather than to the pure
 * functions alone.
 *
 * Every `ReactDOM.render` here is unmounted in `afterEach`.
 */

// `@microsoft/sp-core-library` needs the SPFx page runtime, which plain jsdom has not got;
// `Login` and `Project` both pull it in only for `Log`. This must precede the imports
// below: the TypeScript build emits CommonJS without Babel's `jest.mock` hoisting, so the
// mock has to be registered before the module graph is required.
jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    verbose: (): void => undefined,
    info: (): void => undefined,
    warn: (): void => undefined,
    error: (): void => undefined
  }
}));

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';

import XsProject from './XsProject';
import { UserRole } from '../../../../models/Permissions';
import type { ILoginResponse, ISPFxLoginRequest } from '../../../../models/Auth';
import type { ILoginService } from '../../../../services/LoginService';
import type { ILookupService } from '../../../../services/LookupService';
import type { IProjectService } from '../../../../services/ProjectService';
import type { ILookupOption } from '../../../../models/Building';

function loginResponse(userRoleId: number, domainGroup: number): ILoginResponse {
  return {
    id: 42,
    isExist: true,
    token: 'app-jwt',
    expiration: '2030-01-01T00:00:00Z',
    userName: 'mayur@ifourtechnolab.com',
    isActive: true,
    userRoleId,
    languageId: 1,
    supplierId: 0,
    accountId: 7,
    subscriptionsId: 1,
    isNotification: false,
    tasksNotification: false,
    messagesNotification: false,
    approvalsNotification: false,
    taskEmailNotification: false,
    approveEmailNotification: false,
    domainGroup,
    userMultipleDomain: 0,
    enableAzureSSO: false,
    isNewProjectListLayout: false,
    isFavoriteProject: false,
    isDocumentPreview: false
  } as ILoginResponse;
}

/** A lookup service that resolves immediately, so the form's dropdown load settles. */
function stubLookupService(): ILookupService {
  const options: ILookupOption[] = [{ id: '1', name: 'Base template' }];

  return {
    getProjectTemplates: async (): Promise<ILookupOption[]> => options,
    getUsers: async (): Promise<ILookupOption[]> => options,
    getSpatialBreakdowns: async (): Promise<ILookupOption[]> => options,
    getTenderTemplates: async (): Promise<ILookupOption[]> => options,
    getSuppliers: async (): Promise<ILookupOption[]> => options,
    getProjectFolders: async (): Promise<ILookupOption[]> => [],
    getProjectSubFolders: async (): Promise<ILookupOption[]> => [],
    getProjectSubSubFolders: async (): Promise<ILookupOption[]> => [],
    getProjectSubSubSubFolders: async (): Promise<ILookupOption[]> => []
  };
}

function stubServices(userRoleId: number, domainGroup: number): {
  loginService: ILoginService;
  projectService: IProjectService;
  lookupService: ILookupService;
} {
  const response: ILoginResponse = loginResponse(userRoleId, domainGroup);

  return {
    lookupService: stubLookupService(),
    loginService: {
      login: async (_request: ISPFxLoginRequest) => response,
      loginAsCurrentUser: async () => response
    } as ILoginService,
    projectService: {
      getProjects: async () => [],
      getProjectById: async () => ({}),
      createProject: async () => ({}),
      updateProject: async () => ({}),
      deleteProject: async () => undefined
    } as unknown as IProjectService
  };
}

describe('XsProject', () => {
  let host: HTMLDivElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    window.location.hash = '';
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(host);
    host.remove();
  });

  async function mount(userRoleId: number = UserRole.Admin, domainGroup: number = 2): Promise<void> {
    const services = stubServices(userRoleId, domainGroup);

    await act(async () => {
      ReactDOM.render(
        <XsProject
          description="d"
          isDarkTheme={false}
          environmentMessage="e"
          userDisplayName="Mayur"
          projectService={services.projectService}
          lookupService={services.lookupService}
          loginService={services.loginService}
        />,
        host
      );
    });
  }

  /** Clicks the first button whose visible text matches. */
  async function clickText(text: string): Promise<void> {
    const buttons: HTMLButtonElement[] = Array.prototype.slice.call(host.querySelectorAll('button'));
    // Exact match first, then a prefix match: a parent menu item renders its label
    // followed by a chevron ("Project▾").
    const label = (button: HTMLButtonElement): string => (button.textContent || '').trim();
    const target: HTMLButtonElement | undefined =
      buttons.filter((button: HTMLButtonElement) => label(button) === text)[0]
      || buttons.filter((button: HTMLButtonElement) => label(button).indexOf(text) === 0)[0];

    if (!target) {
      throw new Error(`No button labelled "${text}". Buttons: ${buttons.map((b) => `"${(b.textContent || '').trim()}"`).join(', ')}`);
    }

    await act(async () => {
      target.click();
    });
  }

  function setValue(selector: string, value: string): void {
    const input: HTMLInputElement = host.querySelector(selector) as HTMLInputElement;
    if (!input) {
      throw new Error(`No element for ${selector}`);
    }

    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      setter?.call(input, value);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }

  it('lands on the dashboard after sign-in', async () => {
    await mount();

    expect(window.location.hash).toBe('#/dashboard');
  });

  it('shows the gated Add button to an Admin and reaches the form with it', async () => {
    await mount(UserRole.Admin);

    await clickText('Project');
    expect(window.location.hash).toBe('#/projects');
    expect(host.textContent).toContain('+ Add Project');

    await clickText('+ Add Project');
    expect(window.location.hash).toBe('#/projects/add');
    expect(host.querySelector('#project-buildingName')).not.toBeNull();
  });

  it('shows the Add button to a MaintenanceManager, the role SPFx sign-in provisions', async () => {
    await mount(UserRole.MaintenanceManager);

    await clickText('Project');
    expect(host.textContent).toContain('+ Add Project');
  });

  it('hides the Add button from a Consultant', async () => {
    await mount(UserRole.Consultant);

    await clickText('Project');
    expect(host.textContent).not.toContain('+ Add Project');
  });

  it('creates a project, returns to the listing, and shows it there', async () => {
    await mount(UserRole.Admin);
    await clickText('Project');
    await clickText('+ Add Project');

    setValue('#project-buildingName', 'Smoke test project');
    setValue('#project-address', 'Keizersgracht 1');
    setValue('#project-postcode', '1015 CJ');

    await clickText('Save');

    // Navigated back to the listing, with the success message and the new row.
    expect(window.location.hash).toBe('#/projects');
    expect(host.textContent).toContain('Project  Added successfully!');
    expect(host.textContent).toContain('Smoke test project');
    // 17 sample rows + the new one.
    expect(host.textContent).toContain('of 18 entries');
  });

  it('reopens a created project for editing with its values intact', async () => {
    await mount(UserRole.Admin);
    await clickText('Project');
    await clickText('+ Add Project');

    setValue('#project-buildingName', 'Reopen me');
    setValue('#project-address', 'Damrak 2');
    setValue('#project-postcode', '1012 LP');
    await clickText('Save');

    // A newly created project is the first row now, so its menu is the first one.
    const optionButtons: HTMLButtonElement[] = Array.prototype.slice.call(
      host.querySelectorAll('td button[aria-haspopup="menu"]')
    );
    await act(async () => {
      optionButtons[0].click();
    });
    await clickText('Bewerken');

    expect(window.location.hash).toBe('#/projects/edit/18');
    expect((host.querySelector('#project-buildingName') as HTMLInputElement).value).toBe('Reopen me');
    expect((host.querySelector('#project-address') as HTMLInputElement).value).toBe('Damrak 2');
    // Edit mode locks the template and drops Reset.
    expect((host.querySelector('#project-projectTemplateId') as HTMLSelectElement).disabled).toBe(true);
    expect(host.textContent).not.toContain('Reset');
  });

  it('updates rather than appends when saving an edit', async () => {
    await mount(UserRole.Admin);
    await clickText('Project');
    await clickText('+ Add Project');
    setValue('#project-buildingName', 'First name');
    setValue('#project-address', 'Damrak 2');
    setValue('#project-postcode', '1012 LP');
    await clickText('Save');

    const optionButtons: HTMLButtonElement[] = Array.prototype.slice.call(
      host.querySelectorAll('td button[aria-haspopup="menu"]')
    );
    await act(async () => {
      optionButtons[0].click();
    });
    await clickText('Bewerken');

    setValue('#project-buildingName', 'Renamed');
    await clickText('Save');

    expect(host.textContent).toContain('Project  Updated successfully!');
    expect(host.textContent).toContain('Renamed');
    expect(host.textContent).not.toContain('First name');
    // Still 18: an edit must not create a second row.
    expect(host.textContent).toContain('of 18 entries');
  });

  it('requires no postcode for the McD domain group', async () => {
    await mount(UserRole.Admin, 1);
    await clickText('Project');
    await clickText('+ Add Project');

    // McD also hides Spatial Breakdown.
    expect(host.querySelector('#project-spatialBreakdownIds')).toBeNull();

    setValue('#project-buildingName', 'McD project');
    setValue('#project-address', 'Hoofdstraat 3');
    await clickText('Save');

    expect(window.location.hash).toBe('#/projects');
    expect(host.textContent).toContain('McD project');
  });

  it('keeps the user on the form when validation fails', async () => {
    await mount(UserRole.Admin);
    await clickText('Project');
    await clickText('+ Add Project');

    setValue('#project-address', 'Damrak 2');
    setValue('#project-postcode', '1012 LP');
    // Project name left blank.
    await clickText('Save');

    expect(window.location.hash).toBe('#/projects/add');
    expect(host.textContent).toContain('Project  Name is required');
    // The values already typed survive the failed attempt.
    expect((host.querySelector('#project-address') as HTMLInputElement).value).toBe('Damrak 2');
  });

  it('refuses the form to a SuperAdmin, who still sees the Add button', async () => {
    await mount(UserRole.SuperAdmin);
    await clickText('Project');

    // The listing gate lets a SuperAdmin through...
    expect(host.textContent).toContain('+ Add Project');

    await clickText('+ Add Project');

    // ...but the form gate does not, matching the reference's `return View("Error")`.
    expect(host.textContent).toContain('does not have permission');
    expect(host.querySelector('#project-buildingName')).toBeNull();
  });
});
