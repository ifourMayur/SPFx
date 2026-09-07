import { UserRole, canAddEditProject, canOpenProjectForm } from './Permissions';

/**
 * Mirrors `BuildingController.Index` (lines 112-140), which is the only place the
 * reference application decides who may add or edit a project.
 */
describe('canAddEditProject', () => {
  it('allows an Admin', () => {
    expect(canAddEditProject(UserRole.Admin)).toBe(true);
  });

  it('allows a SuperAdmin', () => {
    expect(canAddEditProject(UserRole.SuperAdmin)).toBe(true);
  });

  it('allows a Client', () => {
    expect(canAddEditProject(UserRole.Client)).toBe(true);
  });

  it('allows a ClientAdmin', () => {
    expect(canAddEditProject(UserRole.ClientAdmin)).toBe(true);
  });

  // Divergence from the reference matrix, on purpose: `AuthenticateController.Add`
  // provisions every SPFx-created user as MaintenanceManager, so denying that role would
  // leave no SPFx user able to reach the form at all.
  it('allows a MaintenanceManager, the role SPFx sign-in provisions', () => {
    expect(canAddEditProject(UserRole.MaintenanceManager)).toBe(true);
  });

  it('denies a Consultant', () => {
    expect(canAddEditProject(UserRole.Consultant)).toBe(false);
  });

  it('denies a Franchisenemer', () => {
    expect(canAddEditProject(UserRole.Franchisenemer)).toBe(false);
  });

  it('denies an McDOffice user', () => {
    expect(canAddEditProject(UserRole.McDOffice)).toBe(false);
  });

  it('denies a role it does not recognise', () => {
    expect(canAddEditProject(9999)).toBe(false);
  });

  it('denies a caller whose role is not known yet', () => {
    expect(canAddEditProject(undefined)).toBe(false);
  });
});

/**
 * The reference gates the *form* more tightly than the listing: `BuildingController.AddEdit`
 * (line 531) returns the error view for a SuperAdmin, even though `Index` shows them the
 * Add button. Reproduced deliberately rather than smoothed over.
 */
describe('canOpenProjectForm', () => {
  it('allows an Admin', () => {
    expect(canOpenProjectForm(UserRole.Admin)).toBe(true);
  });

  it('allows a Client', () => {
    expect(canOpenProjectForm(UserRole.Client)).toBe(true);
  });

  it('allows a ClientAdmin', () => {
    expect(canOpenProjectForm(UserRole.ClientAdmin)).toBe(true);
  });

  it('allows a MaintenanceManager, the role SPFx sign-in provisions', () => {
    expect(canOpenProjectForm(UserRole.MaintenanceManager)).toBe(true);
  });

  it('denies a SuperAdmin, who may see the Add button but not the form', () => {
    expect(canOpenProjectForm(UserRole.SuperAdmin)).toBe(false);
  });

  it('denies a Consultant', () => {
    expect(canOpenProjectForm(UserRole.Consultant)).toBe(false);
  });

  it('denies a caller whose role is not known yet', () => {
    expect(canOpenProjectForm(undefined)).toBe(false);
  });
});
