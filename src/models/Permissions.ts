/**
 * Who may add or edit a project.
 *
 * Part of the models layer: pure, dependency-free and unit-testable, with no imports from
 * `config`, `services`, `components` or any SPFx package.
 *
 * These are the reference application's own rules, expressed against the `userRoleId` the
 * Web API already returns in the sign-in response (`ILoginResponse.userRoleId`), so no
 * extra call is needed to evaluate them.
 *
 * A UI check is a convenience, never a security boundary: the Web API applies its own
 * `[Authorize]` and role checks on every request, and remains the authority.
 */

/** Roles the Web API issues, matching `BMDesk.Business.Enums.UserRoleEnum`. */
export enum UserRole {
  Admin = 1,
  SuperAdmin = 2,
  /** "Manager McD" in the reference UI. */
  Client = 3,
  /** "Uitvoerende partij". */
  MaintenanceManager = 4,
  /** "Manager Extern". */
  ClientAdmin = 5,
  Agent = 6,
  /** "Adviseur". */
  Consultant = 7,
  Franchisenemer = 8,
  Office = 9,
  McDOffice = 10
}

/**
 * Roles allowed to add or edit a project. Every other role is denied.
 *
 * The first four are the ones the reference grants `ViewBag.AddEditAccessRights` to.
 *
 * `MaintenanceManager` is a **deliberate divergence** from that matrix, and the reason is
 * worth knowing before anyone "corrects" it: `AuthenticateController.Add` hard-codes
 * `UserRole = (int)UserRoleEnum.MaintenanceManager` for every user it creates, and that is
 * the path `SPFxLogin` uses - so every user arriving through this web part is a
 * MaintenanceManager. Denying that role, as the reference does, left no SPFx user able to
 * see the Add button or open the form at all.
 *
 * The role means "Uitvoerende partij" (an external executing party) in the MVC
 * application's own user population, which is why it is denied there. Users of this web
 * part are internal Microsoft 365 users of the tenant, who reach it only after a
 * successful `SPFxLogin`; they merely inherit that role by accident of provisioning.
 *
 * This grants nothing the Web API would refuse: `BuildingController.Add`'s role guard,
 * `userRoleId != MaintenanceManager || userRoleId != Consultant || ...`, is a tautology -
 * one integer cannot differ from two values at once - so the API imposes no role
 * restriction on creating a project. The remaining denied roles (Consultant,
 * Franchisenemer, McDOffice) keep the reference's behaviour.
 */
const ADD_EDIT_ROLES: UserRole[] = [
  UserRole.Admin,
  UserRole.SuperAdmin,
  UserRole.Client,
  UserRole.ClientAdmin,
  UserRole.MaintenanceManager
];

/**
 * True when the role may add or edit projects, and therefore see the Add button and the
 * Edit row action in the project listing.
 *
 * Follows `BuildingController.Index` (lines 112-140) apart from `MaintenanceManager` -
 * see {@link ADD_EDIT_ROLES} for why that one is allowed here. Consultant, Franchisenemer,
 * McDOffice and any unlisted role are denied. An undefined role - sign-in has not
 * completed yet - is treated as denied rather than assumed harmless.
 */
export function canAddEditProject(userRoleId: number | undefined): boolean {
  if (userRoleId === undefined) {
    return false;
  }

  return ADD_EDIT_ROLES.indexOf(userRoleId) >= 0;
}

/**
 * True when the role may open the project form itself.
 *
 * Deliberately stricter than {@link canAddEditProject}: `BuildingController.AddEdit`
 * (line 531) returns the error view for a SuperAdmin, even though `Index` shows that same
 * user the Add button. The reference is inconsistent here; keeping the two gates as
 * separate named functions preserves that behavior instead of silently picking one.
 */
export function canOpenProjectForm(userRoleId: number | undefined): boolean {
  return canAddEditProject(userRoleId) && userRoleId !== UserRole.SuperAdmin;
}
