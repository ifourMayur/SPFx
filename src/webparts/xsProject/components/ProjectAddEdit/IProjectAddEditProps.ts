import { IBuildingForm, IProjectFormFeatures } from '../../../../models/Building';
import { ILookupService } from '../../../../services/LookupService';

export interface IProjectAddEditProps {
  /**
   * Reference data for the dropdowns, injected by the web part through `ServiceFactory`.
   *
   * The component depends on the interface only, so it stays free of HTTP concerns and can
   * be rendered against a stub in tests.
   */
  lookupService: ILookupService;
  /**
   * Project being edited. `undefined` (no `:id` in the route) means add mode, which is
   * what switches the Reset button on and unlocks the Project Template dropdown.
   */
  projectId?: number;
  /**
   * Form to start from in edit mode.
   *
   * Absent means the project exists in the listing but its field values are not in memory
   * - the form then starts from {@link fallbackProjectName} and says so, rather than
   * pretending the record loaded. A `GET api/Building?id=` call would remove this case.
   */
  initialForm?: IBuildingForm;
  /** Project name from the listing row, used when {@link initialForm} is absent. */
  fallbackProjectName?: string;
  /** Stands in for the reference view's server-side `@if` gates. */
  features: IProjectFormFeatures;
  /** Signed-in user, who becomes the client of the project. Read-only in the form. */
  clientId: string;
  clientName: string;
  /**
   * Whether this user may use the form at all, from `canOpenProjectForm`. When `false`
   * the component renders a refusal instead of the form, mirroring the reference's
   * `return View("Error")`. The Web API applies its own authorization regardless.
   */
  canSave: boolean;
  /**
   * Persists the form. Rejecting shows its message as the form-level error and keeps the
   * user on the form with their input intact; resolving hands navigation to the caller.
   */
  onSave: (form: IBuildingForm) => Promise<void>;
  /** Leaves the form without saving - the Back button. */
  onCancel: () => void;
}
