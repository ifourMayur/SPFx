/**
 * Validation for the project add/edit form.
 *
 * Pure and dependency-free, in the models layer, so the rules can be unit-tested without
 * rendering anything - see `BuildingValidation.test.ts`.
 *
 * The reference application splits these rules across two places, and both are reproduced
 * here:
 *
 *   1. `BuildingAddModel` data annotations, enforced server-side - project name and
 *      address required, the email length/format, the phone format.
 *   2. `saveProject` in `Views/Building/AddEdit.cshtml` (line 2323), enforced only in the
 *      browser - the postcode (skipped for McD) and the tender template.
 *
 * Two deliberate differences from the reference, both widening what the user is told
 * rather than changing what counts as valid:
 *
 *   - `saveProject` returns on the first failure, so the user fixes one field per attempt.
 *     Every rule is evaluated here and all messages are shown at once.
 *   - The reference shows the email and phone messages through ASP.NET's unobtrusive
 *     validation, which only runs on the round trip. They are checked here directly.
 */

import { IBuildingForm, IBuildingFormErrors, IProjectFormFeatures } from './Building';

/**
 * Message text, taken verbatim from the reference's resource files so a ported screen
 * reads identically to the one it replaces.
 *
 * `buildingNameRequired` carries the double space `BuildingResource.reqProjectName` has.
 * The tender message is the one string the reference hard-codes in the view rather than
 * localising (AddEdit.cshtml line 2377).
 */
export const VALIDATION_MESSAGES = {
  /** `BuildingResource.reqProjectName` */
  buildingNameRequired: 'Project  Name is required',
  /** `BuildingResource.reqAddress` */
  addressRequired: 'Address is required',
  /** `BuildingResource.reqPostcode` */
  postcodeRequired: 'Postcode is required',
  /** `BuildingResource.msgEmail` */
  emailInvalid: 'Invalid Email Address',
  /** `BuildingResource.maxEmail` */
  emailTooLong: 'Max length of email is 255 character',
  /** `ContactResource.msgTelefoon` */
  telefoonInvalid: 'Invalid Phone Number',
  /** Hard-coded in AddEdit.cshtml, not localised there either. */
  tenderTemplateRequired: 'Tender template is required'
} as const;

/** `Constant.EmailRegex` from the reference. */
const EMAIL_PATTERN: RegExp = /^[\w-.]+@([\w-]+\.)+[\w-]{2,}$/;

/** `Constant.MobileNumberRegex` from the reference: digits, spaces and `+` only. */
const PHONE_PATTERN: RegExp = /^[\d +]*$/;

/** `[MaxLength(255)]` on `BuildingAddModel.Email`. */
const EMAIL_MAX_LENGTH: number = 255;

/** True when a text field holds nothing but whitespace, which the reference treats as empty. */
function isBlank(value: string | undefined): boolean {
  return !value || value.trim().length === 0;
}

/**
 * Validates the whole form and returns a message per broken rule.
 *
 * An empty object means the form may be saved. Use {@link hasErrors} rather than checking
 * the object's truthiness, which is always `true`.
 */
export function validateBuildingForm(
  form: IBuildingForm,
  featureFlags: IProjectFormFeatures
): IBuildingFormErrors {
  const errors: IBuildingFormErrors = {};

  if (isBlank(form.buildingName)) {
    errors.buildingName = VALIDATION_MESSAGES.buildingNameRequired;
  }

  if (isBlank(form.address)) {
    errors.address = VALIDATION_MESSAGES.addressRequired;
  }

  // McD is the one tenant grouping the reference exempts, and the server-side [Required]
  // on Postcode is commented out - so for McD nothing enforces it at all.
  if (featureFlags.domainGroup !== 'mcD' && isBlank(form.postcode)) {
    errors.postcode = VALIDATION_MESSAGES.postcodeRequired;
  }

  const emailError: string | undefined = validateEmail(form.email);
  if (emailError) {
    errors.email = emailError;
  }

  // Optional, but must be digits/spaces/plus when supplied.
  if (!isBlank(form.telefoon) && !PHONE_PATTERN.test(form.telefoon)) {
    errors.telefoon = VALIDATION_MESSAGES.telefoonInvalid;
  }

  if (form.isEnableTender && isBlank(form.tenderTemplateId)) {
    errors.tenderTemplateId = VALIDATION_MESSAGES.tenderTemplateRequired;
  }

  // No MOM rule on purpose: `getAllMOMData` collects the row without ever setting its
  // `allValid` flag to false, and never writes `#MOMfolderNotSelect`, so the reference
  // saves a MOM selection with no folder chosen.

  return errors;
}

/**
 * Length before format, matching the order the two attributes are declared in on
 * `BuildingAddModel.Email` - an over-long address reports the length problem, which is
 * the more actionable of the two.
 */
function validateEmail(email: string): string | undefined {
  if (isBlank(email)) {
    return undefined;
  }

  if (email.length > EMAIL_MAX_LENGTH) {
    return VALIDATION_MESSAGES.emailTooLong;
  }

  if (!EMAIL_PATTERN.test(email)) {
    return VALIDATION_MESSAGES.emailInvalid;
  }

  return undefined;
}

/** True when at least one field carries a message, so the form must not be saved. */
export function hasErrors(errors: IBuildingFormErrors): boolean {
  return Object.keys(errors).some((field: string): boolean => {
    return !!errors[field as keyof IBuildingFormErrors];
  });
}
