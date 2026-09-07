import { IBuildingForm, IProjectFormFeatures, createEmptyBuildingForm } from './Building';
import { VALIDATION_MESSAGES, hasErrors, validateBuildingForm } from './BuildingValidation';

/** A form that passes every rule, so each test can break exactly one thing. */
function validForm(overrides?: Partial<IBuildingForm>): IBuildingForm {
  return {
    ...createEmptyBuildingForm(),
    buildingName: 'Roof replacement',
    address: 'Keizersgracht 1',
    postcode: '1015 CJ',
    ...overrides
  };
}

/** Feature gates for a non-McD tenant, where every conditional rule is active. */
function features(overrides?: Partial<IProjectFormFeatures>): IProjectFormFeatures {
  return { domainGroup: 'taskXs', isGrandChildFolderEnabled: true, ...overrides };
}

describe('validateBuildingForm - project name', () => {
  it('requires a project name', () => {
    const errors = validateBuildingForm(validForm({ buildingName: '' }), features());

    expect(errors.buildingName).toBe(VALIDATION_MESSAGES.buildingNameRequired);
  });

  it('rejects a project name of only whitespace', () => {
    const errors = validateBuildingForm(validForm({ buildingName: '   ' }), features());

    expect(errors.buildingName).toBe(VALIDATION_MESSAGES.buildingNameRequired);
  });

  it('accepts a project name', () => {
    const errors = validateBuildingForm(validForm(), features());

    expect(errors.buildingName).toBeUndefined();
  });
});

describe('validateBuildingForm - address', () => {
  it('requires an address', () => {
    const errors = validateBuildingForm(validForm({ address: '' }), features());

    expect(errors.address).toBe(VALIDATION_MESSAGES.addressRequired);
  });

  it('accepts an address', () => {
    const errors = validateBuildingForm(validForm(), features());

    expect(errors.address).toBeUndefined();
  });
});

describe('validateBuildingForm - postcode', () => {
  it('requires a postcode for a non-McD tenant', () => {
    const errors = validateBuildingForm(validForm({ postcode: '' }), features());

    expect(errors.postcode).toBe(VALIDATION_MESSAGES.postcodeRequired);
  });

  // `saveProject` (line 2329) skips the postcode check entirely for the McD domain group,
  // and the server-side [Required] on Postcode is commented out - so nothing enforces it.
  it('does not require a postcode for the McD domain group', () => {
    const errors = validateBuildingForm(validForm({ postcode: '' }), features({ domainGroup: 'mcD' }));

    expect(errors.postcode).toBeUndefined();
  });

  it('accepts a postcode', () => {
    const errors = validateBuildingForm(validForm(), features());

    expect(errors.postcode).toBeUndefined();
  });
});

describe('validateBuildingForm - email', () => {
  it('treats an empty email as valid, because the field is optional', () => {
    const errors = validateBuildingForm(validForm({ email: '' }), features());

    expect(errors.email).toBeUndefined();
  });

  it('rejects an email that does not match the reference pattern', () => {
    const errors = validateBuildingForm(validForm({ email: 'not-an-email' }), features());

    expect(errors.email).toBe(VALIDATION_MESSAGES.emailInvalid);
  });

  it('rejects an email whose domain has no dot', () => {
    const errors = validateBuildingForm(validForm({ email: 'someone@localhost' }), features());

    expect(errors.email).toBe(VALIDATION_MESSAGES.emailInvalid);
  });

  it('accepts a well formed email', () => {
    const errors = validateBuildingForm(validForm({ email: 'jan.jansen@example.co.uk' }), features());

    expect(errors.email).toBeUndefined();
  });

  it('rejects an email longer than 255 characters', () => {
    const localPart: string = new Array(250).join('a');
    const errors = validateBuildingForm(validForm({ email: `${localPart}@example.com` }), features());

    expect(errors.email).toBe(VALIDATION_MESSAGES.emailTooLong);
  });
});

describe('validateBuildingForm - phone', () => {
  it('treats an empty phone number as valid, because the field is optional', () => {
    const errors = validateBuildingForm(validForm({ telefoon: '' }), features());

    expect(errors.telefoon).toBeUndefined();
  });

  it('accepts digits, spaces and a leading plus', () => {
    const errors = validateBuildingForm(validForm({ telefoon: '+31 20 123 4567' }), features());

    expect(errors.telefoon).toBeUndefined();
  });

  it('rejects a phone number containing letters', () => {
    const errors = validateBuildingForm(validForm({ telefoon: '020-CALL-NOW' }), features());

    expect(errors.telefoon).toBe(VALIDATION_MESSAGES.telefoonInvalid);
  });
});

describe('validateBuildingForm - tender template', () => {
  it('requires a tender template when tender is enabled', () => {
    const errors = validateBuildingForm(
      validForm({ isEnableTender: true, tenderTemplateId: '' }),
      features()
    );

    expect(errors.tenderTemplateId).toBe(VALIDATION_MESSAGES.tenderTemplateRequired);
  });

  it('accepts an enabled tender that has a template selected', () => {
    const errors = validateBuildingForm(
      validForm({ isEnableTender: true, tenderTemplateId: '4' }),
      features()
    );

    expect(errors.tenderTemplateId).toBeUndefined();
  });

  it('ignores a missing tender template when tender is disabled', () => {
    const errors = validateBuildingForm(
      validForm({ isEnableTender: false, tenderTemplateId: '' }),
      features()
    );

    expect(errors.tenderTemplateId).toBeUndefined();
  });
});

describe('validateBuildingForm - minutes of meeting', () => {
  // `getAllMOMData` (line 2431) initialises `allValid = true` and never clears it, and
  // `#MOMfolderNotSelect` is never written to, so the reference lets a MOM row through
  // with no folder chosen. Reproduced rather than tightened.
  it('does not require a MOM folder even when saving minutes is enabled', () => {
    const errors = validateBuildingForm(validForm({ isAddMomPdf: true }), features());

    expect(hasErrors(errors)).toBe(false);
  });
});

describe('validateBuildingForm - whole form', () => {
  it('reports no errors for a valid form', () => {
    expect(validateBuildingForm(validForm(), features())).toEqual({});
  });

  it('reports every broken rule at once rather than stopping at the first', () => {
    const errors = validateBuildingForm(
      validForm({ buildingName: '', address: '', postcode: '', email: 'bad' }),
      features()
    );

    expect(errors.buildingName).toBeDefined();
    expect(errors.address).toBeDefined();
    expect(errors.postcode).toBeDefined();
    expect(errors.email).toBeDefined();
  });
});

describe('hasErrors', () => {
  it('is false for an empty error object', () => {
    expect(hasErrors({})).toBe(false);
  });

  it('is true when any field carries a message', () => {
    expect(hasErrors({ postcode: VALIDATION_MESSAGES.postcodeRequired })).toBe(true);
  });
});

describe('createEmptyBuildingForm', () => {
  it('starts in add mode', () => {
    expect(createEmptyBuildingForm().id).toBe(0);
  });

  it('starts with every multi-select empty rather than undefined', () => {
    const form: IBuildingForm = createEmptyBuildingForm();

    expect(form.userIds).toEqual([]);
    expect(form.spatialBreakdownIds).toEqual([]);
  });

  it('starts with an empty MOM folder selection', () => {
    expect(createEmptyBuildingForm().momFolder).toEqual({
      folderId: '',
      subFolderId: '',
      subSubFolderId: '',
      subSubSubFolderId: '',
      supplierId: ''
    });
  });
});
