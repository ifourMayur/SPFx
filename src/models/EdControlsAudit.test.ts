import {
  AUDIT_MESSAGES,
  IAudit,
  IAuditForm,
  createAuditForm,
  createDuplicateDrafts,
  formatDueDate,
  parseDueDate,
  parseDuplicateCount,
  toAuditSaveRequest,
  toAudits,
  toContactOptions,
  toIsoDate,
  toSaveResults,
  toTemplateOptions,
  validateAuditForm
} from './EdControlsAudit';

const AUDIT: IAudit = toAudits([
  {
    id: 4,
    buildingId: 9,
    auditName: 'Fire safety',
    edControlsDatabase: 'db-1',
    edControlsTemplateId: 't-1',
    edControlsTemplateName: 'Template 1',
    supplierId: 7,
    supplierName: 'Van Dijk',
    contactId: null,
    contactName: null,
    dueDate: '2026-12-31T00:00:00',
    folderId: 10,
    subFolderId: 11,
    childSubFolderId: null,
    grandChildSubFolderId: null,
    folderPath: 'Docs / Safety',
    status: 2
  }
])[0];

function validForm(changes: Partial<IAuditForm> = {}): IAuditForm {
  return {
    ...createAuditForm(),
    auditName: ' Fire ',
    templateId: 't-1',
    templateName: 'Template 1',
    supplierId: '7',
    dueDate: '31/12/2026',
    folder: { folderId: '10', subFolderId: '11', subSubFolderId: '', subSubSubFolderId: '', supplierId: '' },
    ...changes
  };
}

describe('EdControlsAudit', () => {
  it('reads the audit list, nulls becoming absent ids and empty names', () => {
    expect(AUDIT).toMatchObject({ id: 4, supplierId: 7, contactId: undefined, contactName: '', subFolderId: 11, childSubFolderId: undefined, status: 2 });
    expect(toAudits([{ auditName: 'no id' }, 'x'])).toEqual([]);
  });

  it('reads templates by their string id, and contacts by name or else e-mail, sorted', () => {
    expect(toTemplateOptions([{ id: 'abc', name: '' }, { name: 'no id' }])).toEqual([{ id: 'abc', name: 'abc' }]);
    expect(toContactOptions([
      { id: 2, name: '', email: 'b@x.nl' },
      { id: 1, name: 'Anna', email: 'a@x.nl' },
      { id: 3, name: '', email: '' }
    ])).toEqual([{ id: '1', name: 'Anna' }, { id: '2', name: 'b@x.nl' }]);
  });

  it('reads the bulk outcome per row', () => {
    expect(toSaveResults([{ index: 1, auditName: 'A', success: false, message: 'No', id: 0 }, {}])).toEqual([
      { index: 1, auditName: 'A', success: false, message: 'No', id: 0 }
    ]);
  });

  describe('dates', () => {
    it('accepts only a real dd/mm/yyyy date', () => {
      expect(parseDueDate('31/12/2026')).toEqual(new Date(2026, 11, 31));
      expect(parseDueDate('31/02/2026')).toBeUndefined();
      expect(parseDueDate('2026-12-31')).toBeUndefined();
      expect(toIsoDate('1/2/2026')).toBe('2026-02-01');
    });

    it('shows an API date without shifting it by the time zone', () => {
      expect(formatDueDate('2026-12-31T00:00:00')).toBe('31/12/2026');
      expect(formatDueDate('')).toBe('');
    });
  });

  describe('validation', () => {
    it('needs a name, a template, a supplier, a valid due date and a folder', () => {
      expect(validateAuditForm(createAuditForm())).toEqual({
        auditName: AUDIT_MESSAGES.nameRequired,
        templateId: AUDIT_MESSAGES.templateRequired,
        supplierId: AUDIT_MESSAGES.supplierRequired,
        dueDate: AUDIT_MESSAGES.dueDateRequired,
        folderId: AUDIT_MESSAGES.folderRequired
      });
      expect(validateAuditForm(validForm({ dueDate: '32/01/2026' }))).toEqual({ dueDate: AUDIT_MESSAGES.dueDateInvalid });
      expect(validateAuditForm(validForm())).toEqual({});
    });
  });

  it('posts ids as numbers, the date as yyyy-MM-dd and empty levels as absent', () => {
    expect(toAuditSaveRequest(validForm(), 9, 'db-1')).toEqual({
      buildingId: 9,
      auditName: 'Fire',
      edControlsDatabase: 'db-1',
      edControlsTemplateId: 't-1',
      edControlsTemplateName: 'Template 1',
      supplierId: 7,
      contactId: undefined,
      dueDate: '2026-12-31',
      folderId: 10,
      subFolderId: 11,
      childSubFolderId: undefined,
      grandChildSubFolderId: undefined
    });
  });

  describe('duplicates', () => {
    it('accepts a whole number from 1 to 20 only', () => {
      expect(parseDuplicateCount('3')).toBe(3);
      ['0', '21', '2.5', '', 'x', '-1'].forEach((value: string) => expect(parseDuplicateCount(value)).toBeUndefined());
    });

    it('copies the audit into numbered drafts with its template, supplier, date and folder', () => {
      const drafts = createDuplicateDrafts(AUDIT, 2);

      expect(drafts.map((draft) => draft.form.auditName)).toEqual(['Fire safety - 1', 'Fire safety - 2']);
      expect(drafts[0].form).toMatchObject({
        templateId: 't-1',
        supplierId: '7',
        contactId: '',
        dueDate: '31/12/2026',
        folderPath: 'Docs / Safety',
        folder: { folderId: '10', subFolderId: '11', subSubFolderId: '', subSubSubFolderId: '' }
      });
      expect(drafts[0].key).not.toBe(drafts[1].key);
    });

    it('keeps a draft name within the limit', () => {
      const long: IAudit = { ...AUDIT, auditName: 'x'.repeat(250) };

      expect(createDuplicateDrafts(long, 1)[0].form.auditName).toHaveLength(250);
    });
  });
});
