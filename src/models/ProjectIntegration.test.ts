import {
  IEdControlsForm,
  IKypForm,
  INTEGRATION_MESSAGES,
  ITagRow,
  createTagRow,
  toEdControlsForm,
  toEdControlsSaveFields,
  toEdControlsSettings,
  toKypForm,
  toKypSaveFields,
  toKypSettings,
  toTagData,
  toTagList,
  toTagOptions,
  validateTagRows
} from './ProjectIntegration';
import { toBuildingSaveRequest, IPreservedProjectFields, IProjectSaveContext } from './BuildingSave';
import { createEmptyBuildingForm, IBuildingForm } from './Building';
import { validateBuildingForm } from './BuildingValidation';

function row(changes: Partial<ITagRow>): ITagRow {
  return { ...createTagRow(), ...changes };
}

function folder(folderId: string): ITagRow['folder'] {
  return { folderId, subFolderId: '11', subSubFolderId: '', subSubSubFolderId: '', supplierId: '7' };
}

const CONTEXT: IProjectSaveContext = {
  clientId: 42,
  clientName: 'Mayur',
  accountId: 3,
  domainName: 'contoso.sharepoint.com',
  companyName: 'Contoso',
  companyLogo: '',
  site: undefined
};

const PRESERVED: IPreservedProjectFields = {
  documentStorageType: 2,
  tagData: '[{"tags":["old"]}]',
  ticketSourceId: 1,
  ticketProjectId: 'stored-db',
  planningSourcetId: 2,
  planningProjectId: 'stored-kyp',
  kypProjectAuthToken: 'stored-token'
};

function validForm(changes: Partial<IBuildingForm> = {}): IBuildingForm {
  return { ...createEmptyBuildingForm(), id: 9, buildingName: 'P', address: 'A', postcode: '1000 AA', ...changes };
}

describe('ProjectIntegration', () => {
  describe('reading the API', () => {
    it('reads the ED Controls settings, with the stored tag rows', () => {
      const settings = toEdControlsSettings({
        isConfigured: true,
        isTicketSource: true,
        ticketProjectId: 'db-1',
        tagRows: [
          {
            periodic: true,
            documentName: ' Weekly ',
            tags: ['a', 'a', '', null, 'b'],
            folderId: '10',
            subFolderId: null,
            subSubFolderId: '',
            subSubSubFolderId: null,
            supplierId: '7'
          },
          'not a row'
        ]
      });

      expect(settings?.isConfigured).toBe(true);
      expect(settings?.ticketProjectId).toBe('db-1');
      expect(settings?.tagRows).toHaveLength(1);
      expect(settings?.tagRows[0]).toMatchObject({
        periodic: true,
        documentName: 'Weekly',
        tags: ['a', 'b'],
        folder: { folderId: '10', subFolderId: '', subSubFolderId: '', subSubSubFolderId: '', supplierId: '7' }
      });
    });

    it('refuses a payload that is not an object', () => {
      expect(toEdControlsSettings('x')).toBeUndefined();
      expect(toKypSettings(null)).toBeUndefined();
    });

    it('reads the KYP settings without ever expecting a token', () => {
      expect(toKypSettings({ isConfigured: true, isPlanningSource: true, planningProjectId: 'k', hasProjectToken: true })).toEqual({
        isConfigured: true,
        isPlanningSource: true,
        planningProjectId: 'k',
        hasProjectToken: true
      });
    });

    it('reads a tag list, dropping blanks and duplicates', () => {
      expect(toTagList(['x', ' x ', '', 3, null])).toEqual(['x', '3']);
      expect(toTagList({})).toEqual([]);
    });
  });

  describe('starting values', () => {
    it('offers one empty tag row in edit mode when none is stored, none while adding', () => {
      const settings = { isConfigured: true, isTicketSource: false, ticketProjectId: '', tagRows: [] };

      expect(toEdControlsForm(settings, true).tagRows).toHaveLength(1);
      expect(toEdControlsForm(settings, false).tagRows).toHaveLength(0);
    });

    it('ticks "Use alternative token" when the project stores one, with an empty field', () => {
      expect(toKypForm({ isConfigured: true, isPlanningSource: true, planningProjectId: 'k', hasProjectToken: true })).toEqual({
        isPlanningSource: true,
        planningProjectId: 'k',
        isUseProjectToken: true,
        projectToken: '',
        hasStoredProjectToken: true
      });
    });

    it('keeps a selected tag selectable when the project no longer has it', () => {
      expect(toTagOptions(['a'], ['a', 'gone'])).toEqual([
        { id: 'a', name: 'a' },
        { id: 'gone', name: 'gone' }
      ]);
    });
  });

  describe('tag rows', () => {
    it('saves only rows with tags, in the reference JSON shape', () => {
      const rows: ITagRow[] = [
        row({ tags: ['a'], periodic: false, documentName: 'ignored', folder: folder('10') }),
        row({ tags: [] })
      ];

      expect(JSON.parse(toTagData(rows) as string)).toEqual([
        {
          periodic: false,
          documentName: '',
          tags: ['a'],
          folderId: '10',
          subFolderId: '11',
          subSubFolderId: '',
          subSubSubFolderId: '',
          supplierId: '7'
        }
      ]);
      expect(toTagData([row({ tags: [] })])).toBeUndefined();
    });

    it('needs a folder on a row with tags when saving, and tags and a folder on every row when adding one', () => {
      const withoutFolder: ITagRow = row({ tags: ['a'] });
      const empty: ITagRow = row({});

      expect(validateTagRows([withoutFolder, empty], false)).toEqual({
        [withoutFolder.key]: { folderId: INTEGRATION_MESSAGES.folderRequired }
      });
      expect(validateTagRows([empty], true)).toEqual({
        [empty.key]: { tags: INTEGRATION_MESSAGES.tagRequired, folderId: INTEGRATION_MESSAGES.folderRequired }
      });
      expect(validateTagRows([row({ tags: ['a'], folder: folder('10') })], true)).toEqual({});
    });
  });

  describe('saving', () => {
    const ed: IEdControlsForm = { isTicketSource: true, ticketProjectId: 'db-2', tagRows: [row({ tags: ['t'], folder: folder('10') })] };
    const kyp: IKypForm = {
      isPlanningSource: true,
      planningProjectId: 'k-2',
      isUseProjectToken: false,
      projectToken: '',
      hasStoredProjectToken: false
    };

    it('adds nothing for a section that was not loaded', () => {
      expect(toEdControlsSaveFields(undefined)).toEqual({});
      expect(toKypSaveFields(undefined, 'stored')).toEqual({});
    });

    it('echoes the stored values when the sections were not loaded', () => {
      const request = toBuildingSaveRequest(validForm(), CONTEXT, PRESERVED);

      expect(request).toMatchObject({
        tagData: PRESERVED.tagData,
        ticketSourceId: 1,
        ticketProjectId: 'stored-db',
        planningSourcetId: 2,
        planningProjectId: 'stored-kyp',
        kypProjectAuthToken: 'stored-token'
      });
    });

    it('posts what the loaded sections hold, over the stored values', () => {
      const request = toBuildingSaveRequest(validForm({ edControls: ed, kyp }), CONTEXT, PRESERVED);

      expect(request.ticketSourceId).toBe(1);
      expect(request.ticketProjectId).toBe('db-2');
      expect(JSON.parse(request.tagData as string)[0].tags).toEqual(['t']);
      expect(request.planningSourcetId).toBe(2);
      expect(request.planningProjectId).toBe('k-2');
      // "Use alternative token" is off, so the stored one is removed, as the reference does.
      expect(request.kypProjectAuthToken).toBe('');
    });

    it('unlinks ED Controls and KYP when switched off', () => {
      const request = toBuildingSaveRequest(
        validForm({ edControls: { ...ed, isTicketSource: false }, kyp: { ...kyp, isPlanningSource: false } }),
        CONTEXT,
        PRESERVED
      );

      expect(request.ticketSourceId).toBe(0);
      expect(request.ticketProjectId).toBeUndefined();
      expect(request.tagData).toBeUndefined();
      expect(request.planningSourcetId).toBe(1);
      expect(request.planningProjectId).toBeUndefined();
    });

    it('keeps the stored alternative token when none was typed, and replaces it when one was', () => {
      const keep = toKypSaveFields({ ...kyp, isUseProjectToken: true, hasStoredProjectToken: true }, 'stored-token');
      const replace = toKypSaveFields({ ...kyp, isUseProjectToken: true, projectToken: ' new ' }, 'stored-token');

      expect(keep.kypProjectAuthToken).toBe('stored-token');
      expect(replace.kypProjectAuthToken).toBe('new');
    });
  });

  describe('validation', () => {
    it('requires the ED Controls and KYP projects while their toggles are on', () => {
      const errors = validateBuildingForm(
        validForm({
          edControls: { isTicketSource: true, ticketProjectId: '', tagRows: [] },
          kyp: { isPlanningSource: true, planningProjectId: '', isUseProjectToken: false, projectToken: '', hasStoredProjectToken: false }
        }),
        { domainGroup: 'taskXs', isGrandChildFolderEnabled: false }
      );

      expect(errors.ticketProjectId).toBe(INTEGRATION_MESSAGES.ticketProjectRequired);
      expect(errors.planningProjectId).toBe(INTEGRATION_MESSAGES.planningProjectRequired);
      expect(errors.tagRows).toBeUndefined();
    });

    it('reports a tag row without a folder', () => {
      const incomplete: ITagRow = row({ tags: ['a'] });
      const errors = validateBuildingForm(
        validForm({ edControls: { isTicketSource: true, ticketProjectId: 'db', tagRows: [incomplete] } }),
        { domainGroup: 'taskXs', isGrandChildFolderEnabled: false }
      );

      expect(errors.tagRows).toEqual({ [incomplete.key]: { folderId: INTEGRATION_MESSAGES.folderRequired } });
    });

    it('ignores the sections while they are switched off', () => {
      const errors = validateBuildingForm(
        validForm({ edControls: { isTicketSource: false, ticketProjectId: '', tagRows: [row({ tags: ['a'] })] } }),
        { domainGroup: 'taskXs', isGrandChildFolderEnabled: false }
      );

      expect(errors).toEqual({});
    });
  });
});
