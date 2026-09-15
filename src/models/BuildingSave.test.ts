import { IBuildingForm, createEmptyBuildingForm } from './Building';
import {
  DOCUMENT_STORAGE_TYPE,
  IBuildingDetails,
  IBuildingSaveRequest,
  IPreservedProjectFields,
  IProjectSaveContext,
  toBuildingSaveRequest,
  toHostName,
  toImageBytes,
  toMomData,
  toPreservedFields,
  toRebindRequest
} from './BuildingSave';

/** A form filled in the way `validateBuildingForm` would let through. */
function form(overrides?: Partial<IBuildingForm>): IBuildingForm {
  return {
    ...createEmptyBuildingForm(),
    buildingName: 'Roof replacement',
    address: 'Keizersgracht 1',
    postcode: '1015 CJ',
    clientId: '42',
    ...overrides
  };
}

function context(overrides?: Partial<IProjectSaveContext>): IProjectSaveContext {
  return {
    clientId: 42,
    clientName: 'Mayur',
    accountId: 7,
    domainName: 'contoso.sharepoint.com',
    companyName: 'Contoso',
    companyLogo: '',
    ...overrides
  };
}

describe('toBuildingSaveRequest - the renamed fields', () => {
  it('sends the project name as projectName', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form(), context());

    expect(request.projectName).toBe('Roof replacement');
  });

  it('sends the human-facing project code as projectId, not the primary key', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12, buldingId: 'PRJ-001' }),
      context()
    );

    expect(request.projectId).toBe('PRJ-001');
    expect(request.id).toBe(12);
  });

  it('keeps id 0 for an insert, which is what routes the API to Add', () => {
    expect(toBuildingSaveRequest(form(), context()).id).toBe(0);
  });
});

describe('toBuildingSaveRequest - text', () => {
  it('trims text, so a whitespace-only value is not stored as a name', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ buildingName: '  Roof  ', email: ' a@b.com ' }),
      context()
    );

    expect(request.projectName).toBe('Roof');
    expect(request.email).toBe('a@b.com');
  });

  it('always sends isActive true - this path never deactivates a project', () => {
    expect(toBuildingSaveRequest(form(), context()).isActive).toBe(true);
  });
});

describe('toBuildingSaveRequest - ids', () => {
  it('parses the multi-selects into number arrays', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ userIds: ['3', '9'], spatialBreakdownIds: ['5'] }),
      context()
    );

    expect(request.userId).toEqual([3, 9]);
    expect(request.spatialBreakdownId).toEqual([5]);
  });

  it('sends empty arrays rather than omitting them - the API dereferences both', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form(), context());

    expect(request.userId).toEqual([]);
    expect(request.spatialBreakdownId).toEqual([]);
  });

  it('drops unparseable option values instead of sending NaN', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ userIds: ['3', '', 'all'] }),
      context()
    );

    expect(request.userId).toEqual([3]);
  });

  it('omits projectTemplateId when none is chosen, so the API picks the base template', () => {
    expect(toBuildingSaveRequest(form(), context()).projectTemplateId).toBeUndefined();
  });

  it('sends the chosen project template', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ projectTemplateId: '4' }),
      context()
    );

    expect(request.projectTemplateId).toBe(4);
  });

  it('falls back to the signed-in user when the form carries no client id', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form({ clientId: '' }), context());

    expect(request.clientId).toBe(42);
  });
});

describe('toBuildingSaveRequest - tender', () => {
  it('sends the tender template while the toggle is on', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ isEnableTender: true, tenderTemplateId: '8' }),
      context()
    );

    expect(request.tenderTemplateId).toBe(8);
  });

  it('does not carry a stale template along with the toggle switched off', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ isEnableTender: false, tenderTemplateId: '8' }),
      context()
    );

    expect(request.tenderTemplateId).toBe(0);
  });
});

/** A project read back through `GET api/Building?id=`, as an update would see it. */
function details(overrides?: Partial<IBuildingDetails>): IBuildingDetails {
  return {
    id: 12,
    documentStorageType: DOCUMENT_STORAGE_TYPE.xsCloud,
    sharepointFolderId: 'drive-item-id',
    imageBytes: 'c3RvcmVk',
    tagData: '[{"tags":["a"]}]',
    ticketSourceId: 2,
    ticketProjectId: 'ED-1',
    planningSourcetId: 3,
    planningProjectId: 'KYP-1',
    isDocumentReviewerEnable: true,
    kypProjectAuthToken: 'kyp-token',
    ...overrides
  };
}

describe('toPreservedFields', () => {
  it('carries every field Update would otherwise null', () => {
    const preserved: IPreservedProjectFields = toPreservedFields(details());

    expect(preserved).toEqual({
      documentStorageType: DOCUMENT_STORAGE_TYPE.xsCloud,
      sharepointFolderId: 'drive-item-id',
      imageBytes: 'c3RvcmVk',
      tagData: '[{"tags":["a"]}]',
      ticketSourceId: 2,
      ticketProjectId: 'ED-1',
      planningSourcetId: 3,
      planningProjectId: 'KYP-1',
      isDocumentReviewerEnable: true,
      kypProjectAuthToken: 'kyp-token'
    });
  });
});

describe('toBuildingSaveRequest - an update echoes what it must not lose', () => {
  it('keeps the SharePoint folder id, which would otherwise orphan the folder tree', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12 }),
      context(),
      toPreservedFields(details())
    );

    expect(request.sharepointFolderId).toBe('drive-item-id');
  });

  it('keeps the stored image when no new one was chosen', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12 }),
      context(),
      toPreservedFields(details())
    );

    expect(request.imageBytes).toBe('c3RvcmVk');
  });

  it('keeps the tag rows together with the ticket source they depend on', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12 }),
      context(),
      toPreservedFields(details())
    );

    expect(request.tagData).toBe('[{"tags":["a"]}]');
    expect(request.ticketSourceId).toBe(2);
  });

  it('keeps the KYP fields', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12 }),
      context(),
      toPreservedFields(details())
    );

    expect(request.planningSourcetId).toBe(3);
    expect(request.planningProjectId).toBe('KYP-1');
    expect(request.kypProjectAuthToken).toBe('kyp-token');
    expect(request.isDocumentReviewerEnable).toBe(true);
  });

  it('takes the storage type from the project, never from the form', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12 }),
      context(),
      toPreservedFields(details())
    );

    // An XS Cloud project stays XS Cloud, even though an insert would say SharePoint.
    expect(request.documentStorageType).toBe(DOCUMENT_STORAGE_TYPE.xsCloud);
  });

  it('leaves a genuinely empty field absent rather than inventing a value', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12 }),
      context(),
      toPreservedFields(details({ sharepointFolderId: undefined, imageBytes: undefined }))
    );

    expect(request.sharepointFolderId).toBeUndefined();
    expect(request.imageBytes).toBeUndefined();
  });

  it('prefers a newly uploaded image over the stored one', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ id: 12, imageDataUrl: 'data:image/png;base64,dXBsb2FkZWQ=' }),
      context(),
      toPreservedFields(details())
    );

    expect(request.imageBytes).toBe('dXBsb2FkZWQ=');
  });
});

describe('toBuildingSaveRequest - storage and session values', () => {
  it('always stores a new project in SharePoint - a SPFx bundle has no XS Cloud fallback', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form(), context());

    expect(request.documentStorageType).toBe(DOCUMENT_STORAGE_TYPE.sharePoint);
  });

  it('sends an uploaded image on an insert', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(
      form({ imageDataUrl: 'data:image/jpeg;base64,dXBsb2FkZWQ=' }),
      context()
    );

    expect(request.imageBytes).toBe('dXBsb2FkZWQ=');
  });

  it('omits the integration fields on an insert - there is nothing to preserve yet', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form(), context());

    expect(request.sharepointFolderId).toBeUndefined();
    expect(request.tagData).toBeUndefined();
    expect(request.imageBytes).toBeUndefined();
  });

  it('passes the session-derived values straight through', () => {
    const request: IBuildingSaveRequest = toBuildingSaveRequest(form(), context());

    expect(request.accountId).toBe(7);
    expect(request.domainName).toBe('contoso.sharepoint.com');
    expect(request.companyName).toBe('Contoso');
  });
});

describe('toMomData', () => {
  it('omits the folder data while the toggle is off', () => {
    expect(toMomData(form({ isAddMomPdf: false }))).toBeUndefined();
  });

  it('serializes the single folder path as a one-row array', () => {
    const momData: string | undefined = toMomData(
      form({
        isAddMomPdf: true,
        momFolder: {
          folderId: '1',
          subFolderId: '2',
          subSubFolderId: '3',
          subSubSubFolderId: '4',
          supplierId: '5'
        }
      })
    );

    expect(JSON.parse(momData as string)).toEqual([
      {
        folderId: '1',
        subFolderId: '2',
        subSubFolderId: '3',
        subSubSubFolderId: '4',
        supplierId: '5'
      }
    ]);
  });
});

describe('toImageBytes', () => {
  it('takes the Base64 payload out of a data URL', () => {
    expect(toImageBytes('data:image/png;base64,aGVsbG8=')).toBe('aGVsbG8=');
  });

  it('ignores a blob URL, which carries no bytes of its own', () => {
    expect(toImageBytes('blob:https://contoso.sharepoint.com/8f2c-4a1b')).toBeUndefined();
  });

  it('ignores a data URL that is not Base64-encoded', () => {
    expect(toImageBytes('data:text/plain,hello')).toBeUndefined();
  });

  it('ignores a data URL with an empty payload', () => {
    expect(toImageBytes('data:image/png;base64,')).toBeUndefined();
  });

  it('ignores an absent image', () => {
    expect(toImageBytes(undefined)).toBeUndefined();
  });
});

describe('toHostName', () => {
  it('takes the host out of a site URL', () => {
    expect(toHostName('https://contoso.sharepoint.com/sites/xs')).toBe('contoso.sharepoint.com');
  });

  it('keeps the port, as Request.Host.Value does', () => {
    expect(toHostName('https://localhost:4321/sites/xs')).toBe('localhost:4321');
  });

  it('passes a bare host name through', () => {
    expect(toHostName('contoso.sharepoint.com')).toBe('contoso.sharepoint.com');
  });

  it('returns an empty string for an empty value, so the caller can fall back', () => {
    expect(toHostName('')).toBe('');
  });
});

describe('toRebindRequest', () => {
  /** The body an insert sent: no preserved fields, and `id: 0` until the API assigns one. */
  function inserted(): IBuildingSaveRequest {
    return toBuildingSaveRequest(
      form({ imageDataUrl: 'data:image/png;base64,dXBsb2FkZWQ=' }),
      context()
    );
  }

  it('binds the id the insert was assigned, which turns the second post into an update', () => {
    const request: IBuildingSaveRequest = toRebindRequest(inserted(), 18, 'drive-item-id');

    expect(request.id).toBe(18);
  });

  it('binds the root folder id the project could not be given before it existed', () => {
    const request: IBuildingSaveRequest = toRebindRequest(inserted(), 18, 'drive-item-id');

    expect(request.sharepointFolderId).toBe('drive-item-id');
  });

  it('re-sends the whole body, because Update reads an omitted field as null', () => {
    // The one rule that makes this a copy of the insert rather than a two-field patch:
    // `ProjectController.Update` assigns the image and the storage type unconditionally,
    // so a slim body would blank what the insert had just stored.
    const request: IBuildingSaveRequest = toRebindRequest(inserted(), 18, 'drive-item-id');

    expect(request.imageBytes).toBe('dXBsb2FkZWQ=');
    expect(request.documentStorageType).toBe(DOCUMENT_STORAGE_TYPE.sharePoint);
    expect(request.projectName).toBe('Roof replacement');
    expect(request.clientId).toBe(42);
  });

  it('leaves the body it was handed untouched, so the printed insert stays what was sent', () => {
    const first: IBuildingSaveRequest = inserted();
    toRebindRequest(first, 18, 'drive-item-id');

    expect(first.id).toBe(0);
    expect(first.sharepointFolderId).toBeUndefined();
  });
});
