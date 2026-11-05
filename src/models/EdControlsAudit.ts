/**
 * The project Audits page: audits created from an ED Controls template in the project's
 * linked ED Controls project, whose completed report is filed into a project folder. Ported
 * from the reference's `Views/Building/_AddEdControlAudit.cshtml`.
 *
 * The audits themselves live in `api/EDControlAudit` (list, create, bulk create, delete);
 * the template list comes from `api/EDControls/GetAuditTemplates`, which calls ED Controls
 * server-side; a supplier's contacts from `api/Contact/GetAllContact`.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package.
 */

import type { IFolderSelection, ILookupOption } from './Building';

/** `EDControlAudit.Status`: waiting for the completed report, or the report has been filed. */
export const AUDIT_STATUS = { completed: 2 } as const;

/** The reference's limits: a duplicate batch, and an audit name. */
export const MAX_DUPLICATES: number = 20;
export const MAX_AUDIT_NAME_LENGTH: number = 250;

/** The reference's own resource strings. */
export const AUDIT_MESSAGES = {
  nameRequired: 'Audit name is required.',
  templateRequired: 'Ed Controls template is required.',
  supplierRequired: 'Supplier is required.',
  dueDateRequired: 'Due date is required.',
  dueDateInvalid: 'Enter a valid date (dd/mm/yyyy).',
  folderRequired: 'Folder is required.',
  countInvalid: `Enter a whole number between 1 and ${MAX_DUPLICATES}.`
} as const;

/** One configured audit, as `GET EDControlAudit/GetByProject` lists it. */
export interface IAudit {
  id: number;
  buildingId: number;
  auditName: string;
  edControlsDatabase: string;
  edControlsTemplateId: string;
  edControlsTemplateName: string;
  supplierId?: number;
  supplierName: string;
  contactId?: number;
  contactName: string;
  /** As the API sends it (ISO); see {@link formatDueDate}. */
  dueDate: string;
  folderId: number;
  subFolderId?: number;
  childSubFolderId?: number;
  grandChildSubFolderId?: number;
  folderPath: string;
  status: number;
}

/**
 * The add / edit-duplicate form. Ids are strings, as `<select>` values always are; the names
 * travel with them so a draft row can show what was chosen without another lookup.
 */
export interface IAuditForm {
  auditName: string;
  templateId: string;
  templateName: string;
  supplierId: string;
  supplierName: string;
  contactId: string;
  contactName: string;
  /** As typed: dd/mm/yyyy. */
  dueDate: string;
  /** The folder path; its `supplierId` is unused (the audit's supplier is {@link supplierId}). */
  folder: IFolderSelection;
  folderPath: string;
}

export interface IAuditFormErrors {
  auditName?: string;
  templateId?: string;
  supplierId?: string;
  dueDate?: string;
  folderId?: string;
}

/** One draft of a duplicate batch: nothing is created in ED Controls until the batch is saved. */
export interface IAuditDraft {
  key: string;
  form: IAuditForm;
  /** Why the save of this row failed, from the bulk result. */
  error?: string;
}

/** `POST EDControlAudit/post` body, and one row of `postBulk`. */
export interface IAuditSaveRequest {
  buildingId: number;
  auditName: string;
  edControlsDatabase: string;
  edControlsTemplateId: string;
  edControlsTemplateName: string;
  supplierId?: number;
  contactId?: number;
  /** yyyy-MM-dd, so the API binds the same date whatever its culture. */
  dueDate?: string;
  folderId: number;
  subFolderId?: number;
  childSubFolderId?: number;
  grandChildSubFolderId?: number;
}

/** One row's outcome of `postBulk`. */
export interface IAuditSaveResult {
  index: number;
  auditName: string;
  success: boolean;
  message: string;
  id: number;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function text(value: unknown): string {
  return typeof value === 'string' ? value.trim() : typeof value === 'number' ? String(value) : '';
}

function optionalId(value: unknown): number | undefined {
  return typeof value === 'number' && value > 0 ? value : undefined;
}

/** Reads the audit list; entries that are not audits are skipped. */
export function toAudits(payload: unknown): IAudit[] {
  const audits: IAudit[] = [];

  (Array.isArray(payload) ? payload : []).forEach((entry: unknown): void => {
    if (!isObject(entry) || typeof entry.id !== 'number') {
      return;
    }
    audits.push({
      id: entry.id,
      buildingId: typeof entry.buildingId === 'number' ? entry.buildingId : 0,
      auditName: text(entry.auditName),
      edControlsDatabase: text(entry.edControlsDatabase),
      edControlsTemplateId: text(entry.edControlsTemplateId),
      edControlsTemplateName: text(entry.edControlsTemplateName),
      supplierId: optionalId(entry.supplierId),
      supplierName: text(entry.supplierName),
      contactId: optionalId(entry.contactId),
      contactName: text(entry.contactName),
      dueDate: text(entry.dueDate),
      folderId: typeof entry.folderId === 'number' ? entry.folderId : 0,
      subFolderId: optionalId(entry.subFolderId),
      childSubFolderId: optionalId(entry.childSubFolderId),
      grandChildSubFolderId: optionalId(entry.grandChildSubFolderId),
      folderPath: text(entry.folderPath),
      status: typeof entry.status === 'number' ? entry.status : 0
    });
  });

  return audits;
}

/** Template options (`{ id, name }`, the id an ED Controls string). */
export function toTemplateOptions(payload: unknown): ILookupOption[] {
  const options: ILookupOption[] = [];
  (Array.isArray(payload) ? payload : []).forEach((entry: unknown): void => {
    if (isObject(entry) && text(entry.id)) {
      options.push({ id: text(entry.id), name: text(entry.name) || text(entry.id) });
    }
  });

  return options;
}

/**
 * Contact options of one supplier, from `Contact/GetAllContact`: the name, else the e-mail
 * (a contact's name is optional), without blanks, by name - as the reference's
 * `GetSupplierContactList` builds them.
 */
export function toContactOptions(payload: unknown): ILookupOption[] {
  const options: ILookupOption[] = [];
  (Array.isArray(payload) ? payload : []).forEach((entry: unknown): void => {
    if (!isObject(entry) || typeof entry.id !== 'number') {
      return;
    }
    const name: string = text(entry.name) || text(entry.email);
    if (name) {
      options.push({ id: String(entry.id), name });
    }
  });

  return options.sort((a: ILookupOption, b: ILookupOption) => a.name.localeCompare(b.name));
}

/** Reads the bulk results; anything unreadable is dropped. */
export function toSaveResults(payload: unknown): IAuditSaveResult[] {
  const results: IAuditSaveResult[] = [];
  (Array.isArray(payload) ? payload : []).forEach((entry: unknown): void => {
    if (isObject(entry) && typeof entry.index === 'number') {
      results.push({
        index: entry.index,
        auditName: text(entry.auditName),
        success: entry.success === true,
        message: text(entry.message),
        id: typeof entry.id === 'number' ? entry.id : 0
      });
    }
  });

  return results;
}

function pad(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/**
 * A typed dd/mm/yyyy date, strictly: a real calendar date or `undefined`. Built in local
 * time and read back, so 31/02/2026 does not roll over into March.
 */
export function parseDueDate(value: string): Date | undefined {
  const match: RegExpExecArray | null = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec((value || '').trim());
  if (!match) {
    return undefined;
  }

  const day: number = Number(match[1]);
  const month: number = Number(match[2]);
  const year: number = Number(match[3]);
  const date: Date = new Date(year, month - 1, day);

  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : undefined;
}

/** yyyy-MM-dd of a typed dd/mm/yyyy date; `undefined` when it is not one. */
export function toIsoDate(value: string): string | undefined {
  const date: Date | undefined = parseDueDate(value);

  return date ? `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` : undefined;
}

/**
 * dd/mm/yyyy of a date the API sent. The date part is read verbatim from an ISO string, so
 * a date-only value is never shifted a day by the browser's time zone.
 */
export function formatDueDate(value: string | undefined): string {
  const match: RegExpExecArray | null = /^(\d{4})-(\d{2})-(\d{2})/.exec((value || '').trim());

  return match ? `${match[3]}/${match[2]}/${match[1]}` : '';
}

/** Messages for the form, the reference's `validateAudit`; empty when it can be saved. */
export function validateAuditForm(form: IAuditForm): IAuditFormErrors {
  const errors: IAuditFormErrors = {};

  if (!form.auditName.trim()) {
    errors.auditName = AUDIT_MESSAGES.nameRequired;
  }
  if (!form.templateId) {
    errors.templateId = AUDIT_MESSAGES.templateRequired;
  }
  if (!form.supplierId) {
    errors.supplierId = AUDIT_MESSAGES.supplierRequired;
  }
  if (!form.dueDate.trim()) {
    errors.dueDate = AUDIT_MESSAGES.dueDateRequired;
  } else if (!parseDueDate(form.dueDate)) {
    errors.dueDate = AUDIT_MESSAGES.dueDateInvalid;
  }
  if (!form.folder.folderId) {
    errors.folderId = AUDIT_MESSAGES.folderRequired;
  }

  return errors;
}

/** An empty form. */
export function createAuditForm(): IAuditForm {
  return {
    auditName: '',
    templateId: '',
    templateName: '',
    supplierId: '',
    supplierName: '',
    contactId: '',
    contactName: '',
    dueDate: '',
    folder: { folderId: '', subFolderId: '', subSubFolderId: '', subSubSubFolderId: '', supplierId: '' },
    folderPath: ''
  };
}

function toId(value: string): number | undefined {
  const id: number = Number(value);

  return isFinite(id) && id > 0 ? Math.floor(id) : undefined;
}

/**
 * The save body. The ED Controls project is the project's saved link (`database`), as on the
 * reference's standalone Audits page; empty folder levels and an unchosen contact stay absent.
 */
export function toAuditSaveRequest(form: IAuditForm, buildingId: number, database: string): IAuditSaveRequest {
  return {
    buildingId,
    auditName: form.auditName.trim().substring(0, MAX_AUDIT_NAME_LENGTH),
    edControlsDatabase: database,
    edControlsTemplateId: form.templateId,
    edControlsTemplateName: form.templateName,
    supplierId: toId(form.supplierId),
    contactId: toId(form.contactId),
    dueDate: toIsoDate(form.dueDate),
    folderId: toId(form.folder.folderId) || 0,
    subFolderId: toId(form.folder.subFolderId),
    childSubFolderId: toId(form.folder.subSubFolderId),
    grandChildSubFolderId: toId(form.folder.subSubSubFolderId)
  };
}

/** The form an existing audit would fill in - the source of its duplicates. */
export function toAuditForm(audit: IAudit): IAuditForm {
  const id = (value: number | undefined): string => (value ? String(value) : '');

  return {
    auditName: audit.auditName,
    templateId: audit.edControlsTemplateId,
    templateName: audit.edControlsTemplateName,
    supplierId: id(audit.supplierId),
    supplierName: audit.supplierName,
    contactId: id(audit.contactId),
    contactName: audit.contactName,
    dueDate: formatDueDate(audit.dueDate),
    folder: {
      folderId: id(audit.folderId),
      subFolderId: id(audit.subFolderId),
      subSubFolderId: id(audit.childSubFolderId),
      subSubSubFolderId: id(audit.grandChildSubFolderId),
      supplierId: ''
    },
    folderPath: audit.folderPath
  };
}

/** A typed duplicate count: a whole number from 1 to {@link MAX_DUPLICATES}, else `undefined`. */
export function parseDuplicateCount(value: string): number | undefined {
  const trimmed: string = (value || '').trim();

  if (!/^\d+$/.test(trimmed)) {
    return undefined;
  }
  const count: number = Number(trimmed);

  return count >= 1 && count <= MAX_DUPLICATES ? count : undefined;
}

let draftSequence: number = 0;

/** `count` drafts copying `audit`, named "<name> - 1", "<name> - 2", ... (clamped to the name limit). */
export function createDuplicateDrafts(audit: IAudit, count: number): IAuditDraft[] {
  const drafts: IAuditDraft[] = [];

  for (let index: number = 1; index <= count; index++) {
    draftSequence += 1;
    drafts.push({
      key: `audit-draft-${draftSequence}`,
      form: { ...toAuditForm(audit), auditName: `${audit.auditName} - ${index}`.substring(0, MAX_AUDIT_NAME_LENGTH) }
    });
  }

  return drafts;
}
