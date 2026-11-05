/* eslint-disable @rushstack/no-new-null -- The settings DTOs describe the BMDesk Web API's
   wire format, which serializes with System.Text.Json defaults: absent values arrive as an
   explicit `null`. */

/**
 * The project form's third-party integrations: ED Controls (tickets / audits, tag rows) and
 * KYP (planning).
 *
 * Every call to ED Controls or KYP is made by the Web API (`api/EDControls`, `api/KYP`), so
 * no credential or token is ever read here - the API answers lists and "is it configured".
 * What the user chooses is saved with the project through `POST api/Building`, in the
 * fields the reference's `Views/Building/AddEdit.cshtml` posts: `TicketSourceId`,
 * `TicketProjectId`, `TagData`, `PlanningSourcetId`, `PlanningProjectId`,
 * `KypProjectAuthToken`.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package.
 */

import type { IFolderSelection, ILookupOption } from './Building';

/** `TicketSourceTypeEnum`: what `TicketSourceId` is set to by the ED Controls toggle (0 = none). */
export const TICKET_SOURCE = { none: 0, edControls: 1 } as const;

/** `PlanningSourceTypeEnum`: the KYP toggle switches between these, exactly as the reference does. */
export const PLANNING_SOURCE = { msProject: 1, kyp: 2 } as const;

/** Validation messages, the reference's own resource strings. */
export const INTEGRATION_MESSAGES = {
  ticketProjectRequired: 'Please select a project',
  planningProjectRequired: 'Please select a project',
  tagRequired: 'Please select at least one tag',
  folderRequired: 'Please select a folder'
} as const;

/**
 * One ED Controls tag row: audits carrying these tags are filed into this folder path (and
 * supplier). `periodic` reveals `documentName`, the name the recurring document gets.
 */
export interface ITagRow {
  /** Client-only identity, so React and the per-row folder options survive add/remove. */
  key: string;
  periodic: boolean;
  documentName: string;
  tags: string[];
  folder: IFolderSelection;
}

/** The ED Controls part of the project form. Absent on the form = not loaded: the save keeps what is stored. */
export interface IEdControlsForm {
  /** The "ED controls" toggle. */
  isTicketSource: boolean;
  /** The ED Controls project ("database") the project is linked to. */
  ticketProjectId: string;
  /** Edited in edit mode only: folders exist once the project does. */
  tagRows: ITagRow[];
}

/** The KYP part of the project form. Absent on the form = not loaded: the save keeps what is stored. */
export interface IKypForm {
  /** The "KYP Project" toggle. */
  isPlanningSource: boolean;
  planningProjectId: string;
  /** "Use alternative token": the project's own KYP token instead of the account's. */
  isUseProjectToken: boolean;
  /** A token typed in this session. Empty keeps the stored one. Never read back from the API. */
  projectToken: string;
  /** The project already stores an alternative token (its value is never sent to the browser). */
  hasStoredProjectToken: boolean;
}

/** Per-row validation messages. */
export interface ITagRowErrors {
  tags?: string;
  folderId?: string;
}

/** `GET api/EDControls/GetProjectSettings`. */
export interface IEdControlsSettings {
  isConfigured: boolean;
  /** The project's name; empty while adding. */
  projectName?: string;
  isTicketSource: boolean;
  ticketProjectId: string;
  tagRows: ITagRow[];
}

/** `GET api/KYP/GetProjectSettings`. */
export interface IKypSettings {
  isConfigured: boolean;
  isPlanningSource: boolean;
  planningProjectId: string;
  hasProjectToken: boolean;
}

/** `POST api/KYP/GetProjects` body. */
export interface IKypProjectsRequest {
  projectId: number;
  useProjectToken: boolean;
  authToken: string;
}

/** The fields of `POST api/Building` these sections own. */
export interface IIntegrationSaveFields {
  ticketSourceId?: number;
  ticketProjectId?: string;
  tagData?: string;
  planningSourcetId?: number;
  planningProjectId?: string;
  kypProjectAuthToken?: string;
}

/** One stored tag row as the API returns it (`TagRowData`). */
interface IApiTagRow {
  periodic?: boolean | null;
  documentName?: string | null;
  tags?: (string | null)[] | null;
  folderId?: string | null;
  subFolderId?: string | null;
  subSubFolderId?: string | null;
  subSubSubFolderId?: string | null;
  supplierId?: string | null;
}

let tagRowSequence: number = 0;

/** A fresh, empty tag row. */
export function createTagRow(): ITagRow {
  tagRowSequence += 1;

  return {
    key: `tag-row-${tagRowSequence}`,
    periodic: false,
    documentName: '',
    tags: [],
    folder: { folderId: '', subFolderId: '', subSubFolderId: '', subSubSubFolderId: '', supplierId: '' }
  };
}

function text(value: unknown): string {
  return typeof value === 'string' ? value.trim() : typeof value === 'number' ? String(value) : '';
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

/** Maps one stored row; tags are de-duplicated and blanks dropped. */
function toTagRow(entry: unknown): ITagRow | undefined {
  if (!isObject(entry)) {
    return undefined;
  }

  const item: IApiTagRow = entry as IApiTagRow;
  const tags: string[] = [];
  (Array.isArray(item.tags) ? item.tags : []).forEach((tag: unknown): void => {
    const value: string = text(tag);
    if (value && tags.indexOf(value) < 0) {
      tags.push(value);
    }
  });

  return {
    ...createTagRow(),
    periodic: item.periodic === true,
    documentName: text(item.documentName),
    tags,
    folder: {
      folderId: text(item.folderId),
      subFolderId: text(item.subFolderId),
      subSubFolderId: text(item.subSubFolderId),
      subSubSubFolderId: text(item.subSubSubFolderId),
      supplierId: text(item.supplierId)
    }
  };
}

/** Reads `GetProjectSettings`; `undefined` when the payload is not the promised object. */
export function toEdControlsSettings(payload: unknown): IEdControlsSettings | undefined {
  if (!isObject(payload)) {
    return undefined;
  }

  const rows: ITagRow[] = [];
  (Array.isArray(payload.tagRows) ? payload.tagRows : []).forEach((entry: unknown): void => {
    const row: ITagRow | undefined = toTagRow(entry);
    if (row) {
      rows.push(row);
    }
  });

  return {
    isConfigured: payload.isConfigured === true,
    projectName: text(payload.projectName),
    isTicketSource: payload.isTicketSource === true,
    ticketProjectId: text(payload.ticketProjectId),
    tagRows: rows
  };
}

/** Reads `GetProjectSettings`; `undefined` when the payload is not the promised object. */
export function toKypSettings(payload: unknown): IKypSettings | undefined {
  if (!isObject(payload)) {
    return undefined;
  }

  return {
    isConfigured: payload.isConfigured === true,
    isPlanningSource: payload.isPlanningSource === true,
    planningProjectId: text(payload.planningProjectId),
    hasProjectToken: payload.hasProjectToken === true
  };
}

/** Reads the tag list: a JSON array of strings, blanks and duplicates dropped. */
export function toTagList(payload: unknown): string[] {
  const tags: string[] = [];
  (Array.isArray(payload) ? payload : []).forEach((tag: unknown): void => {
    const value: string = text(tag);
    if (value && tags.indexOf(value) < 0) {
      tags.push(value);
    }
  });

  return tags;
}

/**
 * The form section's starting value. In edit mode an empty row is offered when none is stored,
 * like the reference's blank first row.
 */
export function toEdControlsForm(settings: IEdControlsSettings, isEditMode: boolean): IEdControlsForm {
  return {
    isTicketSource: settings.isTicketSource,
    ticketProjectId: settings.ticketProjectId,
    tagRows: settings.tagRows.length > 0 || !isEditMode ? settings.tagRows : [createTagRow()]
  };
}

export function toKypForm(settings: IKypSettings): IKypForm {
  return {
    isPlanningSource: settings.isPlanningSource,
    planningProjectId: settings.planningProjectId,
    isUseProjectToken: settings.hasProjectToken,
    projectToken: '',
    hasStoredProjectToken: settings.hasProjectToken
  };
}

/** Tag options for a multi-select: the project's tags plus any a row already holds (a renamed or removed tag stays visible). */
export function toTagOptions(tags: string[], selected: string[]): ILookupOption[] {
  const all: string[] = tags.slice();
  selected.forEach((tag: string): void => {
    if (all.indexOf(tag) < 0) {
      all.push(tag);
    }
  });

  return all.map((tag: string): ILookupOption => ({ id: tag, name: tag }));
}

/**
 * Tag row messages. Saving checks only rows that carry tags (rows without tags are not
 * saved, as in the reference's `getAllTagData`); adding a row (`requireTags`) needs every
 * existing row complete, as its `validateAllRowsOnEdit` does. Empty when all is well.
 */
export function validateTagRows(rows: ITagRow[], requireTags: boolean): { [key: string]: ITagRowErrors } {
  const errors: { [key: string]: ITagRowErrors } = {};

  rows.forEach((row: ITagRow): void => {
    const rowErrors: ITagRowErrors = {};
    if (requireTags && row.tags.length === 0) {
      rowErrors.tags = INTEGRATION_MESSAGES.tagRequired;
    }
    if ((requireTags || row.tags.length > 0) && !row.folder.folderId) {
      rowErrors.folderId = INTEGRATION_MESSAGES.folderRequired;
    }
    if (rowErrors.tags || rowErrors.folderId) {
      errors[row.key] = rowErrors;
    }
  });

  return errors;
}

/** `TagData`: the rows that carry tags, in the reference's JSON shape; `undefined` when none. */
export function toTagData(rows: ITagRow[]): string | undefined {
  const saved = rows
    .filter((row: ITagRow) => row.tags.length > 0)
    .map((row: ITagRow) => ({
      periodic: row.periodic,
      documentName: row.periodic ? row.documentName.trim() : '',
      tags: row.tags.slice(),
      folderId: row.folder.folderId,
      subFolderId: row.folder.subFolderId,
      subSubFolderId: row.folder.subSubFolderId,
      subSubSubFolderId: row.folder.subSubSubFolderId,
      supplierId: row.folder.supplierId
    }));

  return saved.length > 0 ? JSON.stringify(saved) : undefined;
}

/**
 * The ED Controls fields of the save request. `{}` when the section was not loaded, so the
 * values read back from the project are echoed unchanged. Switched off, the source is `0` and
 * the project and rows are dropped (the server stores `TagData` only for ED Controls anyway).
 */
export function toEdControlsSaveFields(form: IEdControlsForm | undefined): IIntegrationSaveFields {
  if (!form) {
    return {};
  }
  if (!form.isTicketSource) {
    return { ticketSourceId: TICKET_SOURCE.none, ticketProjectId: undefined, tagData: undefined };
  }

  return {
    ticketSourceId: TICKET_SOURCE.edControls,
    ticketProjectId: form.ticketProjectId || undefined,
    tagData: toTagData(form.tagRows)
  };
}

/**
 * The KYP fields of the save request. `{}` when the section was not loaded. The alternative
 * token is cleared when "Use alternative token" is off (as the reference's `AddEdit` POST
 * does), replaced when one was typed, and otherwise the stored one (`storedToken`, read back
 * by the save) is kept.
 */
export function toKypSaveFields(form: IKypForm | undefined, storedToken: string | undefined): IIntegrationSaveFields {
  if (!form) {
    return {};
  }

  const typed: string = form.projectToken.trim();

  return {
    planningSourcetId: form.isPlanningSource ? PLANNING_SOURCE.kyp : PLANNING_SOURCE.msProject,
    planningProjectId: form.isPlanningSource ? form.planningProjectId || undefined : undefined,
    kypProjectAuthToken: !form.isUseProjectToken ? '' : typed || storedToken || undefined
  };
}
