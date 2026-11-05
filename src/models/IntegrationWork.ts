/**
 * The Tickets & Planning page: the signed-in user's open ED Controls tickets and the KYP
 * planning items they are responsible for, grouped per project - the reference's "My tasks"
 * overview (`Views/Task/Index.cshtml`, its Tickets and Planning tabs).
 *
 * Read from `api/EDControls/GetMyTickets` and `api/KYP/GetMyPlannings`; the Web API calls ED
 * Controls and KYP itself. Ticket photos are not offered: ED Controls only serves them with
 * its access token in the URL, which must not reach the browser - each ticket links to ED
 * Controls instead.
 *
 * Part of the models layer: no imports from `config`, `services`, `components` or any SPFx
 * package.
 */

export interface ITicket {
  id: string;
  ticketNumber: string;
  title: string;
  /** Plain text - rendered as text, never as HTML (the reference's `Html.Raw` is not ported). */
  body: string;
  /** ED Controls' state, e.g. `created` or `started`. */
  state: string;
  createdOn: string;
  dueDate: string;
  url: string;
  attachmentCount: number;
}

export interface ITicketGroup {
  projectId: number;
  projectName: string;
  tickets: ITicket[];
  /** Why this project's tickets could not be read; empty when they could. */
  message: string;
}

export interface IPlanningItem {
  name: string;
  startDate: string;
  endDate: string;
  totalDays: number;
  elapsedDays: number;
}

export interface IPlanningGroup {
  projectId: number;
  projectName: string;
  /** The project's planning in KYP. */
  url: string;
  items: IPlanningItem[];
  message: string;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function text(value: unknown): string {
  return typeof value === 'string' ? value.trim() : typeof value === 'number' ? String(value) : '';
}

function count(value: unknown): number {
  return typeof value === 'number' && isFinite(value) ? value : 0;
}

/** Only https links are kept, so a value from a third party can never become a `javascript:` link. */
function httpsUrl(value: unknown): string {
  const url: string = text(value);

  return /^https:\/\//i.test(url) ? url : '';
}

function list(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

export function toTicketGroups(payload: unknown): ITicketGroup[] {
  return list(payload).filter(isObject).map((group: Record<string, unknown>): ITicketGroup => ({
    projectId: count(group.projectId),
    projectName: text(group.projectName),
    message: text(group.message),
    tickets: list(group.tickets).filter(isObject).map((ticket: Record<string, unknown>): ITicket => ({
      id: text(ticket.id),
      ticketNumber: text(ticket.ticketNumber),
      title: text(ticket.title),
      body: text(ticket.body),
      state: text(ticket.state).toLowerCase(),
      createdOn: text(ticket.createdOn),
      dueDate: text(ticket.dueDate),
      url: httpsUrl(ticket.url),
      attachmentCount: count(ticket.attachmentCount)
    }))
  }));
}

export function toPlanningGroups(payload: unknown): IPlanningGroup[] {
  return list(payload).filter(isObject).map((group: Record<string, unknown>): IPlanningGroup => ({
    projectId: count(group.projectId),
    projectName: text(group.projectName),
    url: httpsUrl(group.url),
    message: text(group.message),
    items: list(group.items).filter(isObject).map((item: Record<string, unknown>): IPlanningItem => ({
      name: text(item.name),
      startDate: text(item.startDate),
      endDate: text(item.endDate),
      totalDays: count(item.totalDays),
      elapsedDays: count(item.elapsedDays)
    }))
  }));
}

/** dd/mm/yyyy of an ISO date, read verbatim so the browser's time zone cannot shift it; empty when absent. */
export function formatWorkDate(value: string): string {
  const match: RegExpExecArray | null = /^(\d{4})-(\d{2})-(\d{2})/.exec((value || '').trim());

  return match ? `${match[3]}/${match[2]}/${match[1]}` : '';
}
