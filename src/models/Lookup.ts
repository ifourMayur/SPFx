/**
 * The Web API's `LookupItem` shape and its mapping onto dropdown options.
 *
 * Part of the models layer: pure, with no imports outside `models`.
 *
 * Every reference-data endpoint this solution calls - project templates, users, spatial
 * breakdowns, tender templates, suppliers and each folder level - returns
 * `List<LookupItem>` inside the `ResponseDetail` envelope, so one mapper serves all of
 * them.
 */

/* eslint-disable @rushstack/no-new-null -- These interfaces describe the BMDesk Web API's
   wire format, which is the "legacy API" case the rule exempts: it serializes with
   System.Text.Json defaults, so absent values arrive as an explicit `null` rather than
   being omitted. Declaring them as `| null` is what makes the null checks below type-check
   instead of looking redundant. */

import { ILookupOption } from './Building';

/**
 * `BMDesk.Business.Common.LookupItem`, camelCased by System.Text.Json.
 *
 * Only the three fields this solution reads are declared; the server type also carries
 * colour, index and SharePoint-path fields that no dropdown here uses. Every field is
 * optional and nullable because System.Text.Json writes `null` rather than omitting keys.
 */
export interface IApiLookupItem {
  id?: number | null;
  /** String identifier used by the endpoints whose keys are not integers. */
  strId?: string | null;
  name?: string | null;
}

/** Returns the trimmed text when it actually holds something, otherwise undefined. */
function text(value: string | null | undefined): string | undefined {
  const trimmed: string = (value || '').trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/**
 * Maps an API lookup collection onto `<option>` values.
 *
 * `strId` wins when present, otherwise the numeric `id` is stringified - a `<select>`
 * value is always a string, and the reference binds these with
 * `new SelectList(list, "Id", "Name")`.
 *
 * Anything unusable is dropped rather than allowed to produce a broken option: a
 * non-array payload yields an empty list (so a failed load renders an empty dropdown
 * instead of throwing), entries with no identifier are skipped, and a `null` name becomes
 * an empty label rather than the string "null".
 */
export function toLookupOptions(payload: unknown): ILookupOption[] {
  if (!Array.isArray(payload)) {
    return [];
  }

  const options: ILookupOption[] = [];

  payload.forEach((entry: unknown): void => {
    if (typeof entry !== 'object' || entry === null) {
      return;
    }

    const item: IApiLookupItem = entry as IApiLookupItem;
    const id: string | undefined =
      text(item.strId) || (typeof item.id === 'number' ? String(item.id) : undefined);

    if (id === undefined) {
      return;
    }

    options.push({ id, name: text(item.name) || '' });
  });

  return options;
}
