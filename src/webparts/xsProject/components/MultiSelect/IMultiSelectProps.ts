import { ILookupOption } from '../../../../models/Building';

export interface IMultiSelectProps {
  /** Ties the label to the control and keeps ids unique when several are on one screen. */
  inputId: string;
  /** Visible label. */
  label: string;
  /** Everything selectable. An empty list renders the {@link emptyText} hint. */
  options: ILookupOption[];
  /** Currently selected option ids. */
  selectedIds: string[];
  /** Called with the complete next selection. */
  onChange: (selectedIds: string[]) => void;
  /** Text shown while nothing is selected. */
  placeholder?: string;
  /**
   * Shown in place of the option list when `options` is empty, so "the API returned
   * nothing" never looks the same as "this control is broken".
   */
  emptyText?: string;
  isDisabled?: boolean;
}
