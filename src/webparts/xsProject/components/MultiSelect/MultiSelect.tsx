import * as React from 'react';

import styles from './MultiSelect.module.scss';
import type { IMultiSelectProps } from './IMultiSelectProps';
import { ILookupOption } from '../../../../models/Building';

/** Shown when the caller supplies no `emptyText`. */
const DEFAULT_EMPTY_TEXT: string = 'No options available.';

interface IMultiSelectState {
  isOpen: boolean;
  search: string;
}

/**
 * Multi-select with removable chips and a checkbox list.
 *
 * The reference application renders these fields as `@Html.ListBoxFor(..., multiple)` and
 * then upgrades them with select2, which is what gives them chips, a search box and
 * click-to-toggle. A bare `<select multiple>` is technically the same control but is
 * unusable in practice - it needs ctrl-click, shows no chips, and silently drops the whole
 * selection on a stray click - so this reproduces the select2 behaviour without adding the
 * dependency.
 *
 * Selection lives entirely in props; only "is the panel open" and the search term are local
 * state, so the parent form stays the single source of truth.
 */
export default class MultiSelect extends React.Component<IMultiSelectProps, IMultiSelectState> {
  private readonly _rootRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();

  public constructor(props: IMultiSelectProps) {
    super(props);

    this.state = { isOpen: false, search: '' };
  }

  public componentDidMount(): void {
    document.addEventListener('mousedown', this._onDocumentMouseDown);
  }

  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this._onDocumentMouseDown);
  }

  public render(): React.ReactElement<IMultiSelectProps> {
    const { inputId, label, options, selectedIds, isDisabled, placeholder, emptyText } = this.props;
    const { isOpen } = this.state;
    const selected: ILookupOption[] = options.filter(
      (option: ILookupOption) => selectedIds.indexOf(option.id) >= 0
    );

    return (
      <div className={styles.multiSelect} ref={this._rootRef}>
        <label className={styles.label} htmlFor={inputId}>{label}</label>

        <button
          type="button"
          id={inputId}
          className={styles.control}
          disabled={isDisabled}
          aria-haspopup="listbox"
          aria-expanded={isOpen}
          onClick={this._onToggleOpen}
        >
          {selected.length === 0
            ? <span className={styles.placeholder}>{placeholder || 'Select…'}</span>
            : (
              <span className={styles.chips}>
                {selected.map((option: ILookupOption) => (
                  <span key={option.id} className={styles.chip}>
                    <span className={styles.chipLabel}>{option.name}</span>
                    {/* A span, not a button: a button inside a button is invalid markup, and
                        the click is handled here on the parent via the data attribute. */}
                    <span
                      className={styles.chipRemove}
                      role="button"
                      aria-label={`Remove ${option.name}`}
                      data-id={option.id}
                      onClick={this._onRemoveChip}
                    >
                      &times;
                    </span>
                  </span>
                ))}
              </span>
            )}
          <span className={styles.caret} aria-hidden="true">{isOpen ? '▴' : '▾'}</span>
        </button>

        {isOpen && this._renderPanel()}

        {options.length === 0 && (
          <span className={styles.emptyHint}>{emptyText || DEFAULT_EMPTY_TEXT}</span>
        )}
      </div>
    );
  }

  private _renderPanel(): React.ReactElement {
    const { options, selectedIds, emptyText } = this.props;
    const term: string = this.state.search.trim().toLowerCase();
    const visible: ILookupOption[] = term
      ? options.filter((option: ILookupOption) => option.name.toLowerCase().indexOf(term) >= 0)
      : options;

    return (
      <div className={styles.panel} role="listbox" aria-multiselectable="true">
        <input
          type="search"
          className={styles.search}
          value={this.state.search}
          placeholder="Search"
          aria-label="Search options"
          onChange={this._onSearchChange}
          onKeyDown={this._onSearchKeyDown}
        />

        <ul className={styles.list}>
          {visible.map((option: ILookupOption) => {
            const isSelected: boolean = selectedIds.indexOf(option.id) >= 0;

            return (
              <li key={option.id}>
                <label className={styles.option}>
                  <input
                    type="checkbox"
                    checked={isSelected}
                    data-id={option.id}
                    onChange={this._onToggleOption}
                  />
                  <span>{option.name}</span>
                </label>
              </li>
            );
          })}

          {visible.length === 0 && (
            <li className={styles.empty}>
              {options.length === 0 ? (emptyText || DEFAULT_EMPTY_TEXT) : 'No matches.'}
            </li>
          )}
        </ul>
      </div>
    );
  }

  // Assigned as properties so `this` is bound without a per-render closure, matching the
  // pattern the other components here use.
  private _onToggleOpen = (): void => {
    this.setState((state: IMultiSelectState) => ({ isOpen: !state.isOpen, search: '' }));
  };

  private _onSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ search: event.target.value });
  };

  private _onSearchKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
    if (event.key === 'Escape') {
      this.setState({ isOpen: false, search: '' });
    }
  };

  private _onToggleOption = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const id: string = event.target.dataset.id as string;
    const { selectedIds } = this.props;
    const isSelected: boolean = selectedIds.indexOf(id) >= 0;

    // The panel deliberately stays open: choosing several options is the normal case.
    this.props.onChange(
      isSelected
        ? selectedIds.filter((selectedId: string) => selectedId !== id)
        : selectedIds.concat([id])
    );
  };

  private _onRemoveChip = (event: React.MouseEvent<HTMLSpanElement>): void => {
    // Without this the click would also reach the control and toggle the panel.
    event.stopPropagation();

    const id: string = event.currentTarget.dataset.id as string;
    this.props.onChange(this.props.selectedIds.filter((selectedId: string) => selectedId !== id));
  };

  private _onDocumentMouseDown = (event: MouseEvent): void => {
    if (!this.state.isOpen) {
      return;
    }

    const root: HTMLDivElement | null = this._rootRef.current;
    if (root && !root.contains(event.target as Node)) {
      this.setState({ isOpen: false, search: '' });
    }
  };
}
