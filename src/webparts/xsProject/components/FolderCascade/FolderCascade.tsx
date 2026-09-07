import * as React from 'react';

import styles from './FolderCascade.module.scss';
import type { IFolderCascadeProps } from './IFolderCascadeProps';
import { ILookupOption } from '../../../../models/Building';
import { FolderLevel, selectFolderLevel } from '../../../../models/FolderTree';

/** Label per cascade level, from `BuildingResource`. */
const LEVEL_LABELS: { [level in FolderLevel]: string } = {
  folderId: 'Folder',
  subFolderId: 'Sub Folder',
  subSubFolderId: 'Sub Sub Folder',
  subSubSubFolderId: 'Sub Sub Sub Folder'
};

/** Placeholder option per level, matching the reference's "Select folder" text. */
const LEVEL_PLACEHOLDERS: { [level in FolderLevel]: string } = {
  folderId: 'Select folder',
  subFolderId: 'Select sub folder',
  subSubFolderId: 'Select sub sub folder',
  subSubSubFolderId: 'Select sub sub sub folder'
};

/** `BuildingResource.msgLoading`, shown while a level is being fetched. */
const LOADING_TEXT: string = 'Loading...';

/**
 * The `Folder -> Sub Folder -> Sub Sub Folder -> Sub Sub Sub Folder` dropdowns plus their
 * supplier, rendered as the reference's grey panel.
 *
 * Extracted as its own component because the reference repeats this exact group (per ED
 * Controls tag row, and again for MOM), so a second caller needs no new markup.
 *
 * Fully controlled and stateless: the selection and each level's options come in through
 * props, and every change goes straight back out. Which deeper levels a change discards
 * lives in `src/models/FolderTree.ts`, where it is unit-tested; fetching the next level is
 * the parent's job, which is why `onChange` reports the level that changed.
 *
 * While a level is loading it shows the reference's own "Loading..." placeholder and is
 * disabled, mirroring `setDropdownLoading` / `clearDropdownLoading` in the cshtml.
 */
export default class FolderCascade extends React.Component<IFolderCascadeProps> {
  public render(): React.ReactElement<IFolderCascadeProps> {
    const { isGrandChildFolderEnabled } = this.props;

    return (
      <div className={styles.folderCascadeGroup}>
        <div className={styles.row}>
          {this._renderLevel('folderId')}
          {this._renderLevel('subFolderId')}
          {this._renderLevel('subSubFolderId')}
          {isGrandChildFolderEnabled && this._renderLevel('subSubSubFolderId')}
        </div>

        <div className={styles.row}>
          {this._renderSupplier()}
        </div>
      </div>
    );
  }

  private _renderLevel(level: FolderLevel): React.ReactElement {
    const { idPrefix, options, value, isDisabled } = this.props;
    const levelOptions: ILookupOption[] = options[level] || [];
    const isLoading: boolean = (this.props.loadingLevels || []).indexOf(level) >= 0;
    const inputId: string = `${idPrefix}-${level}`;

    return (
      <div className={styles.field} key={level}>
        <label className={styles.label} htmlFor={inputId}>{LEVEL_LABELS[level]}</label>
        <select
          id={inputId}
          className={`${styles.select} ${isLoading ? styles.loading : ''}`}
          value={value[level]}
          // A level with nothing to offer - its parent is unselected, or it is still
          // loading - is disabled rather than presented as an empty, clickable dropdown.
          disabled={isDisabled || isLoading || levelOptions.length === 0}
          data-level={level}
          onChange={this._onLevelChange}
        >
          <option value="">{isLoading ? LOADING_TEXT : LEVEL_PLACEHOLDERS[level]}</option>
          {levelOptions.map((option: ILookupOption) => (
            <option key={option.id} value={option.id}>{option.name}</option>
          ))}
        </select>
      </div>
    );
  }

  private _renderSupplier(): React.ReactElement {
    const { idPrefix, suppliers, value, isDisabled } = this.props;
    const inputId: string = `${idPrefix}-supplierId`;

    return (
      <div className={styles.field}>
        <label className={styles.label} htmlFor={inputId}>Supplier</label>
        <select
          id={inputId}
          className={styles.select}
          value={value.supplierId}
          disabled={isDisabled}
          onChange={this._onSupplierChange}
        >
          <option value="">Select supplier</option>
          {suppliers.map((option: ILookupOption) => (
            <option key={option.id} value={option.id}>{option.name}</option>
          ))}
        </select>
      </div>
    );
  }

  // Assigned as properties so the `this` pointer is bound without a per-render closure,
  // matching the pattern `Login` / `Menu` use for their own handlers. One handler covers
  // all four levels, keyed off `data-level`.
  private _onLevelChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const level: FolderLevel = event.target.dataset.level as FolderLevel;

    this.props.onChange(selectFolderLevel(this.props.value, level, event.target.value), level);
  };

  private _onSupplierChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.props.onSupplierChange(event.target.value);
  };
}
