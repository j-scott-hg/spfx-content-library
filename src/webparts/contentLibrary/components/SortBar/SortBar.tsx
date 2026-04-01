import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { SortDirection } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

export interface ISortBarProps {
  sortField: string;
  sortDirection: SortDirection;
  sortFields: Array<{ key: string; text: string }>;
  onSortChange: (field: string, direction: SortDirection) => void;
  /** When true, renders inline with the toolbar search bar (no separate row) */
  inToolbar?: boolean;
}

const SortBar: React.FC<ISortBarProps> = ({
  sortField,
  sortDirection,
  sortFields,
  onSortChange,
  inToolbar,
}) => {
  const options: IDropdownOption[] = sortFields.map(f => ({ key: f.key, text: f.text }));

  const handleFieldChange = (_: React.FormEvent, option?: IDropdownOption): void => {
    if (option) onSortChange(String(option.key), sortDirection);
  };

  const toggleDirection = (): void => {
    onSortChange(sortField, sortDirection === 'asc' ? 'desc' : 'asc');
  };

  const dirLabel = sortDirection === 'asc' ? 'Ascending' : 'Descending';
  const dirIcon = sortDirection === 'asc' ? 'SortUp' : 'SortDown';

  return (
    <div
      className={inToolbar ? styles.sortBarInline : styles.sortBarStandalone}
      role="group"
      aria-label="Sort controls"
    >
      <span className={styles.sortBarLabel}>Sort by</span>
      <Dropdown
        selectedKey={sortField || undefined}
        options={options}
        onChange={handleFieldChange}
        placeholder="Select field"
        styles={{
          root: { minWidth: 140 },
          dropdown: { fontSize: 13, height: 30 },
          title: { lineHeight: '28px', height: 30, fontSize: 13 },
          caretDownWrapper: { lineHeight: '28px', height: 30 },
        }}
        ariaLabel="Sort by field"
      />
      <button
        className={styles.sortDirButton}
        onClick={toggleDirection}
        title={`Sort ${dirLabel}`}
        aria-label={`Sort ${dirLabel} — click to toggle`}
        aria-pressed={sortDirection === 'asc'}
      >
        <Icon iconName={dirIcon} style={{ fontSize: 14 }} />
        <span className={styles.sortDirLabel}>{dirLabel}</span>
      </button>
    </div>
  );
};

export default SortBar;
