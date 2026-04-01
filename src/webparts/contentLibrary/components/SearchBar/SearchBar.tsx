import * as React from 'react';
import { useCallback, useRef } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IconButton } from '@fluentui/react/lib/Button';
import { ISearchBarProps } from './ISearchBarProps';
import styles from '../../styles/ContentLibrary.module.scss';

const SearchBar: React.FC<ISearchBarProps> = ({
  value,
  onChange,
  placeholder = 'Search...',
  style,
  wrapperStyle,
  showSortControl,
  sortFields,
  currentSortField,
  currentSortAsc,
  onSortChange,
}) => {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      onChange(e.target.value);
    },
    [onChange]
  );

  const handleClear = useCallback(() => {
    onChange('');
    inputRef.current?.focus();
  }, [onChange]);

  const searchInput = (
    <div className={styles.searchBarMinimal} role="search">
      <span className={styles.searchIcon} aria-hidden="true">
        <Icon iconName="Search" />
      </span>
      <input
        ref={inputRef}
        type="search"
        value={value}
        onChange={handleChange}
        placeholder={placeholder}
        aria-label={placeholder}
        autoComplete="off"
      />
      {value && (
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          title="Clear search"
          ariaLabel="Clear search"
          onClick={handleClear}
          styles={{
            root: { width: 24, height: 24, padding: 0 },
            icon: { fontSize: 12, color: '#605e5c' },
          }}
        />
      )}
    </div>
  );

  if (style === 'elevated') {
    return (
      <div className={styles.searchBarElevated} style={wrapperStyle}>
        {searchInput}
      </div>
    );
  }

  if (style === 'toolbar') {
    const sortOptions: IDropdownOption[] = (sortFields ?? []).map(f => ({
      key: f.key,
      text: f.text,
    }));

    const handleSortFieldChange = (_: React.FormEvent, option?: IDropdownOption): void => {
      if (option && onSortChange) {
        onSortChange(String(option.key), currentSortAsc !== false);
      }
    };

    const toggleSortDir = (): void => {
      if (onSortChange && currentSortField) {
        onSortChange(currentSortField, !currentSortAsc);
      }
    };

    return (
      <div className={styles.searchBarToolbar} role="toolbar" aria-label="Search and sort toolbar">
        <div className={styles.toolbarSpacer} />
        {showSortControl && sortOptions.length > 0 && (
          <>
            <Dropdown
              placeholder="Sort by"
              selectedKey={currentSortField}
              options={sortOptions}
              onChange={handleSortFieldChange}
              styles={{ root: { width: 160 }, dropdown: { fontSize: 13 } }}
              ariaLabel="Sort by field"
            />
            <IconButton
              iconProps={{ iconName: currentSortAsc ? 'SortUp' : 'SortDown' }}
              title={currentSortAsc ? 'Sort ascending' : 'Sort descending'}
              ariaLabel={currentSortAsc ? 'Sort ascending' : 'Sort descending'}
              onClick={toggleSortDir}
              styles={{ root: { height: 32 } }}
            />
          </>
        )}
        {searchInput}
      </div>
    );
  }

  // Default: minimal — apply wrapperStyle if provided
  if (wrapperStyle) {
    return <div style={wrapperStyle}>{searchInput}</div>;
  }
  return searchInput;
};

export default SearchBar;
