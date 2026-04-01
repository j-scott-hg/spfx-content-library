import * as React from 'react';
import { useCallback, useState, useRef, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { IListItem } from '../../models/IListItem';
import { IColumnDef } from '../../services/ViewMapper';
import { IItemIconOverride } from '../../models/IWebPartConfig';
import { formatDate, formatFieldValue } from '../../helpers/fieldFormatting';
import { getFileIconInfo, getFileIconColorHex } from '../../helpers/fileIconMapping';
import { LinkTarget, SortDirection } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

export interface IDocumentTableViewProps {
  items: IListItem[];
  columns: IColumnDef[];
  showColumnHeaders: boolean;
  showFileTypeIcon: boolean;
  showModifiedDate: boolean;
  showModifiedBy: boolean;
  isDocumentLibrary: boolean;
  linkTarget: LinkTarget;
  sortField: string;
  sortDirection: SortDirection;
  allowUserSort: boolean;
  onSortChange: (field: string, direction: SortDirection) => void;
  onItemClick: (item: IListItem) => void;
  isEditMode: boolean;
  itemIconOverrides: Record<string, IItemIconOverride>;
  onEditItemIcon: (itemId: string, defaultIconName: string, defaultIconColor: string, title: string) => void;
  itemFontSize: number;
  itemIconSize: number;
}

const DocumentTableView: React.FC<IDocumentTableViewProps> = ({
  items,
  columns,
  showColumnHeaders,
  showFileTypeIcon,
  showModifiedDate,
  showModifiedBy,
  isDocumentLibrary,
  linkTarget,
  sortField,
  sortDirection,
  allowUserSort,
  onSortChange,
  onItemClick,
  isEditMode,
  itemIconOverrides,
  onEditItemIcon,
  itemFontSize,
  itemIconSize,
}) => {
  // ── Column picker state ───────────────────────────────────────────────────
  const [pickerOpen, setPickerOpen] = useState(false);
  const [hiddenCols, setHiddenCols] = useState<Record<string, boolean>>({});

  const effectiveColumns = columns.filter(
    c => c.internalName !== 'Title' && c.internalName !== 'FileLeafRef'
  );

  // ── Column widths (pixel values, user-resizable) ──────────────────────────
  // Default widths computed from field type and label length
  const getDefaultWidth = (col: IColumnDef): number => {
    const ft = col.fieldType;
    if (ft === 'DateTime') return 150;
    if (ft === 'User' || ft === 'UserMulti') return 160;
    if (ft === 'Boolean') return 90;
    if (ft === 'Number' || ft === 'Currency') return 110;
    if (ft === 'Note') return 220;
    return Math.max(110, col.displayName.length * 9 + 32);
  };

  // colWidths maps a stable key → pixel width
  const buildDefaultWidths = useCallback((): Record<string, number> => {
    const w: Record<string, number> = { __name: 220 };
    effectiveColumns.forEach(c => { w[c.internalName] = getDefaultWidth(c); });
    if (showModifiedDate) w['__modified'] = 150;
    if (showModifiedBy)  w['__modifiedBy'] = 160;
    return w;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [columns, showModifiedDate, showModifiedBy]);

  const [colWidths, setColWidths] = useState<Record<string, number>>(buildDefaultWidths);

  // Reset widths and hidden state when the column set changes
  const colKey = columns.map(c => c.internalName).join(',');
  useEffect(() => {
    setHiddenCols({});
    setColWidths(buildDefaultWidths());
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [colKey]);

  // ── Resize drag logic ─────────────────────────────────────────────────────
  const resizingCol = useRef<string | null>(null);
  const resizeStartX = useRef(0);
  const resizeStartWidth = useRef(0);

  const onResizeMouseDown = useCallback((e: React.MouseEvent, colId: string) => {
    e.preventDefault();
    e.stopPropagation();
    resizingCol.current = colId;
    resizeStartX.current = e.clientX;
    resizeStartWidth.current = colWidths[colId] ?? 120;

    const onMouseMove = (ev: MouseEvent): void => {
      if (!resizingCol.current) return;
      const delta = ev.clientX - resizeStartX.current;
      const newWidth = Math.max(60, resizeStartWidth.current + delta);
      setColWidths(prev => ({ ...prev, [resizingCol.current as string]: newWidth }));
    };

    const onMouseUp = (): void => {
      resizingCol.current = null;
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  }, [colWidths]);

  // Build gridTemplateColumns from pixel widths
  const visibleEffectiveCols = effectiveColumns.filter(col => !hiddenCols[col.internalName]);
  const showName    = !hiddenCols['__name'];
  const showModDate = showModifiedDate && !hiddenCols['__modified'];
  const showModBy   = showModifiedBy   && !hiddenCols['__modifiedBy'];

  const colParts = [
    isEditMode ? '36px' : '',
    showName    ? `${colWidths['__name'] ?? 220}px` : '',
    ...visibleEffectiveCols.map(col => `${colWidths[col.internalName] ?? 120}px`),
    showModDate ? `${colWidths['__modified'] ?? 150}px` : '',
    showModBy   ? `${colWidths['__modifiedBy'] ?? 160}px` : '',
  ].filter(Boolean);
  const gridTemplateColumns = colParts.join(' ');

  // ── Column picker helpers ─────────────────────────────────────────────────
  const allPickerCols: Array<{ internalName: string; displayName: string }> = [
    { internalName: '__name', displayName: 'Name' },
    ...effectiveColumns,
    ...(showModifiedDate ? [{ internalName: '__modified', displayName: 'Modified' }] : []),
    ...(showModifiedBy  ? [{ internalName: '__modifiedBy', displayName: 'Modified By' }] : []),
  ];

  const toggleCol = (internalName: string): void => {
    setHiddenCols(prev => ({ ...prev, [internalName]: !prev[internalName] }));
  };

  const handleSortClick = useCallback(
    (colName: string) => {
      if (!allowUserSort) return;
      if (sortField === colName) {
        onSortChange(colName, sortDirection === 'asc' ? 'desc' : 'asc');
      } else {
        onSortChange(colName, 'asc');
      }
    },
    [allowUserSort, sortField, sortDirection, onSortChange]
  );

  const SortIcon: React.FC<{ field: string }> = ({ field }) => {
    if (sortField !== field) return <Icon iconName="Sort" style={{ fontSize: 10, opacity: 0.4 }} />;
    return <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} style={{ fontSize: 10 }} />;
  };

  return (
    <>
      {/* Column picker trigger bar */}
      <div className={styles.tableToolbar}>
        <button
          className={styles.tableColPickerBtn}
          onClick={() => setPickerOpen(true)}
          title="Show/hide columns"
          aria-label="Show or hide columns"
        >
          <Icon iconName="ColumnOptions" style={{ fontSize: 14 }} />
          <span>Columns</span>
        </button>
      </div>

      {/* Column visibility picker panel */}
      <Panel
        isOpen={pickerOpen}
        onDismiss={() => setPickerOpen(false)}
        type={PanelType.custom}
        customWidth="260px"
        isLightDismiss
        hasCloseButton
        closeButtonAriaLabel="Close"
        headerText="Edit view columns"
        styles={{
          main: { boxShadow: '0 8px 32px rgba(0,0,0,0.18)' },
          header: { paddingTop: 20 },
          content: { paddingTop: 8 },
        }}
      >
        <p className={styles.colPickerHint}>
          Select the columns to display in the table view.
        </p>
        <div className={styles.colPickerList}>
          {allPickerCols.map(col => (
            <div key={col.internalName} className={styles.colPickerRow}>
              <Checkbox
                label={col.displayName}
                checked={!hiddenCols[col.internalName]}
                onChange={() => toggleCol(col.internalName)}
                styles={{
                  root: { padding: '6px 0' },
                  label: { fontSize: 14, fontWeight: 400 },
                }}
              />
            </div>
          ))}
        </div>
      </Panel>

      {/* Table */}
      <div className={styles.tableScrollWrapper}>
        <div className={styles.tableContainer} role="table" aria-label="Document list">
          {showColumnHeaders && (
            <div className={styles.tableHeader} style={{ gridTemplateColumns }} role="row">
              {isEditMode && <div className={styles.tableHeaderCell} role="columnheader" aria-label="Icon" />}
              {showName && (
                <div className={styles.tableHeaderCell} role="columnheader"
                  onClick={() => handleSortClick(isDocumentLibrary ? 'FileLeafRef' : 'Title')}
                  onKeyDown={e => e.key === 'Enter' && handleSortClick(isDocumentLibrary ? 'FileLeafRef' : 'Title')}
                  tabIndex={allowUserSort ? 0 : -1}
                  aria-sort={sortField === (isDocumentLibrary ? 'FileLeafRef' : 'Title') ? (sortDirection === 'asc' ? 'ascending' : 'descending') : 'none'}
                >
                  <span className={styles.headerLabel}>Name</span>
                  {allowUserSort && <SortIcon field={isDocumentLibrary ? 'FileLeafRef' : 'Title'} />}
                  <div className={styles.colResizeHandle} onMouseDown={e => onResizeMouseDown(e, '__name')} aria-hidden="true" />
                </div>
              )}
              {visibleEffectiveCols.map(col => (
                <div key={col.internalName} className={styles.tableHeaderCell} role="columnheader"
                  onClick={() => handleSortClick(col.internalName)}
                  onKeyDown={e => e.key === 'Enter' && handleSortClick(col.internalName)}
                  tabIndex={allowUserSort ? 0 : -1}
                  aria-sort={sortField === col.internalName ? (sortDirection === 'asc' ? 'ascending' : 'descending') : 'none'}
                >
                  <span className={styles.headerLabel}>{col.displayName}</span>
                  {allowUserSort && <SortIcon field={col.internalName} />}
                  <div className={styles.colResizeHandle} onMouseDown={e => onResizeMouseDown(e, col.internalName)} aria-hidden="true" />
                </div>
              ))}
              {showModDate && (
                <div className={styles.tableHeaderCell} role="columnheader"
                  onClick={() => handleSortClick('Modified')}
                  onKeyDown={e => e.key === 'Enter' && handleSortClick('Modified')}
                  tabIndex={allowUserSort ? 0 : -1}
                  aria-sort={sortField === 'Modified' ? (sortDirection === 'asc' ? 'ascending' : 'descending') : 'none'}
                >
                  <span className={styles.headerLabel}>Modified</span>
                  {allowUserSort && <SortIcon field="Modified" />}
                  <div className={styles.colResizeHandle} onMouseDown={e => onResizeMouseDown(e, '__modified')} aria-hidden="true" />
                </div>
              )}
              {showModBy && (
                <div className={styles.tableHeaderCell} role="columnheader">
                  <span className={styles.headerLabel}>Modified By</span>
                  <div className={styles.colResizeHandle} onMouseDown={e => onResizeMouseDown(e, '__modifiedBy')} aria-hidden="true" />
                </div>
              )}
            </div>
          )}

          <div role="rowgroup">
            {items.map(item => {
              const defaultIconInfo = getFileIconInfo(item.fileType, item.isFolder);
              const defaultIconColor = getFileIconColorHex(item.fileType, item.isFolder);
              const override = itemIconOverrides[item.id];
              const iconName = override?.iconName ?? defaultIconInfo.iconName;
              const iconColor = override?.iconColor ?? defaultIconColor;

              const displayTitle = item.isFolder
                ? (item.fileLeafRef ?? item.title)
                : (isDocumentLibrary ? (item.fileLeafRef ?? item.title) : item.title);

              return (
                <div
                  key={item.id}
                  className={styles.tableRow}
                  style={{ gridTemplateColumns }}
                  role="row"
                >
                  {isEditMode && (
                    <div className={`${styles.tableCell} ${styles.tableEditCell}`} role="cell">
                      <button
                        className={styles.itemEditButton}
                        onClick={() => onEditItemIcon(item.id, defaultIconInfo.iconName, defaultIconColor, displayTitle)}
                        title="Edit icon"
                        aria-label={`Edit icon for ${displayTitle}`}
                        style={{ position: 'static', opacity: 1 }}
                      >
                        <Icon iconName="Color" style={{ fontSize: 12 }} />
                      </button>
                    </div>
                  )}

                  {showName && (
                    <div className={`${styles.tableCell} ${styles.tableTitleCell}`} role="cell">
                      {showFileTypeIcon && (
                        <span
                          className={styles.fileIconBase}
                          style={{ color: iconColor, fontSize: itemIconSize }}
                          aria-label={defaultIconInfo.label}
                          role="img"
                        >
                          <Icon iconName={iconName} />
                        </span>
                      )}
                      <button
                        className={styles.tableItemLink}
                        title={displayTitle}
                        aria-label={displayTitle}
                        onClick={() => onItemClick(item)}
                        style={{ fontSize: itemFontSize }}
                      >
                        {displayTitle}
                      </button>
                    </div>
                  )}

                  {visibleEffectiveCols.map(col => (
                    <div key={col.internalName} className={styles.tableCell} role="cell">
                      {formatFieldValue(item[col.internalName], col.fieldType)}
                    </div>
                  ))}

                  {showModDate && (
                    <div className={styles.tableCell} role="cell">
                      {formatDate(item.modified)}
                    </div>
                  )}

                  {showModBy && (
                    <div className={styles.tableCell} role="cell">
                      {item.modifiedBy ?? ''}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </>
  );
};

export default DocumentTableView;
