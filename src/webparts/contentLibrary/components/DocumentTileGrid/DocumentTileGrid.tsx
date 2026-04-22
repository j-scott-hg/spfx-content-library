import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IListItem } from '../../models/IListItem';
import { IItemIconOverride } from '../../models/IWebPartConfig';
import { formatDate, formatFieldValue } from '../../helpers/fieldFormatting';
import { normalizeChoiceValues, getChoiceBadgeStyle } from '../../helpers/choiceBadgeUtils';
import { IColumnDef } from '../../services/ViewMapper';
import { getFileIconInfo, getFileIconColorHex } from '../../helpers/fileIconMapping';
import { getContrastTextColor, tintColor } from '../../helpers/colorUtils';
import { LinkTarget } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

export interface IDocumentTileGridProps {
  items: IListItem[];
  gridColumns: number;
  showFileTypeIcon: boolean;
  isDocumentLibrary: boolean;
  linkTarget: LinkTarget;
  cardCornerRadius: number;
  isEditMode: boolean;
  itemIconOverrides: Record<string, IItemIconOverride>;
  onEditItemIcon: (itemId: string, defaultIconName: string, defaultIconColor: string, title: string) => void;
  onItemClick: (item: IListItem) => void;
  /** Internal field name for first meta line (empty = hide) */
  cardMeta1Field: string;
  /** Internal field name for second meta line (empty = hide) */
  cardMeta2Field: string;
  cardMeta1Icon: string;
  cardMeta2Icon: string;
  showChoicePillsOnCards: boolean;
  /** All available columns for resolving display names and types */
  allColumns: IColumnDef[];
  enableCategoryColors: boolean;
  categoryColors: Record<string, string>;
  categoryTextOverrides: Record<string, 'light' | 'dark'>;
  filterFieldInternalName: string;
  itemFontSize: number;
  itemIconSize: number;
}

const DocumentTileGrid: React.FC<IDocumentTileGridProps> = ({
  items,
  gridColumns,
  showFileTypeIcon,
  isDocumentLibrary,
  linkTarget,
  cardCornerRadius,
  isEditMode,
  itemIconOverrides,
  onEditItemIcon,
  onItemClick,
  cardMeta1Field,
  cardMeta2Field,
  cardMeta1Icon,
  cardMeta2Icon,
  showChoicePillsOnCards,
  allColumns,
  enableCategoryColors,
  categoryColors,
  categoryTextOverrides,
  filterFieldInternalName,
  itemFontSize,
  itemIconSize,
}) => {
  const colMap: Record<string, IColumnDef> = {};
  allColumns.forEach(c => { colMap[c.internalName] = c; });

  const resolveMetaValue = (item: IListItem, fieldName: string): string => {
    if (!fieldName) return '';
    if (fieldName === 'Modified') return item.modified ? formatDate(item.modified) : '';
    if (fieldName === 'Created') return item.created ? formatDate(item.created) : '';
    if (fieldName === 'Editor') return item.modifiedBy ?? '';
    if (fieldName === 'Author') return item.createdBy ?? '';
    const col = colMap[fieldName];
    const raw = item[fieldName];
    if (raw === undefined || raw === null || raw === '') return '';
    return col ? formatFieldValue(raw, col.fieldType) : String(raw);
  };

  const renderMetaValue = (item: IListItem, fieldName: string): React.ReactNode => {
    const col = colMap[fieldName];
    const raw = item[fieldName];
    const isChoiceField = col?.fieldType === 'Choice' || col?.fieldType === 'MultiChoice';
    if (showChoicePillsOnCards && isChoiceField) {
      const choices = normalizeChoiceValues(raw);
      if (choices.length > 0) {
        return (
          <span className={styles.metaChoicePills}>
            {choices.map(choice => {
              const pillStyle = getChoiceBadgeStyle(choice);
              return (
                <span key={choice} className={styles.choiceBadge} style={pillStyle}>
                  {choice}
                </span>
              );
            })}
          </span>
        );
      }
    }
    return <span className={styles.metaText}>{resolveMetaValue(item, fieldName)}</span>;
  };

  return (
    <div
      className={styles.tileGrid}
      style={{ '--grid-cols': gridColumns } as React.CSSProperties}
      role="list"
      aria-label="Document tiles"
    >
      {items.map(item => {
        const defaultIconInfo = getFileIconInfo(item.fileType, item.isFolder);
        const defaultIconColor = getFileIconColorHex(item.fileType, item.isFolder);
        const override = itemIconOverrides[item.id];
        const iconName = override?.iconName ?? defaultIconInfo.iconName;
        const iconColor = override?.iconColor ?? defaultIconColor;

        const displayTitle = isDocumentLibrary
          ? (item.fileLeafRef ?? item.title)
          : item.title;

        // Category colour coding
        const catValue = filterFieldInternalName ? String(item[filterFieldInternalName] ?? '') : '';
        const catBgHex = enableCategoryColors && catValue && categoryColors[catValue] ? categoryColors[catValue] : undefined;
        const tileBg = catBgHex ? tintColor(catBgHex, 0.18) : undefined;
        const textOverride = catValue ? categoryTextOverrides?.[catValue] : undefined;
        const tileTextColor = catBgHex
          ? (textOverride === 'light' ? '#ffffff' : textOverride === 'dark' ? '#1b1b1b' : getContrastTextColor(tintColor(catBgHex, 0.18)))
          : undefined;
        const effectiveIconColor = catBgHex ? catBgHex : iconColor;

        const tile = (
          <div
            role="button"
            tabIndex={0}
            className={styles.documentTile}
            style={{
              borderRadius: cardCornerRadius,
              cursor: 'pointer',
              ...(tileBg ? { background: tileBg, borderColor: `${catBgHex}60` } : {}),
            }}
            aria-label={displayTitle}
            onClick={() => onItemClick(item)}
            onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && onItemClick(item)}
          >
            {showFileTypeIcon && (
              <div
                className={styles.tileIconWrapper}
                style={{ background: `${effectiveIconColor}18`, borderRadius: cardCornerRadius / 2 }}
                aria-hidden="true"
              >
                <span style={{ color: effectiveIconColor, fontSize: itemIconSize }}>
                  <Icon iconName={iconName} />
                </span>
              </div>
            )}
            <div className={styles.tileTitle} style={{ fontSize: itemFontSize, ...(tileTextColor ? { color: tileTextColor } : {}) }} title={displayTitle}>
              {displayTitle}
            </div>
            {[cardMeta1Field, cardMeta2Field].map((fieldName, idx) => {
              const val = resolveMetaValue(item, fieldName);
              if (!val) return null;
              return (
                <div key={idx} className={styles.tileMeta} style={{ ...(tileTextColor ? { color: tileTextColor, opacity: 0.7 } : {}) }}>
                  <span className={styles.tileMetaLine}>
                    <Icon iconName={idx === 0 ? cardMeta1Icon : cardMeta2Icon} aria-hidden="true" className={styles.tileMetaIcon} />
                    {renderMetaValue(item, fieldName)}
                  </span>
                </div>
              );
            })}
          </div>
        );

        if (!isEditMode) {
          return <div key={item.id} role="listitem">{tile}</div>;
        }

        return (
          <div key={item.id} role="listitem" className={styles.itemEditWrapper}>
            {tile}
            <div className={styles.itemEditOverlay}>
              <button
                className={styles.itemEditButton}
                onClick={e => {
                  e.preventDefault();
                  e.stopPropagation();
                  onEditItemIcon(item.id, defaultIconInfo.iconName, defaultIconColor, displayTitle);
                }}
                title="Edit icon"
                aria-label={`Edit icon for ${displayTitle}`}
              >
                <Icon iconName="Color" style={{ fontSize: 13 }} />
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );
};

export default DocumentTileGrid;
