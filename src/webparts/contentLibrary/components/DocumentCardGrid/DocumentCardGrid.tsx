import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IListItem } from '../../models/IListItem';
import { IColumnDef } from '../../services/ViewMapper';
import { IItemIconOverride } from '../../models/IWebPartConfig';
import { formatDate, formatFieldValue, truncate } from '../../helpers/fieldFormatting';
import { getFileIconInfo, getFileIconColorHex } from '../../helpers/fileIconMapping';
import { getContrastTextColor, tintColor } from '../../helpers/colorUtils';
import { LinkTarget } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

export interface IDocumentCardGridProps {
  items: IListItem[];
  columns: IColumnDef[];
  gridColumns: number;
  showFileTypeIcon: boolean;
  showDescription: boolean;
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
  /** All available columns for resolving display names */
  allColumns: IColumnDef[];
  /** Category colour coding */
  enableCategoryColors: boolean;
  categoryColors: Record<string, string>;
  categoryTextOverrides: Record<string, 'light' | 'dark'>;
  filterFieldInternalName: string;
  itemFontSize: number;
  itemIconSize: number;
}

const DocumentCardGrid: React.FC<IDocumentCardGridProps> = ({
  items,
  columns,
  gridColumns,
  showFileTypeIcon,
  showDescription,
  isDocumentLibrary,
  linkTarget,
  cardCornerRadius,
  isEditMode,
  itemIconOverrides,
  onEditItemIcon,
  onItemClick,
  cardMeta1Field,
  cardMeta2Field,
  allColumns,
  enableCategoryColors,
  categoryColors,
  categoryTextOverrides,
  filterFieldInternalName,
  itemFontSize,
  itemIconSize,
}) => {
  // Build a lookup of internalName → IColumnDef for resolving meta field display names and types
  const colMap: Record<string, IColumnDef> = {};
  allColumns.forEach(c => { colMap[c.internalName] = c; });

  // Resolve a meta field value for display — handles built-in fields (Modified, Editor, etc.)
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

  // Pick an icon for a meta field
  const metaIcon = (fieldName: string): string => {
    if (fieldName === 'Modified' || fieldName === 'Created') return 'Clock';
    if (fieldName === 'Editor' || fieldName === 'Author') return 'Contact';
    const col = colMap[fieldName];
    if (!col) return 'Info';
    if (col.fieldType === 'DateTime') return 'Clock';
    if (col.fieldType === 'User') return 'Contact';
    if (col.fieldType === 'URL') return 'Link';
    if (col.fieldType === 'Boolean') return 'CheckMark';
    return 'Tag';
  };

  const descriptionCol = columns.find(
    c => c.internalName === 'Description' || c.fieldType === 'Note' || c.internalName === '_ExtendedDescription'
  );

  return (
    <div
      className={styles.cardGrid}
      style={{ '--grid-cols': gridColumns } as React.CSSProperties}
      role="list"
      aria-label="Document cards"
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

        const descriptionText = showDescription
          ? (descriptionCol ? formatFieldValue(item[descriptionCol.internalName], descriptionCol.fieldType) : (item.description ? String(item.description) : ''))
          : '';

        // Category colour coding
        const catValue = filterFieldInternalName ? String(item[filterFieldInternalName] ?? '') : '';
        const catBgHex = enableCategoryColors && catValue && categoryColors[catValue] ? categoryColors[catValue] : undefined;
        const cardBg = catBgHex ? tintColor(catBgHex, 0.18) : undefined;
        const textOverride = catValue ? categoryTextOverrides?.[catValue] : undefined;
        const cardTextColor = catBgHex
          ? (textOverride === 'light' ? '#ffffff' : textOverride === 'dark' ? '#1b1b1b' : getContrastTextColor(tintColor(catBgHex, 0.18)))
          : undefined;
        const cardBorderColor = catBgHex ? `${catBgHex}60` : undefined;
        const effectiveIconColor = catBgHex ? catBgHex : iconColor;

        const cardStyle: React.CSSProperties = {
          borderRadius: cardCornerRadius,
          ...(cardBg ? { background: cardBg, borderColor: cardBorderColor } : {}),
        };

        const card = (
          <div
            role="button"
            tabIndex={0}
            className={styles.documentCard}
            style={{ ...cardStyle, cursor: 'pointer' }}
            aria-label={displayTitle}
            onClick={() => onItemClick(item)}
            onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && onItemClick(item)}
          >
            {/* Category colour accent bar */}
            {catBgHex && (
              <div
                style={{
                  height: 3,
                  background: catBgHex,
                  margin: `-16px -16px 12px`,
                  borderRadius: `${cardCornerRadius}px ${cardCornerRadius}px 0 0`,
                }}
                aria-hidden="true"
              />
            )}
            <div className={styles.cardHeader}>
              {showFileTypeIcon && (
                <div
                  className={styles.cardIconWrapper}
                  style={{ background: `${effectiveIconColor}18`, borderRadius: cardCornerRadius / 2 }}
                  aria-hidden="true"
                >
                  <span style={{ color: effectiveIconColor, fontSize: itemIconSize }}>
                    <Icon iconName={iconName} />
                  </span>
                </div>
              )}
              <div className={styles.cardTitle} style={{ fontSize: itemFontSize, ...(cardTextColor ? { color: cardTextColor } : {}) }} title={displayTitle}>
                {displayTitle}
              </div>
            </div>

            {showDescription && descriptionText && (
              <div className={styles.cardDescription} style={{ ...(cardTextColor ? { color: cardTextColor, opacity: 0.8 } : {}) }}>
                {truncate(descriptionText, 120)}
              </div>
            )}

            <div className={styles.cardMeta}>
              {[cardMeta1Field, cardMeta2Field].map((fieldName, idx) => {
                const val = resolveMetaValue(item, fieldName);
                if (!val) return null;
                return (
                  <span key={idx} className={styles.cardMetaItem} style={{ ...(cardTextColor ? { color: cardTextColor, opacity: 0.7 } : {}) }}>
                    <Icon iconName={metaIcon(fieldName)} aria-hidden="true" />
                    {val}
                  </span>
                );
              })}
            </div>
          </div>
        );

        if (!isEditMode) {
          return <div key={item.id} role="listitem">{card}</div>;
        }

        return (
          <div key={item.id} role="listitem" className={styles.itemEditWrapper}>
            {card}
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

export default DocumentCardGrid;
