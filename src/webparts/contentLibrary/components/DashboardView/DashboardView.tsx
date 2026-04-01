import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IListItem } from '../../models/IListItem';
import { IItemIconOverride } from '../../models/IWebPartConfig';
import { formatDate } from '../../helpers/fieldFormatting';
import { getFileIconInfo, getFileIconColorHex } from '../../helpers/fileIconMapping';
import { LinkTarget } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

export interface IDashboardViewProps {
  items: IListItem[];
  showFileTypeIcon: boolean;
  showModifiedDate: boolean;
  showModifiedBy: boolean;
  isDocumentLibrary: boolean;
  linkTarget: LinkTarget;
  cardCornerRadius: number;
  isEditMode: boolean;
  itemIconOverrides: Record<string, IItemIconOverride>;
  onEditItemIcon: (itemId: string, defaultIconName: string, defaultIconColor: string, title: string) => void;
  onItemClick: (item: IListItem) => void;
}

const DashboardView: React.FC<IDashboardViewProps> = ({
  items,
  showFileTypeIcon,
  showModifiedDate,
  showModifiedBy,
  isDocumentLibrary,
  linkTarget,
  cardCornerRadius,
  isEditMode,
  itemIconOverrides,
  onEditItemIcon,
  onItemClick,
}) => {
  const [starredIds, setStarredIds] = useState<string[]>([]);

  const toggleStar = (id: string, e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    setStarredIds(prev => {
      if (prev.indexOf(id) !== -1) return prev.filter(x => x !== id);
      return [...prev, id];
    });
  };

  const recentItems = items.slice(0, 8);
  const starredItems = items.filter(i => starredIds.indexOf(i.id) !== -1);

  const renderDocRow = (item: IListItem): React.ReactElement => {
    const defaultIconInfo = getFileIconInfo(item.fileType, item.isFolder);
    const defaultIconColor = getFileIconColorHex(item.fileType, item.isFolder);
    const override = itemIconOverrides[item.id];
    const iconName = override?.iconName ?? defaultIconInfo.iconName;
    const iconColor = override?.iconColor ?? defaultIconColor;

    const displayTitle = isDocumentLibrary
      ? (item.fileLeafRef ?? item.title)
      : item.title;
    const isStarred = starredIds.indexOf(item.id) !== -1;

    return (
      <div key={item.id} className={styles.itemEditWrapper} style={{ display: 'block' }}>
        <div
          role="button"
          tabIndex={0}
          className={styles.dashboardDocumentRow}
          aria-label={displayTitle}
          style={{ cursor: 'pointer' }}
          onClick={() => onItemClick(item)}
          onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && onItemClick(item)}
        >
          {showFileTypeIcon && (
            <span style={{ color: iconColor, fontSize: 20, flexShrink: 0 }} aria-hidden="true">
              <Icon iconName={iconName} />
            </span>
          )}
          <div className={styles.dashboardDocTitle} title={displayTitle}>
            {displayTitle}
          </div>
          {showModifiedBy && item.modifiedBy && (
            <div className={styles.dashboardDocMeta}>{item.modifiedBy}</div>
          )}
          {showModifiedDate && item.modified && (
            <div className={styles.dashboardDocMeta}>{formatDate(item.modified)}</div>
          )}
          <button
            className={`${styles.dashboardDocStar} ${isStarred ? styles.starred : ''}`}
            onClick={e => toggleStar(item.id, e)}
            aria-label={isStarred ? `Unstar ${displayTitle}` : `Star ${displayTitle}`}
            aria-pressed={isStarred}
            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}
          >
            <Icon iconName={isStarred ? 'FavoriteStarFill' : 'FavoriteStar'} />
          </button>
        </div>
        {isEditMode && (
          <div className={styles.itemEditOverlay}>
            <button
              className={styles.itemEditButton}
              onClick={() => onEditItemIcon(item.id, defaultIconInfo.iconName, defaultIconColor, displayTitle)}
              title="Edit icon"
              aria-label={`Edit icon for ${displayTitle}`}
            >
              <Icon iconName="Color" style={{ fontSize: 13 }} />
            </button>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={styles.dashboardLayout}>
      <div className={styles.dashboardSection} style={{ borderRadius: cardCornerRadius }}>
        <div className={styles.dashboardSectionHeader}>
          <div className={styles.dashboardSectionTitle}>Recent Documents</div>
          {items.length > 8 && (
            <span className={styles.dashboardSeeAll} aria-label="See all documents">See all</span>
          )}
        </div>
        <div className={styles.dashboardDocumentList} role="list" aria-label="Recent documents">
          {recentItems.map(item => renderDocRow(item))}
        </div>
      </div>

      <div className={styles.dashboardSection} style={{ borderRadius: cardCornerRadius }}>
        <div className={styles.dashboardSectionHeader}>
          <div className={styles.dashboardSectionTitle}>Favourite Documents</div>
        </div>
        <div className={styles.dashboardDocumentList} role="list" aria-label="Favourite documents">
          {starredItems.length === 0 ? (
            <div style={{ padding: '24px 20px', textAlign: 'center', color: '#605e5c', fontSize: 13 }}>
              <Icon iconName="FavoriteStar" style={{ fontSize: 24, display: 'block', margin: '0 auto 8px', color: '#c8c6c4' }} />
              Star documents to add them here
            </div>
          ) : (
            starredItems.map(item => renderDocRow(item))
          )}
        </div>
      </div>
    </div>
  );
};

export default DashboardView;
