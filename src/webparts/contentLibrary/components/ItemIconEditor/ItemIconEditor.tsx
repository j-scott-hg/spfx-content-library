import * as React from 'react';
import { useState, useCallback, useMemo } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { DefaultButton, PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { IItemIconOverride } from '../../models/IWebPartConfig';
import styles from '../../styles/ContentLibrary.module.scss';

// ─── Curated icon list ────────────────────────────────────────────────────────
// A broad set of Fluent UI icons grouped by category for easy browsing.
const ICON_GROUPS: Array<{ label: string; icons: string[] }> = [
  {
    label: 'Documents & Files',
    icons: [
      'Document', 'WordDocument', 'ExcelDocument', 'PowerPointDocument', 'PDF',
      'OneNoteLogo', 'VisioDocument', 'TextDocument', 'FileCode', 'FileImage',
      'ZipFolder', 'FabricFolder', 'FabricOpenFolderHorizontal', 'Attach',
    ],
  },
  {
    label: 'Actions & Navigation',
    icons: [
      'Link', 'Globe', 'OpenInNewTab', 'NavigateExternalInline', 'Share',
      'Download', 'Upload', 'Send', 'Forward', 'Reply', 'Pin', 'Pinned',
      'Bookmark', 'BookmarkSolid', 'Tag', 'TagSolid',
    ],
  },
  {
    label: 'Communication',
    icons: [
      'Mail', 'MailSolid', 'Chat', 'Comment', 'CommentSolid', 'Phone',
      'Video', 'VideoSolid', 'Megaphone', 'Bullhorn', 'Ringer', 'RingerSolid',
    ],
  },
  {
    label: 'People & Org',
    icons: [
      'Contact', 'ContactSolid', 'People', 'Group', 'Team', 'Org',
      'ReminderPerson', 'UserEvent', 'FollowUser', 'AddFriend',
    ],
  },
  {
    label: 'Data & Analytics',
    icons: [
      'BarChart4', 'LineChart', 'PieDouble', 'AreaChart', 'ScatterChart',
      'BullseyeTarget', 'Trending12', 'NumberField', 'Calculator', 'Database',
      'Table', 'GridViewSmall', 'List', 'BulletedList',
    ],
  },
  {
    label: 'Settings & Tools',
    icons: [
      'Settings', 'SettingsSolid', 'Edit', 'EditSolid', 'Delete', 'Add',
      'Search', 'Filter', 'Sort', 'Refresh', 'Sync', 'Build', 'Repair',
      'Lock', 'Unlock', 'Shield', 'ShieldSolid',
    ],
  },
  {
    label: 'Status & Alerts',
    icons: [
      'Completed', 'CompletedSolid', 'StatusCircleCheckmark', 'Warning',
      'WarningSolid', 'Error', 'ErrorBadge', 'Info', 'InfoSolid',
      'Flag', 'FlagSolid', 'Like', 'LikeSolid', 'Dislike', 'DislikeSolid',
    ],
  },
  {
    label: 'Stars & Symbols',
    icons: [
      'FavoriteStar', 'FavoriteStarFill', 'Heart', 'HeartFill',
      'Diamond', 'Ribbon', 'Trophy', 'Trophy2', 'Medal', 'Crown',
      'Lightbulb', 'LightbulbSolid', 'Rocket', 'AsteriskSolid',
    ],
  },
  {
    label: 'Media & Content',
    icons: [
      'Photo2', 'Camera', 'Video', 'MusicNote', 'Play', 'Pause',
      'PageAdd', 'Page', 'News', 'Blog', 'Library', 'ReadingMode',
      'Education', 'Certificate', 'Presentation',
    ],
  },
  {
    label: 'Location & Time',
    icons: [
      'MapPin', 'Location', 'World', 'Globe', 'Calendar', 'CalendarDay',
      'Clock', 'Timer', 'DateTime', 'ScheduleEventAction', 'Arrivals',
    ],
  },
];

// Flatten for search
const ALL_ICONS: string[] = ICON_GROUPS.reduce<string[]>(
  (acc, g) => [...acc, ...g.icons], []
);

// Preset colour swatches
const COLOR_SWATCHES = [
  '#0078d4', '#106ebe', '#185abd', '#107c41', '#217346',
  '#c43e1c', '#d93025', '#7719aa', '#038387', '#881798',
  '#ffb900', '#ff8c00', '#e3008c', '#00b294', '#8a8886',
  '#201f1e', '#323130', '#605e5c',
];

export interface IItemIconEditorProps {
  isOpen: boolean;
  itemTitle: string;
  itemId: string;
  currentOverride: IItemIconOverride | undefined;
  defaultIconName: string;
  defaultIconColor: string;
  onSave: (itemId: string, override: IItemIconOverride | undefined) => void;
  onDismiss: () => void;
}

const ItemIconEditor: React.FC<IItemIconEditorProps> = ({
  isOpen,
  itemTitle,
  itemId,
  currentOverride,
  defaultIconName,
  defaultIconColor,
  onSave,
  onDismiss,
}) => {
  const [selectedIcon, setSelectedIcon] = useState<string>(
    currentOverride?.iconName ?? defaultIconName
  );
  const [selectedColor, setSelectedColor] = useState<string>(
    currentOverride?.iconColor ?? defaultIconColor
  );
  const [searchQuery, setSearchQuery] = useState('');
  const [customColor, setCustomColor] = useState(
    currentOverride?.iconColor ?? defaultIconColor
  );

  // Reset local state whenever the panel opens for a (possibly different) item
  React.useEffect(() => {
    if (isOpen) {
      setSelectedIcon(currentOverride?.iconName ?? defaultIconName);
      setSelectedColor(currentOverride?.iconColor ?? defaultIconColor);
      setCustomColor(currentOverride?.iconColor ?? defaultIconColor);
      setSearchQuery('');
    }
  }, [isOpen, itemId, currentOverride, defaultIconName, defaultIconColor]);

  const filteredIcons = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return null; // show groups when no search
    return ALL_ICONS.filter(name => name.toLowerCase().indexOf(q) !== -1);
  }, [searchQuery]);

  const handleSave = useCallback(() => {
    onSave(itemId, { iconName: selectedIcon, iconColor: selectedColor });
    onDismiss();
  }, [itemId, selectedIcon, selectedColor, onSave, onDismiss]);

  const handleReset = useCallback(() => {
    onSave(itemId, undefined);
    onDismiss();
  }, [itemId, onSave, onDismiss]);

  const handleCustomColorChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setCustomColor(val);
    // Only apply if it looks like a valid hex
    if (/^#[0-9a-fA-F]{6}$/.test(val)) {
      setSelectedColor(val);
    }
  }, []);

  const renderIconGrid = (icons: string[]): React.ReactElement => (
    <div className={styles.iconPickerGrid}>
      {icons.map(name => (
        <button
          key={name}
          className={`${styles.iconPickerCell} ${selectedIcon === name ? styles.iconPickerCellActive : ''}`}
          onClick={() => setSelectedIcon(name)}
          title={name}
          aria-label={name}
          aria-pressed={selectedIcon === name}
        >
          <Icon iconName={name} style={{ fontSize: 20, color: selectedIcon === name ? selectedColor : '#323130' }} />
        </button>
      ))}
    </div>
  );

  const panelFooter = (): React.ReactElement => (
    <div className={styles.iconEditorFooter}>
      <PrimaryButton text="Apply" onClick={handleSave} />
      <DefaultButton text="Cancel" onClick={onDismiss} style={{ marginLeft: 8 }} />
      {currentOverride && (
        <DefaultButton
          text="Reset to default"
          onClick={handleReset}
          style={{ marginLeft: 'auto' }}
          iconProps={{ iconName: 'Refresh' }}
        />
      )}
    </div>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.smallFixedFar}
      headerText="Edit icon"
      isFooterAtBottom
      onRenderFooterContent={panelFooter}
      closeButtonAriaLabel="Close icon editor"
      styles={{
        main: { maxWidth: 340 },
        content: { paddingBottom: 0 },
      }}
    >
      <div className={styles.iconEditorBody}>
        {/* Preview */}
        <div className={styles.iconEditorPreview}>
          <div
            className={styles.iconEditorPreviewBadge}
            style={{ background: `${selectedColor}18` }}
          >
            <Icon iconName={selectedIcon} style={{ fontSize: 32, color: selectedColor }} />
          </div>
          <div className={styles.iconEditorPreviewTitle} title={itemTitle}>
            {itemTitle}
          </div>
        </div>

        {/* Colour picker */}
        <div className={styles.iconEditorSection}>
          <div className={styles.iconEditorSectionLabel}>Colour</div>
          <div className={styles.colorSwatches}>
            {COLOR_SWATCHES.map(hex => (
              <button
                key={hex}
                className={`${styles.colorSwatch} ${selectedColor === hex ? styles.colorSwatchActive : ''}`}
                style={{ background: hex }}
                onClick={() => { setSelectedColor(hex); setCustomColor(hex); }}
                aria-label={`Colour ${hex}`}
                aria-pressed={selectedColor === hex}
                title={hex}
              />
            ))}
          </div>
          <div className={styles.colorCustomRow}>
            <input
              type="color"
              value={customColor.startsWith('#') && customColor.length === 7 ? customColor : '#0078d4'}
              onChange={e => { setCustomColor(e.target.value); setSelectedColor(e.target.value); }}
              className={styles.colorNativePicker}
              aria-label="Custom colour"
              title="Pick a custom colour"
            />
            <input
              type="text"
              value={customColor}
              onChange={handleCustomColorChange}
              placeholder="#0078d4"
              className={styles.colorHexInput}
              aria-label="Hex colour value"
              maxLength={7}
            />
          </div>
        </div>

        {/* Icon search */}
        <div className={styles.iconEditorSection}>
          <div className={styles.iconEditorSectionLabel}>Icon</div>
          <div className={styles.iconSearchWrapper}>
            <Icon iconName="Search" className={styles.iconSearchIcon} aria-hidden="true" />
            <input
              type="search"
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              placeholder="Search icons…"
              className={styles.iconSearchInput}
              aria-label="Search icons"
            />
            {searchQuery && (
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                onClick={() => setSearchQuery('')}
                styles={{ root: { width: 24, height: 24, padding: 0 }, icon: { fontSize: 11 } }}
                ariaLabel="Clear search"
              />
            )}
          </div>
        </div>

        {/* Icon grid */}
        <div className={styles.iconPickerScroll}>
          {filteredIcons ? (
            filteredIcons.length > 0
              ? renderIconGrid(filteredIcons)
              : <div className={styles.iconPickerEmpty}>No icons match &ldquo;{searchQuery}&rdquo;</div>
          ) : (
            ICON_GROUPS.map(group => (
              <div key={group.label} className={styles.iconPickerGroup}>
                <div className={styles.iconPickerGroupLabel}>{group.label}</div>
                {renderIconGrid(group.icons)}
              </div>
            ))
          )}
        </div>
      </div>
    </Panel>
  );
};

export default ItemIconEditor;
