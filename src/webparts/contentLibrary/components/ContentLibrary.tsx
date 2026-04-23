import * as React from 'react';
import { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { formatDate, formatFieldValue } from '../helpers/fieldFormatting';
import { normalizeChoiceValues, getChoiceBadgeStyle } from '../helpers/choiceBadgeUtils';
import { getFileIconInfo, getFileIconColorHex } from '../helpers/fileIconMapping';
import { IContentLibraryProps } from './IContentLibraryProps';
import { SharePointDataService } from '../services/SharePointDataService';
import { ViewMapper } from '../services/ViewMapper';
import { IListItem, IFieldDefinition, IViewInfo } from '../models/IListItem';
import { SortDirection, IItemIconOverride, ICategoryColorMap } from '../models/IWebPartConfig';
import { extractCategories, itemMatchesCategory } from '../helpers/categoryExtraction';
import { filterItemsBySearch, sortItems } from '../helpers/searchUtils';
import { autoAssignCategoryColors, getAccentCssProperties } from '../helpers/colorUtils';
import { resolveItemThumbnailUrl } from '../helpers/thumbnailUtils';
import SearchBar from './SearchBar/SearchBar';
import SortBar from './SortBar/SortBar';
import FilterBar from './FilterBar/FilterBar';
import DocumentTableView from './DocumentTableView/DocumentTableView';
import DocumentCardGrid from './DocumentCardGrid/DocumentCardGrid';
import DocumentTileGrid from './DocumentTileGrid/DocumentTileGrid';
import DashboardView from './DashboardView/DashboardView';
import DocumentPreviewGrid from './DocumentPreviewGrid/DocumentPreviewGrid';
import EmptyState from './EmptyState/EmptyState';
import LoadingState from './LoadingState/LoadingState';
import ItemIconEditor from './ItemIconEditor/ItemIconEditor';
import styles from '../styles/ContentLibrary.module.scss';

const ContentLibrary: React.FC<IContentLibraryProps> = ({ config, context, isEditMode, onIconOverrideSave, onCategoryKeysChange }) => {
  const [allItems, setAllItems] = useState<IListItem[]>([]);
  const [fields, setFields] = useState<IFieldDefinition[]>([]);
  const [viewInfo, setViewInfo] = useState<IViewInfo | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [debouncedQuery, setDebouncedQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState('');
  const [sortField, setSortField] = useState(config.defaultSortField || 'Modified');
  const [sortDirection, setSortDirection] = useState<SortDirection>(config.defaultSortDirection || 'desc');

  // Store the timer ID so we can reliably cancel it
  const debounceTimerId = useRef<ReturnType<typeof setTimeout> | null>(null);
  const dataService = useMemo(() => new SharePointDataService(context), [context]);

  // ── Icon editor state ────────────────────────────────────────────────────
  const [editorOpen, setEditorOpen] = useState(false);
  const [editorItemId, setEditorItemId] = useState('');
  const [editorItemTitle, setEditorItemTitle] = useState('');
  const [editorDefaultIconName, setEditorDefaultIconName] = useState('Document');
  const [editorDefaultIconColor, setEditorDefaultIconColor] = useState('#8a8886');

  // ── List item detail panel state ─────────────────────────────────────────
  const [detailItem, setDetailItem] = useState<IListItem | null>(null);
  /** When thumbnail image errors in details modal, fall back to icon layout */
  const [detailThumbFailed, setDetailThumbFailed] = useState(false);

  // Parse persisted overrides from JSON string
  const itemIconOverrides = useMemo((): Record<string, IItemIconOverride> => {
    try {
      return JSON.parse(config.itemIconOverridesJson || '{}') as Record<string, IItemIconOverride>;
    } catch {
      return {};
    }
  }, [config.itemIconOverridesJson]);

  useEffect(() => {
    setDetailThumbFailed(false);
  }, [detailItem?.id]);

  const accentRootStyle = useMemo(
    (): React.CSSProperties => getAccentCssProperties(config.accentColorHex) as React.CSSProperties,
    [config.accentColorHex]
  );

  const handleEditItemIcon = useCallback((
    itemId: string,
    defaultIconName: string,
    defaultIconColor: string,
    title: string
  ) => {
    setEditorItemId(itemId);
    setEditorItemTitle(title);
    setEditorDefaultIconName(defaultIconName);
    setEditorDefaultIconColor(defaultIconColor);
    setEditorOpen(true);
  }, []);

  const handleIconOverrideSave = useCallback((itemId: string, override: IItemIconOverride | undefined) => {
    onIconOverrideSave(itemId, override);
  }, [onIconOverrideSave]);

  /**
   * Saves a custom thumbnail URL for a Preview card while preserving the
   * existing icon/color settings already stored for that item.
   */
  const handleSaveThumbnail = useCallback((itemId: string, thumbnailUrl: string | undefined) => {
    const existing: IItemIconOverride | undefined = itemIconOverrides[itemId];
    if (!thumbnailUrl) {
      // Remove thumbnail — keep icon/color if present, otherwise clear the override entirely
      if (existing && (existing.iconName || existing.iconColor)) {
        onIconOverrideSave(itemId, { iconName: existing.iconName, iconColor: existing.iconColor });
      } else {
        onIconOverrideSave(itemId, undefined);
      }
    } else {
      onIconOverrideSave(itemId, {
        iconName: existing?.iconName ?? '',
        iconColor: existing?.iconColor ?? '',
        customThumbnailUrl: thumbnailUrl,
      });
    }
  }, [itemIconOverrides, onIconOverrideSave]);

  // ── Item click: open doc in tab or show native DispForm panel ────────────
  const handleItemClick = useCallback((item: IListItem) => {
    if (config.isDocumentLibrary) {
      // fileRef is a server-relative path like /sites/mysite/LibName/file.docx
      // Prepend the tenant origin (protocol + hostname) to make it absolute.
      const fileUrl = item.fileRef
        ? (item.fileRef.startsWith('http')
            ? item.fileRef
            : (() => {
                const abs = context.pageContext.web.absoluteUrl;
                const origin = abs.substring(0, abs.indexOf('/', abs.indexOf('//') + 2));
                return `${origin}${item.fileRef}`;
              })())
        : undefined;
      if (fileUrl) {
        window.open(fileUrl, config.linkTarget === '_blank' ? '_blank' : '_self');
      }
    } else {
      // Open native SharePoint item details panel
      setDetailItem(item);
    }
  }, [config.isDocumentLibrary, config.linkTarget, context.pageContext.web.absoluteUrl]);

  // ── Effect 1: Load fields & views when list/site/view selection changes ──
  // Intentionally separate from the items fetch so that cardMeta field changes
  // can trigger a re-fetch of items without redundantly re-fetching metadata.
  useEffect(() => {
    if (!config.listId) {
      setAllItems([]);
      setFields([]);
      setViewInfo(null);
      return;
    }

    let cancelled = false;

    const loadMeta = async (): Promise<void> => {
      try {
        const [fetchedFields, views] = await Promise.all([
          dataService.getFields(config.listId, config.siteUrl || undefined),
          dataService.getViews(config.listId, config.siteUrl || undefined),
        ]);

        if (cancelled) return;

        setFields(fetchedFields);

        const selectedView = config.viewId
          ? views.find(v => v.id === config.viewId) ?? views[0]
          : views[0];

        setViewInfo(selectedView ?? null);
      } catch (err) {
        if (!cancelled) {
          console.error('[ContentLibrary] Meta load error:', err);
        }
      }
    };

    loadMeta().catch(console.error);
    return () => { cancelled = true; };
  }, [config.listId, config.viewId, config.siteUrl]);

  // ── Effect 2: Fetch items whenever viewInfo or any field-selection changes ─
  // viewInfo is the resolved state from Effect 1, so this always runs with
  // up-to-date viewFields. cardMeta fields are included here so changing the
  // detail-line dropdowns immediately re-fetches with the new columns selected.
  useEffect(() => {
    if (!config.listId) {
      setAllItems([]);
      return;
    }

    let cancelled = false;

    const loadItems = async (): Promise<void> => {
      setLoading(true);
      setError(null);

      try {
        const viewFields = viewInfo?.viewFields ?? [];

        const VIRTUAL_FIELDS = new Set([
          'LinkFilename', 'LinkTitle', 'LinkFilenameNoMenu', 'DocIcon',
          'Edit', 'SelectTitle', '_UIVersionString',
          'Modified', 'Created', 'Editor', 'Author', '',
        ]);
        const extraFields: string[] = [];
        const addExtra = (f: string): void => {
          if (!VIRTUAL_FIELDS.has(f) && extraFields.indexOf(f) === -1) extraFields.push(f);
        };

        if (config.enableFilters && config.filterFieldInternalName) {
          addExtra(config.filterFieldInternalName);
        }
        addExtra(config.cardMeta1Field || '');
        addExtra(config.cardMeta2Field || '');
        viewFields.forEach(f => addExtra(f));

        const items = await dataService.getItems(
          config.listId,
          viewFields,
          config.itemLimit || 100,
          config.isDocumentLibrary,
          config.defaultSortField || 'Modified',
          config.defaultSortDirection !== 'asc',
          config.siteUrl || undefined,
          extraFields
        );

        if (!cancelled) {
          setAllItems(items);
        }
      } catch (err) {
        if (!cancelled) {
          console.error('[ContentLibrary] Items load error:', err);
          setError('Failed to load items. Please check the web part configuration.');
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    loadItems().catch(console.error);
    return () => { cancelled = true; };
  }, [viewInfo, config.listId, config.siteUrl, config.itemLimit, config.isDocumentLibrary, config.defaultSortField, config.defaultSortDirection, config.enableFilters, config.filterFieldInternalName, config.cardMeta1Field, config.cardMeta2Field]);

  // ── Debounce search ──────────────────────────────────────────────────────
  useEffect(() => {
    const delayMs = config.searchDebounceMs > 0 ? config.searchDebounceMs : 300;
    // Cancel any pending timer before scheduling a new one
    if (debounceTimerId.current !== null) {
      clearTimeout(debounceTimerId.current);
    }
    debounceTimerId.current = setTimeout(() => {
      setDebouncedQuery(searchQuery);
      debounceTimerId.current = null;
    }, delayMs);
    return () => {
      if (debounceTimerId.current !== null) {
        clearTimeout(debounceTimerId.current);
        debounceTimerId.current = null;
      }
    };
  }, [searchQuery, config.searchDebounceMs]);

  // ── Reset category when list changes ────────────────────────────────────
  useEffect(() => {
    setSelectedCategory('');
    setSearchQuery('');
    setDebouncedQuery('');
  }, [config.listId]);

  // ── Keep sort state aligned with property-pane defaults ─────────────────
  useEffect(() => {
    setSortField(config.defaultSortField || 'Modified');
    setSortDirection(config.defaultSortDirection || 'desc');
  }, [config.defaultSortField, config.defaultSortDirection]);

  // ── Computed: filtered + searched + sorted items ─────────────────────────
  const filteredItems = useMemo(() => {
    let result = allItems;

    if (config.enableFilters && selectedCategory && config.filterFieldInternalName) {
      result = result.filter(item =>
        itemMatchesCategory(item, config.filterFieldInternalName, selectedCategory)
      );
    }

    if (config.enableSearch && debouncedQuery) {
      const searchFields = config.searchFields?.length
        ? config.searchFields
        : (config.isDocumentLibrary ? ['FileLeafRef', 'Title'] : ['Title']);
      result = filterItemsBySearch(result, debouncedQuery, searchFields);
    }

    result = sortItems(result, sortField, sortDirection === 'asc');

    return result;
  }, [allItems, selectedCategory, debouncedQuery, sortField, sortDirection, config]);

  // ── Computed: categories ─────────────────────────────────────────────────
  const categories = useMemo(() => {
    if (!config.enableFilters || !config.filterFieldInternalName) return [];
    return extractCategories(
      allItems,
      config.filterFieldInternalName,
      config.categorySortOrder,
      config.allOptionLabel
    );
  }, [allItems, config.enableFilters, config.filterFieldInternalName, config.categorySortOrder, config.allOptionLabel]);

  // ── Notify web part when category keys change (for property pane picker) ──
  useEffect(() => {
    if (onCategoryKeysChange) {
      onCategoryKeysChange(categories.map(c => c.key));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [categories]);

  // ── Category colour map ──────────────────────────────────────────────────
  const categoryColors = useMemo((): ICategoryColorMap => {
    if (!config.enableCategoryColors) return {};
    let stored: ICategoryColorMap = {};
    try {
      stored = JSON.parse(config.categoryColorsJson || '{}') as ICategoryColorMap;
    } catch {
      stored = {};
    }
    // Auto-fill any category keys that don't yet have an assigned colour
    const allKeys = categories.map(c => c.key);
    const needsAutoAssign = allKeys.some(k => !stored[k]);
    if (needsAutoAssign) {
      const autoMap = autoAssignCategoryColors(allKeys);
      const merged: ICategoryColorMap = {};
      allKeys.forEach(k => {
        merged[k] = stored[k] || autoMap[k];
      });
      return merged;
    }
    return stored;
  }, [config.enableCategoryColors, config.categoryColorsJson, categories]);

  // ── Category text overrides ──────────────────────────────────────────────
  const categoryTextOverrides = useMemo((): Record<string, 'light' | 'dark'> => {
    if (!config.enableCategoryColors) return {};
    try {
      return JSON.parse(config.categoryTextOverridesJson || '{}') as Record<string, 'light' | 'dark'>;
    } catch {
      return {};
    }
  }, [config.enableCategoryColors, config.categoryTextOverridesJson]);

  // ── Computed: column definitions ─────────────────────────────────────────
  const columnDefs = useMemo(() => {
    if (!viewInfo) return [];
    return ViewMapper.mapViewFields(viewInfo.viewFields, fields);
  }, [viewInfo, fields]);

  // All list fields as IColumnDef — used by card/tile meta resolvers so they can
  // look up the correct fieldType even for columns not in the selected view.
  const allColumnDefs = useMemo(() => {
    return fields.map(f => ({
      internalName: f.internalName,
      displayName: f.displayName,
      fieldType: f.fieldType,
    }));
  }, [fields]);

  // ── Handlers ─────────────────────────────────────────────────────────────
  const handleSearchChange = useCallback((val: string) => {
    setSearchQuery(val);
  }, []);

  const handleCategoryChange = useCallback((key: string) => {
    setSelectedCategory(key);
  }, []);

  const handleSortChange = useCallback((field: string, direction: SortDirection) => {
    setSortField(field);
    setSortDirection(direction);
  }, []);

  const handleTableSortChange = useCallback((field: string, direction: SortDirection) => {
    setSortField(field);
    setSortDirection(direction);
  }, []);

  // ── Not configured state ─────────────────────────────────────────────────
  if (!config.listId) {
    return (
      <div className={styles.contentLibraryRoot} style={accentRootStyle}>
        {config.showTitle && config.webPartTitle && (
          <h2 className={styles.webPartTitle}>{config.webPartTitle}</h2>
        )}
        <div className={styles.notConfigured}>
          <div className={styles.notConfiguredIcon}>📋</div>
          <div className={styles.notConfiguredTitle}>Content Library</div>
          <div className={styles.notConfiguredMessage}>
            {isEditMode
              ? 'Open the property pane to connect this web part to a SharePoint list or document library.'
              : 'This web part has not been configured yet.'}
          </div>
        </div>
      </div>
    );
  }

  // ── Error state ──────────────────────────────────────────────────────────
  if (error) {
    return (
      <div className={styles.contentLibraryRoot} style={accentRootStyle}>
        {config.showTitle && config.webPartTitle && (
          <h2 className={styles.webPartTitle}>{config.webPartTitle}</h2>
        )}
        <div className={styles.emptyState} role="alert">
          <div className={styles.emptyStateIcon}>⚠️</div>
          <div className={styles.emptyStateTitle}>Something went wrong</div>
          <div className={styles.emptyStateMessage}>{error}</div>
        </div>
      </div>
    );
  }

  // ── Determine layout classes ─────────────────────────────────────────────
  // Any filter style can be positioned left/right (not just vertical-rail)
  const isFilterLeft = config.enableFilters && config.filterPosition === 'left';
  const isFilterRight = config.enableFilters && config.filterPosition === 'right';
  const isFilterTop = config.enableFilters && !isFilterLeft && !isFilterRight;

  const layoutClass = isFilterLeft
    ? styles.layoutLeft
    : isFilterRight
    ? styles.layoutRight
    : styles.layoutTop;

  const densityClass = config.density === 'compact'
    ? styles.densityCompact
    : config.density === 'comfortable'
    ? styles.densityComfortable
    : '';

  const shadowClass = config.shadowIntensity === 'none'
    ? styles.shadowNone
    : config.shadowIntensity === 'medium'
    ? styles.shadowMedium
    : config.shadowIntensity === 'strong'
    ? styles.shadowStrong
    : '';

  const rootClass = [
    styles.contentLibraryRoot,
    densityClass,
    shadowClass,
    config.customCssClass,
  ].filter(Boolean).join(' ');

  // ── Sort fields for toolbar ──────────────────────────────────────────────
  const sortFieldOptions = columnDefs.map(c => ({ key: c.internalName, text: c.displayName }));

  // ── Search bar position helpers ──────────────────────────────────────────
  const isSearchTopRight = config.enableSearch
    && config.searchBarStyle !== 'toolbar'
    && config.searchBarPosition === 'top-right';

  const searchWrapperStyle: React.CSSProperties = isSearchTopRight
    ? { display: 'flex', justifyContent: 'flex-end', marginBottom: 16 }
    : { marginBottom: 16 };

  const searchBarInlineStyle: React.CSSProperties = isSearchTopRight
    ? { width: 300, maxWidth: '100%' }
    : {};

  // ── Render search bar ────────────────────────────────────────────────────
  const renderSearchBar = (): React.ReactElement | null => {
    if (!config.enableSearch) return null;
    return (
      <SearchBar
        value={searchQuery}
        onChange={handleSearchChange}
        placeholder={config.searchPlaceholder || 'Search...'}
        style={config.searchBarStyle}
        position={config.searchBarPosition}
        debounceMs={config.searchDebounceMs}
        wrapperStyle={searchBarInlineStyle}
        showSortControl={config.enableSortControl && config.allowUserSort}
        sortFields={sortFieldOptions}
        currentSortField={sortField}
        currentSortAsc={sortDirection === 'asc'}
        onSortChange={(field, asc) => handleSortChange(field, asc ? 'asc' : 'desc')}
      />
    );
  };

  // ── Render filter bar ────────────────────────────────────────────────────
  const renderFilterBar = (): React.ReactElement | null => {
    if (!config.enableFilters || categories.length === 0) return null;
    return (
      <FilterBar
        categories={categories}
        selectedCategory={selectedCategory}
        onCategoryChange={handleCategoryChange}
        style={config.filterStyle}
        showAllOption={config.showAllOption}
        allOptionLabel={config.allOptionLabel || 'All'}
        showCounts={config.showCategoryCounts}
        maxVisible={config.maxVisibleCategories}
        enableCategoryColors={config.enableCategoryColors}
        categoryColors={categoryColors}
        categoryTextOverrides={categoryTextOverrides}
      />
    );
  };

  // ── Render sort bar (standalone, shown when sort is enabled) ──────────────
  const renderSortBar = (): React.ReactElement | null => {
    if (!config.enableSortControl) return null;
    // Toolbar style already embeds sort controls inside the search bar
    if (config.enableSearch && config.searchBarStyle === 'toolbar') return null;
    if (sortFieldOptions.length === 0) return null;
    return (
      <SortBar
        sortField={sortField}
        sortDirection={sortDirection}
        sortFields={sortFieldOptions}
        onSortChange={(field, dir) => handleSortChange(field, dir)}
      />
    );
  };

  // ── Render items ─────────────────────────────────────────────────────────
  const renderItems = (): React.ReactElement => {
    if (loading) {
      return <LoadingState displayStyle={config.itemDisplayStyle} itemCount={6} />;
    }

    if (filteredItems.length === 0) {
      return (
        <EmptyState
          message={config.emptyStateMessage}
          hasSearch={config.enableSearch && !!debouncedQuery}
          hasFilter={config.enableFilters && !!selectedCategory}
        />
      );
    }

    switch (config.itemDisplayStyle) {
      case 'table':
        return (
          <DocumentTableView
            items={filteredItems}
            columns={columnDefs}
            allColumns={allColumnDefs}
            showColumnHeaders={config.showColumnHeaders}
            showFileTypeIcon={config.showFileTypeIcon}
            showModifiedDate={config.showModifiedDate}
            showModifiedBy={config.showModifiedBy}
            isDocumentLibrary={config.isDocumentLibrary}
            linkTarget={config.linkTarget}
            sortField={sortField}
            sortDirection={sortDirection}
            allowUserSort={config.allowUserSort}
            onSortChange={handleTableSortChange}
            onItemClick={handleItemClick}
            isEditMode={isEditMode}
            itemIconOverrides={itemIconOverrides}
            onEditItemIcon={handleEditItemIcon}
            itemFontSize={config.itemFontSize || 13}
            itemIconSize={config.itemIconSize || 24}
          />
        );

      case 'tile-grid':
        return (
          <DocumentTileGrid
            items={filteredItems}
            gridColumns={config.gridColumns || 4}
            showFileTypeIcon={config.showFileTypeIcon}
            isDocumentLibrary={config.isDocumentLibrary}
            linkTarget={config.linkTarget}
            cardCornerRadius={config.cardCornerRadius || 8}
            isEditMode={isEditMode}
            itemIconOverrides={itemIconOverrides}
            onEditItemIcon={handleEditItemIcon}
            onItemClick={handleItemClick}
            cardMeta1Field={config.cardMeta1Field || 'Modified'}
            cardMeta2Field={config.cardMeta2Field || 'Editor'}
            cardMeta1Icon={config.cardMeta1Icon || 'Clock'}
            cardMeta2Icon={config.cardMeta2Icon || 'Contact'}
            showChoicePillsOnCards={config.showChoicePillsOnCards !== false}
            allColumns={allColumnDefs}
            enableCategoryColors={config.enableCategoryColors}
            categoryColors={categoryColors}
            categoryTextOverrides={categoryTextOverrides}
            filterFieldInternalName={config.filterFieldInternalName || ''}
            itemFontSize={config.itemFontSize || 13}
            itemIconSize={config.itemIconSize || 24}
          />
        );

      case 'icon-grid':
        return (
          <DocumentTileGrid
            items={filteredItems}
            gridColumns={config.gridColumns || 5}
            showFileTypeIcon={config.showFileTypeIcon}
            isDocumentLibrary={config.isDocumentLibrary}
            linkTarget={config.linkTarget}
            cardCornerRadius={config.cardCornerRadius || 8}
            isEditMode={isEditMode}
            itemIconOverrides={itemIconOverrides}
            onEditItemIcon={handleEditItemIcon}
            onItemClick={handleItemClick}
            cardMeta1Field={''}
            cardMeta2Field={''}
            cardMeta1Icon={config.cardMeta1Icon || 'Tag'}
            cardMeta2Icon={config.cardMeta2Icon || 'Tag'}
            showChoicePillsOnCards={config.showChoicePillsOnCards !== false}
            allColumns={allColumnDefs}
            enableCategoryColors={config.enableCategoryColors}
            categoryColors={categoryColors}
            categoryTextOverrides={categoryTextOverrides}
            filterFieldInternalName={config.filterFieldInternalName || ''}
            itemFontSize={config.itemFontSize || 13}
            itemIconSize={config.itemIconSize || 24}
          />
        );

      case 'dashboard':
        return (
          <DashboardView
            items={filteredItems}
            showFileTypeIcon={config.showFileTypeIcon}
            showModifiedDate={config.showModifiedDate}
            showModifiedBy={config.showModifiedBy}
            isDocumentLibrary={config.isDocumentLibrary}
            linkTarget={config.linkTarget}
            cardCornerRadius={config.cardCornerRadius || 8}
            isEditMode={isEditMode}
            itemIconOverrides={itemIconOverrides}
            onEditItemIcon={handleEditItemIcon}
            onItemClick={handleItemClick}
          />
        );

      case 'preview':
        return (
          <DocumentPreviewGrid
            items={filteredItems}
            columns={columnDefs}
            gridColumns={config.gridColumns || 3}
            isDocumentLibrary={config.isDocumentLibrary}
            linkTarget={config.linkTarget}
            cardCornerRadius={config.cardCornerRadius || 8}
            isEditMode={isEditMode}
            context={context}
            itemIconOverrides={itemIconOverrides}
            onEditItemIcon={handleEditItemIcon}
            onSaveThumbnail={handleSaveThumbnail}
            onItemClick={handleItemClick}
            cardMeta1Field={config.cardMeta1Field || 'Modified'}
            cardMeta2Field={config.cardMeta2Field || 'Editor'}
            cardMeta1Icon={config.cardMeta1Icon || 'Clock'}
            cardMeta2Icon={config.cardMeta2Icon || 'Contact'}
            showChoicePillsOnCards={config.showChoicePillsOnCards !== false}
            allColumns={allColumnDefs}
            enableCategoryColors={config.enableCategoryColors}
            categoryColors={categoryColors}
            categoryTextOverrides={categoryTextOverrides}
            filterFieldInternalName={config.filterFieldInternalName || ''}
            itemFontSize={config.itemFontSize || 13}
            itemIconSize={config.itemIconSize || 24}
          />
        );

      case 'card-grid':
      default:
        return (
          <DocumentCardGrid
            items={filteredItems}
            columns={columnDefs}
            gridColumns={config.gridColumns || 3}
            showFileTypeIcon={config.showFileTypeIcon}
            showDescription={config.showDescription}
            isDocumentLibrary={config.isDocumentLibrary}
            linkTarget={config.linkTarget}
            cardCornerRadius={config.cardCornerRadius || 8}
            isEditMode={isEditMode}
            itemIconOverrides={itemIconOverrides}
            onEditItemIcon={handleEditItemIcon}
            onItemClick={handleItemClick}
            cardMeta1Field={config.cardMeta1Field || 'Modified'}
            cardMeta2Field={config.cardMeta2Field || 'Editor'}
            cardMeta1Icon={config.cardMeta1Icon || 'Clock'}
            cardMeta2Icon={config.cardMeta2Icon || 'Contact'}
            showChoicePillsOnCards={config.showChoicePillsOnCards !== false}
            allColumns={allColumnDefs}
            enableCategoryColors={config.enableCategoryColors}
            categoryColors={categoryColors}
            categoryTextOverrides={categoryTextOverrides}
            filterFieldInternalName={config.filterFieldInternalName || ''}
            itemFontSize={config.itemFontSize || 13}
            itemIconSize={config.itemIconSize || 24}
          />
        );
    }
  };

  // ── Main render ──────────────────────────────────────────────────────────
  return (
    <div className={rootClass} style={accentRootStyle}>
      {config.showTitle && config.webPartTitle && (
        <h2 className={styles.webPartTitle}>{config.webPartTitle}</h2>
      )}

      {/* Toolbar-style search sits above everything, full width */}
      {config.enableSearch && config.searchBarStyle === 'toolbar' && renderSearchBar()}

      {/* Top search (non-toolbar) — position controls alignment */}
      {config.enableSearch && config.searchBarStyle !== 'toolbar' && (
        <div style={searchWrapperStyle}>
          {renderSearchBar()}
        </div>
      )}

      {/* Standalone sort bar — shown when sort is on but search style is not toolbar */}
      {renderSortBar()}

      {/* Top filter bar (pills, cards, compact — not side rail) */}
      {isFilterTop && (
        <div style={{ marginBottom: 16 }}>
          {renderFilterBar()}
        </div>
      )}

      {/* Side layout wrapper for vertical-rail filter position */}
      {(isFilterLeft || isFilterRight) ? (
        <div className={`${styles.layoutWrapper} ${layoutClass}`}>
          <div className={styles.sidePanel}>
            {renderFilterBar()}
          </div>
          <div className={styles.mainContent}>
            {renderItems()}
          </div>
        </div>
      ) : (
        renderItems()
      )}

      {/* Per-item icon editor panel — only mounted in edit mode */}
      {isEditMode && (
        <ItemIconEditor
          isOpen={editorOpen}
          itemId={editorItemId}
          itemTitle={editorItemTitle}
          currentOverride={editorItemId ? itemIconOverrides[editorItemId] : undefined}
          defaultIconName={editorDefaultIconName}
          defaultIconColor={editorDefaultIconColor}
          onSave={handleIconOverrideSave}
          onDismiss={() => setEditorOpen(false)}
        />
      )}

      {/* List item detail modal — modern popup showing view columns */}
      {!config.isDocumentLibrary && detailItem && (
        <Modal
          isOpen={!!detailItem}
          onDismiss={() => setDetailItem(null)}
          isBlocking={false}
          styles={{
            main: {
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
            },
          }}
        >
          <div className={styles.detailModal}>
            <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '8px 8px 0 8px' }}>
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title="Close"
                ariaLabel="Close details"
                onClick={() => setDetailItem(null)}
              />
            </div>
          {(() => {
            const override = itemIconOverrides[detailItem.id];
            const defaultIconInfo = getFileIconInfo(detailItem.fileType, detailItem.isFolder);
            const defaultIconColor = getFileIconColorHex(detailItem.fileType, detailItem.isFolder);
            const iconName = (override?.iconName && String(override.iconName).trim())
              ? override.iconName
              : defaultIconInfo.iconName;
            const iconColor = (override?.iconColor && String(override.iconColor).trim())
              ? override.iconColor
              : defaultIconColor;

            const thumbUrl = resolveItemThumbnailUrl(detailItem, override);
            const useThumbInHero = (config.detailsShowThumbnail !== false) && !!thumbUrl && !detailThumbFailed;
            const layoutAbove = config.detailsThumbnailLayout === 'above';
            const heroTitleText = detailItem.title || 'Item Details';

            const heroHeader = ((): React.ReactElement => {
              if (useThumbInHero && layoutAbove) {
                return (
                  <div className={styles.detailPanelHeroStack}>
                    <div className={styles.detailPanelHeroThumbAbove}>
                      <img
                        src={thumbUrl}
                        alt=""
                        onError={() => setDetailThumbFailed(true)}
                      />
                    </div>
                    <div className={styles.detailPanelHeroTitleAbove}>{heroTitleText}</div>
                  </div>
                );
              }
              if (useThumbInHero && !layoutAbove) {
                return (
                  <div className={styles.detailPanelHeroRow}>
                    <div className={styles.detailPanelHeroThumb}>
                      <img
                        src={thumbUrl}
                        alt=""
                        onError={() => setDetailThumbFailed(true)}
                      />
                    </div>
                    <div className={styles.detailPanelHeroTitleLeft}>{heroTitleText}</div>
                  </div>
                );
              }
              return (
                <div className={styles.detailPanelHero}>
                  <div
                    className={styles.detailPanelHeroIcon}
                    style={{ background: `${iconColor}18` }}
                    aria-hidden="true"
                  >
                    <Icon iconName={iconName} style={{ color: iconColor }} />
                  </div>
                  <div className={styles.detailPanelHeroTitle}>{heroTitleText}</div>
                </div>
              );
            })();

            // Build a lookup for field type resolution from all list fields
            const colTypeMap: Record<string, string> = {};
            const colLabelMap: Record<string, string> = {};
            allColumnDefs.forEach(c => {
              colTypeMap[c.internalName] = c.fieldType;
              colLabelMap[c.internalName] = c.displayName;
            });

            const SKIP = new Set([
              'Title', 'FileLeafRef', 'LinkFilename', 'LinkTitle', 'LinkFilenameNoMenu',
              'DocIcon', 'Edit', 'SelectTitle', '_UIVersionString', 'ContentTypeId',
              'ContentType', 'ComplianceAssetId', 'ID', 'GUID', 'UniqueId',
              'FSObjType', 'SMTotalSize', 'File_x0020_Type',
            ]);

            const rows: Array<{ label: string; value: string; internalName: string }> = [];
            const shownNames = new Set<string>();

            // Helper: resolve a raw field value to a display string
            const resolveValue = (internalName: string, raw: unknown): string => {
              if (raw === undefined || raw === null || raw === '') return '';
              const ft = colTypeMap[internalName] ?? 'Text';
              // Handle Editor/Author expanded objects
              if (typeof raw === 'object' && raw !== null) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const obj = raw as any;
                if (obj.Title) return String(obj.Title);
                if (obj.LookupValue) return String(obj.LookupValue);
              }
              return formatFieldValue(raw, ft);
            };

            // 1. Walk view columns in order — these are the fields the author chose to show
            columnDefs.forEach(col => {
              if (SKIP.has(col.internalName)) return;
              shownNames.add(col.internalName);
              // Built-in mapped fields
              if (col.internalName === 'Modified' || col.internalName === 'Editor') return; // handled below
              const raw = detailItem[col.internalName];
              const val = resolveValue(col.internalName, raw);
              if (!val) return;
              rows.push({ label: col.displayName, value: val, internalName: col.internalName });
            });

            // 2. Always append Modified and Modified By (they are almost always in the view)
            if (detailItem.modified) {
              rows.push({ label: 'Modified', value: formatDate(detailItem.modified), internalName: 'Modified' });
              shownNames.add('Modified');
            }
            if (detailItem.modifiedBy) {
              rows.push({ label: 'Modified by', value: detailItem.modifiedBy, internalName: 'Editor' });
              shownNames.add('Editor');
            }

            return (
              <>
                {heroHeader}

                {/* Field rows */}
                <div className={styles.detailPanelBody}>
                  <div className={styles.detailPanelSection}>
                    {rows.map(row => {
                      const isDate = row.internalName === 'Modified' || row.internalName === 'Created' || colTypeMap[row.internalName] === 'DateTime';
                      const isUser = row.internalName === 'Editor' || row.internalName === 'Author' || colTypeMap[row.internalName] === 'User';
                      const isChoice = colTypeMap[row.internalName] === 'Choice' || colTypeMap[row.internalName] === 'MultiChoice';
                      const choiceValues = isChoice ? normalizeChoiceValues(detailItem[row.internalName]) : [];
                      return (
                        <div key={row.internalName} className={styles.detailPanelRow}>
                          <div className={styles.detailPanelLabel}>{row.label}</div>
                          <div className={isDate || isUser ? styles.detailPanelValueMuted : styles.detailPanelValue}>
                            {isDate && <Icon iconName="Clock" style={{ fontSize: 12 }} aria-hidden="true" />}
                            {isUser && <Icon iconName="Contact" style={{ fontSize: 12 }} aria-hidden="true" />}
                            {isChoice && choiceValues.length > 0 ? (
                              <span className={styles.detailChoicePills}>
                                {choiceValues.map(choice => {
                                  const pillStyle = getChoiceBadgeStyle(choice);
                                  return (
                                    <span key={choice} className={styles.choiceBadge} style={pillStyle}>
                                      {choice}
                                    </span>
                                  );
                                })}
                              </span>
                            ) : row.value}
                          </div>
                        </div>
                      );
                    })}
                    {rows.length === 0 && (
                      <div style={{ padding: '24px 0', color: '#605e5c', fontSize: 14, textAlign: 'center' }}>
                        No field data available for this item.
                      </div>
                    )}
                  </div>
                </div>

                {/* Footer with metadata */}
                {(detailItem.created || detailItem.createdBy) && (
                  <div className={styles.detailPanelFooter}>
                    <Icon iconName="Info" style={{ fontSize: 12 }} aria-hidden="true" />
                    {detailItem.created && <>Created {formatDate(detailItem.created)}</>}
                    {detailItem.createdBy && <> &middot; {detailItem.createdBy}</>}
                  </div>
                )}
              </>
            );
          })()}
          </div>
        </Modal>
      )}
    </div>
  );
};

export default ContentLibrary;
