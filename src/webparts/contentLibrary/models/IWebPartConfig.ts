// ─── Category colour coding ───────────────────────────────────────────────────

/**
 * Maps a category key (the raw field value) to a hex background colour.
 * Text and icon colour are derived automatically for contrast.
 * Stored as a JSON string in itemIconOverridesJson's sibling property.
 */
export type ICategoryColorMap = Record<string, string>;

// ─── Per-item icon override ───────────────────────────────────────────────────

export interface IItemIconOverride {
  /** Fluent UI icon name, e.g. "Globe", "Star", "Heart" */
  iconName: string;
  /** Hex colour string, e.g. "#0078d4" */
  iconColor: string;
}

// ─── Enumerations ────────────────────────────────────────────────────────────

export type SearchBarStyle = 'minimal' | 'elevated' | 'toolbar';
export type SearchBarPosition = 'top-full' | 'top-right' | 'toolbar';

export type FilterStyle = 'pills' | 'vertical-rail' | 'cards' | 'compact-buttons';
export type FilterPosition = 'top' | 'left' | 'right';
export type CategorySortOrder = 'alpha' | 'count';

export type ItemDisplayStyle = 'table' | 'card-grid' | 'tile-grid' | 'dashboard' | 'icon-grid';
export type Density = 'compact' | 'normal' | 'comfortable';
export type ShadowIntensity = 'none' | 'subtle' | 'medium' | 'strong';
export type LinkTarget = '_self' | '_blank';
export type SortDirection = 'asc' | 'desc';

// ─── Main Config Interface ────────────────────────────────────────────────────

export interface IWebPartConfig {
  // ── Data Source ────────────────────────────────────────────────────────────
  webPartTitle: string;
  showTitle: boolean;
  siteUrl: string;
  listId: string;
  listTitle: string;
  viewId: string;
  viewTitle: string;
  itemLimit: number;
  isDocumentLibrary: boolean;

  // ── Search ─────────────────────────────────────────────────────────────────
  enableSearch: boolean;
  searchPlaceholder: string;
  searchFields: string[]; // internal field names to search across
  searchBarStyle: SearchBarStyle;
  searchBarPosition: SearchBarPosition;
  searchDebounceMs: number;

  // ── Filters ────────────────────────────────────────────────────────────────
  enableFilters: boolean;
  filterFieldInternalName: string;
  filterFieldDisplayName: string;
  filterStyle: FilterStyle;
  filterPosition: FilterPosition;
  showAllOption: boolean;
  allOptionLabel: string;
  showCategoryCounts: boolean;
  categorySortOrder: CategorySortOrder;
  maxVisibleCategories: number;

  // ── Display ────────────────────────────────────────────────────────────────
  itemDisplayStyle: ItemDisplayStyle;
  density: Density;
  cardCornerRadius: number; // px
  shadowIntensity: ShadowIntensity;
  showFileTypeIcon: boolean;
  showModifiedDate: boolean;
  showModifiedBy: boolean;
  showDescription: boolean;
  showColumnHeaders: boolean;
  gridColumns: number; // for card/tile modes
  // Internal field name to show as the first meta line on cards/tiles (empty = hide)
  cardMeta1Field: string;
  // Internal field name to show as the second meta line on cards/tiles (empty = hide)
  cardMeta2Field: string;

  // ── Sorting ────────────────────────────────────────────────────────────────
  enableSortControl: boolean;
  defaultSortField: string;
  defaultSortDirection: SortDirection;
  allowUserSort: boolean;
  linkTarget: LinkTarget;

  // ── Advanced ───────────────────────────────────────────────────────────────
  emptyStateMessage: string;
  loadingText: string;
  customCssClass: string;

  // ── Per-item icon overrides ────────────────────────────────────────────────
  // Stored as a JSON string (Record<itemId, IItemIconOverride>) to survive
  // property pane serialisation. Parse with JSON.parse before use.
  itemIconOverridesJson: string;

  // ── Category colour coding ─────────────────────────────────────────────────
  enableCategoryColors: boolean;
  // JSON string of ICategoryColorMap (Record<categoryKey, hexColor>)
  categoryColorsJson: string;
  // JSON string of Record<categoryKey, 'light'|'dark'> — overrides auto-contrast text colour
  categoryTextOverridesJson: string;

  // ── Font & icon sizing ─────────────────────────────────────────────────────
  itemFontSize: number;   // px, e.g. 13
  itemIconSize: number;   // px, e.g. 24
}

export const DEFAULT_CONFIG: IWebPartConfig = {
  webPartTitle: 'Content Library',
  showTitle: true,
  siteUrl: '',
  listId: '',
  listTitle: '',
  viewId: '',
  viewTitle: '',
  itemLimit: 50,
  isDocumentLibrary: false,

  enableSearch: true,
  searchPlaceholder: 'Search...',
  searchFields: ['Title', 'FileLeafRef'],
  searchBarStyle: 'minimal',
  searchBarPosition: 'top-full',
  searchDebounceMs: 300,

  enableFilters: false,
  filterFieldInternalName: '',
  filterFieldDisplayName: '',
  filterStyle: 'pills',
  filterPosition: 'top',
  showAllOption: true,
  allOptionLabel: 'All',
  showCategoryCounts: true,
  categorySortOrder: 'alpha',
  maxVisibleCategories: 10,

  itemDisplayStyle: 'card-grid',
  density: 'normal',
  cardCornerRadius: 8,
  shadowIntensity: 'subtle',
  showFileTypeIcon: true,
  showModifiedDate: true,
  showModifiedBy: true,
  showDescription: true,
  showColumnHeaders: true,
  gridColumns: 3,
  cardMeta1Field: 'Modified',
  cardMeta2Field: 'Editor',

  enableSortControl: false,
  defaultSortField: 'Modified',
  defaultSortDirection: 'desc',
  allowUserSort: true,
  linkTarget: '_blank',

  emptyStateMessage: 'No items found.',
  loadingText: 'Loading...',
  customCssClass: '',
  itemIconOverridesJson: '{}',
  enableCategoryColors: false,
  categoryColorsJson: '{}',
  categoryTextOverridesJson: '{}',
  itemFontSize: 13,
  itemIconSize: 24,
};
