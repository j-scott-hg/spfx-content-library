import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneLabel,
  PropertyPaneHorizontalRule,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  IPropertyPaneCustomFieldProps,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import ContentLibrary from './components/ContentLibrary';
import { IContentLibraryProps } from './components/IContentLibraryProps';
import { IWebPartConfig, DEFAULT_CONFIG, IItemIconOverride } from './models/IWebPartConfig';
import { SharePointDataService } from './services/SharePointDataService';
import { IListInfo, IViewInfo, IFieldDefinition } from './models/IListItem';
import { autoAssignCategoryColors } from './helpers/colorUtils';
import CategoryColorPicker from './components/CategoryColorPicker/CategoryColorPicker';

export interface IContentLibraryWebPartProps extends IWebPartConfig {
  // All config fields are stored directly on properties
}

export default class ContentLibraryWebPart extends BaseClientSideWebPart<IContentLibraryWebPartProps> {
  private _metaIconOptions: IPropertyPaneDropdownOption[] = [
    { key: 'Tag', text: 'Tag' },
    { key: 'Info', text: 'Info' },
    { key: 'Clock', text: 'Clock' },
    { key: 'Contact', text: 'Person' },
    { key: 'Calendar', text: 'Calendar' },
    { key: 'CheckMark', text: 'Check mark' },
    { key: 'Link', text: 'Link' },
    { key: 'Attach', text: 'Attachment' },
    { key: 'Document', text: 'Document' },
    { key: 'Globe', text: 'Globe' },
    { key: 'Starburst', text: 'Star' },
    { key: 'StatusCircleCheckmark', text: 'Status' },
  ];


  private _dataService!: SharePointDataService;
  private _lists: IListInfo[] = [];
  private _views: IViewInfo[] = [];
  private _fields: IFieldDefinition[] = [];
  private _listsLoading = false;
  private _viewsLoading = false;
  private _fieldsLoading = false;
  /** Category keys extracted from loaded items — used by the colour picker */
  private _categoryKeys: string[] = [];

  // ── Lifecycle ─────────────────────────────────────────────────────────────

  protected async onInit(): Promise<void> {
    // Register all Fluent UI MDL2 icons so every icon name in the picker resolves correctly.
    // The `disableWarnings` option suppresses re-registration warnings from SharePoint's own call.
    initializeIcons(undefined, { disableWarnings: true });

    this._dataService = new SharePointDataService(this.context);

    // Apply defaults for any missing properties
    const defaults = DEFAULT_CONFIG;
    for (const key of Object.keys(defaults) as Array<keyof IWebPartConfig>) {
      if (this.properties[key] === undefined || this.properties[key] === null) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (this.properties as any)[key] = defaults[key];
      }
    }

    // Pre-load lists for the current site
    await this._loadLists();

    if (this.properties.listId) {
      await Promise.all([
        this._loadViews(this.properties.listId),
        this._loadFields(this.properties.listId),
      ]);
    }
  }

  public render(): void {
    const config: IWebPartConfig = { ...DEFAULT_CONFIG, ...this.properties };

    const element: React.ReactElement<IContentLibraryProps> = React.createElement(
      ContentLibrary,
      {
        config,
        context: this.context,
        isEditMode: this.displayMode === 2, // DisplayMode.Edit = 2
        onIconOverrideSave: this._handleIconOverrideSave.bind(this),
        onCategoryKeysChange: this.updateCategoryKeys.bind(this),
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Persists a per-item icon override into the web part properties JSON string.
   * Passing null removes the override for that item (resets to default).
   */
  private _handleIconOverrideSave(itemId: string, override: IItemIconOverride | undefined): void {
    let overrides: Record<string, IItemIconOverride> = {};
    try {
      overrides = JSON.parse(this.properties.itemIconOverridesJson || '{}') as Record<string, IItemIconOverride>;
    } catch {
      overrides = {};
    }

    if (override === undefined) {
      delete overrides[itemId];
    } else {
      overrides[itemId] = override;
    }

    this.properties.itemIconOverridesJson = JSON.stringify(overrides);
    // Re-render to reflect the saved override immediately
    this.render();
  }

  /**
   * Called by the CategoryColorPicker custom field whenever a colour or text
   * override changes. Writes both maps back to properties and refreshes.
   */
  private _handleCategoryColorsChange(
    updatedColors: Record<string, string>,
    updatedTextOverrides: Record<string, 'light' | 'dark'>
  ): void {
    this.properties.categoryColorsJson = JSON.stringify(updatedColors);
    this.properties.categoryTextOverridesJson = JSON.stringify(updatedTextOverrides);
    this.render();
    this.context.propertyPane.refresh();
  }

  /**
   * Called from ContentLibrary component (via onCategoryKeysChange) whenever
   * the set of live category keys changes, so the colour picker stays in sync.
   */
  public updateCategoryKeys(keys: string[]): void {
    const changed = keys.join(',') !== this._categoryKeys.join(',');
    this._categoryKeys = keys;
    if (changed && this.context.propertyPane.isRenderedByWebPart()) {
      this.context.propertyPane.refresh();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  // ── Property Pane Change Handling ─────────────────────────────────────────

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'siteUrl' && newValue !== oldValue) {
      this.properties.listId = '';
      this.properties.listTitle = '';
      this.properties.viewId = '';
      this.properties.viewTitle = '';
      this._views = [];
      this._fields = [];
      await this._loadLists();
      this.context.propertyPane.refresh();
      this.render();
    }

    if (propertyPath === 'listId' && newValue !== oldValue) {
      const selectedList = this._lists.find(l => l.id === String(newValue));
      this.properties.listTitle = selectedList?.title ?? '';
      this.properties.isDocumentLibrary = selectedList?.isDocumentLibrary ?? false;
      this.properties.viewId = '';
      this.properties.viewTitle = '';
      this._views = [];
      this._fields = [];
      if (newValue) {
        await Promise.all([
          this._loadViews(String(newValue)),
          this._loadFields(String(newValue)),
        ]);
      }
      this.context.propertyPane.refresh();
      this.render();
    }

    if (propertyPath === 'viewId' && newValue !== oldValue) {
      const selectedView = this._views.find(v => v.id === String(newValue));
      this.properties.viewTitle = selectedView?.title ?? '';
      this.render();
    }

    if (propertyPath === 'defaultSortPreset' && newValue !== oldValue) {
      const preset = String(newValue);
      switch (preset) {
        case 'alphaAsc':
          this.properties.defaultSortField = this.properties.isDocumentLibrary ? 'FileLeafRef' : 'Title';
          this.properties.defaultSortDirection = 'asc';
          break;
        case 'alphaDesc':
          this.properties.defaultSortField = this.properties.isDocumentLibrary ? 'FileLeafRef' : 'Title';
          this.properties.defaultSortDirection = 'desc';
          break;
        case 'createdAsc':
          this.properties.defaultSortField = 'Created';
          this.properties.defaultSortDirection = 'asc';
          break;
        case 'createdDesc':
          this.properties.defaultSortField = 'Created';
          this.properties.defaultSortDirection = 'desc';
          break;
        case 'modifiedAsc':
          this.properties.defaultSortField = 'Modified';
          this.properties.defaultSortDirection = 'asc';
          break;
        case 'modifiedDesc':
          this.properties.defaultSortField = 'Modified';
          this.properties.defaultSortDirection = 'desc';
          break;
        case 'idAsc':
          this.properties.defaultSortField = 'Id';
          this.properties.defaultSortDirection = 'asc';
          break;
        case 'idDesc':
          this.properties.defaultSortField = 'Id';
          this.properties.defaultSortDirection = 'desc';
          break;
      }
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  // ── Data Loading ──────────────────────────────────────────────────────────

  private async _loadLists(): Promise<void> {
    if (this._listsLoading) return;
    this._listsLoading = true;
    try {
      this._lists = await this._dataService.getLists(this.properties.siteUrl || undefined);
    } catch (e) {
      console.error('[ContentLibraryWebPart] _loadLists error:', e);
      this._lists = [];
    } finally {
      this._listsLoading = false;
    }
  }

  private async _loadViews(listId: string): Promise<void> {
    if (this._viewsLoading || !listId) return;
    this._viewsLoading = true;
    try {
      this._views = await this._dataService.getViews(listId, this.properties.siteUrl || undefined);
    } catch (e) {
      console.error('[ContentLibraryWebPart] _loadViews error:', e);
      this._views = [];
    } finally {
      this._viewsLoading = false;
    }
  }

  private async _loadFields(listId: string): Promise<void> {
    if (this._fieldsLoading || !listId) return;
    this._fieldsLoading = true;
    try {
      this._fields = await this._dataService.getFields(listId, this.properties.siteUrl || undefined);
    } catch (e) {
      console.error('[ContentLibraryWebPart] _loadFields error:', e);
      this._fields = [];
    } finally {
      this._fieldsLoading = false;
    }
  }

  // ── Property Pane Configuration ───────────────────────────────────────────

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const listOptions: IPropertyPaneDropdownOption[] = this._lists.map(l => ({
      key: l.id,
      text: `${l.title}${l.isDocumentLibrary ? ' (Library)' : ''}`,
    }));

    const viewOptions: IPropertyPaneDropdownOption[] = this._views.map(v => ({
      key: v.id,
      text: v.title,
    }));

    const filterableFieldOptions: IPropertyPaneDropdownOption[] = [
      { key: '', text: '— Select a field —' },
      ...this._fields
        .filter(f => ['Text', 'Choice', 'MultiChoice', 'Boolean', 'Lookup', 'TaxonomyFieldType'].indexOf(f.fieldType) !== -1)
        .map(f => ({
          key: f.internalName,
          text: `${f.displayName} (${f.fieldType})`,
        })),
    ];

    // ── Colour picker custom field ─────────────────────────────────────────
    const colorPickerEnabled = !!(
      this.properties.enableFilters &&
      this.properties.filterFieldInternalName &&
      this.properties.enableCategoryColors
    );

    // Build the resolved colour map (auto-fill missing keys)
    let storedColorMap: Record<string, string> = {};
    try {
      storedColorMap = JSON.parse(this.properties.categoryColorsJson || '{}') as Record<string, string>;
    } catch { storedColorMap = {}; }

    let storedTextOverrides: Record<string, 'light' | 'dark'> = {};
    try {
      storedTextOverrides = JSON.parse(this.properties.categoryTextOverridesJson || '{}') as Record<string, 'light' | 'dark'>;
    } catch { storedTextOverrides = {}; }

    const autoMap = autoAssignCategoryColors(this._categoryKeys);
    const resolvedColorMap: Record<string, string> = {};
    this._categoryKeys.forEach(k => {
      resolvedColorMap[k] = storedColorMap[k] || autoMap[k];
    });

    // PropertyPaneCustomField is excluded from the public API in this SPFx version.
    // Construct the field descriptor manually using the IPropertyPaneField shape.
    const colorPickerProps: IPropertyPaneCustomFieldProps = {
      key: 'categoryColorPicker',
      onRender: (elem: HTMLElement) => {
        ReactDom.render(
          React.createElement(CategoryColorPicker, {
            colorMap: resolvedColorMap,
            textOverrides: storedTextOverrides,
            categoryKeys: this._categoryKeys,
            onChange: this._handleCategoryColorsChange.bind(this),
            disabled: !colorPickerEnabled,
          }),
          elem
        );
      },
      onDispose: (elem: HTMLElement) => {
        ReactDom.unmountComponentAtNode(elem);
      },
    };
    const colorPickerField: IPropertyPaneField<IPropertyPaneCustomFieldProps> = {
      // PropertyPaneFieldType.Custom = 1 — not exported in this SPFx version's public API
      type: 1 as IPropertyPaneField<IPropertyPaneCustomFieldProps>['type'],
      targetProperty: 'categoryColorPicker',
      properties: colorPickerProps,
      shouldFocus: false,
    };

    return {
      pages: [
        // ── Page 1: Data Source + Display Style ──────────────────────────
        {
          header: { description: 'Connect to a list or library and choose how to display it' },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: '📋 Data Source',
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('webPartTitle', {
                  label: 'Web part title',
                  placeholder: 'Content Library',
                }),
                PropertyPaneToggle('showTitle', {
                  label: 'Show title',
                  onText: 'Visible',
                  offText: 'Hidden',
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  placeholder: 'Leave blank for current site',
                  description: 'Enter the full URL of another site in your tenant, or leave blank to use the current site.',
                }),
                PropertyPaneDropdown('listId', {
                  label: 'List or document library',
                  options: listOptions.length > 0 ? listOptions : [{ key: '', text: 'Loading...' }],
                  selectedKey: this.properties.listId,
                }),
                PropertyPaneDropdown('viewId', {
                  label: 'View to display',
                  options: viewOptions.length > 0 ? viewOptions : [{ key: '', text: this.properties.listId ? 'Loading...' : 'Select a list first' }],
                  selectedKey: this.properties.viewId,
                  disabled: !this.properties.listId,
                }),
                PropertyPaneSlider('itemLimit', {
                  label: 'Maximum items to load',
                  min: 10,
                  max: 500,
                  step: 10,
                  value: this.properties.itemLimit || 50,
                  showValue: true,
                }),
              ],
            },
            {
              groupName: '🎨 Display Style',
              isCollapsed: false,
              groupFields: [
                PropertyPaneDropdown('itemDisplayStyle', {
                  label: 'Item display style',
                  options: [
                    { key: 'card-grid', text: 'Card grid — modern cards with metadata' },
                    { key: 'preview', text: 'Preview — thumbnail cards with image support' },
                    { key: 'table', text: 'Table / list view — document library style' },
                    { key: 'tile-grid', text: 'Tile grid — icon-focused tiles' },
                    { key: 'icon-grid', text: 'Icon grid — compact quick access' },
                    { key: 'dashboard', text: 'Dashboard — recent + favourites panels' },
                  ],
                  selectedKey: this.properties.itemDisplayStyle || 'card-grid',
                }),
                PropertyPaneDropdown('density', {
                  label: 'Spacing density',
                  options: [
                    { key: 'compact', text: 'Compact' },
                    { key: 'normal', text: 'Normal' },
                    { key: 'comfortable', text: 'Comfortable' },
                  ],
                  selectedKey: this.properties.density || 'normal',
                }),
                PropertyPaneSlider('gridColumns', {
                  label: 'Columns (card/tile modes)',
                  min: 1,
                  max: 6,
                  step: 1,
                  value: this.properties.gridColumns || 3,
                  showValue: true,
                }),
                PropertyPaneSlider('cardCornerRadius', {
                  label: 'Card corner radius (px)',
                  min: 0,
                  max: 24,
                  step: 2,
                  value: this.properties.cardCornerRadius || 8,
                  showValue: true,
                }),
                PropertyPaneDropdown('shadowIntensity', {
                  label: 'Shadow intensity',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'subtle', text: 'Subtle' },
                    { key: 'medium', text: 'Medium' },
                    { key: 'strong', text: 'Strong' },
                  ],
                  selectedKey: this.properties.shadowIntensity || 'subtle',
                }),
                PropertyPaneSlider('itemFontSize', {
                  label: 'Item text size (px)',
                  min: 10,
                  max: 20,
                  step: 1,
                  value: this.properties.itemFontSize || 13,
                  showValue: true,
                }),
                PropertyPaneSlider('itemIconSize', {
                  label: 'Item icon size (px)',
                  min: 14,
                  max: 48,
                  step: 2,
                  value: this.properties.itemIconSize || 24,
                  showValue: true,
                }),
              ],
            },
          ],
        },

        // ── Page 2: Search, Filter, Fields & Advanced ────────────────────
        {
          header: { description: 'Search, filters, visible fields, and advanced options' },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: '🔍 Search Settings',
              isCollapsed: false,
              groupFields: [
                PropertyPaneToggle('enableSearch', {
                  label: 'Enable search bar',
                  onText: 'On',
                  offText: 'Off',
                }),
                PropertyPaneTextField('searchPlaceholder', {
                  label: 'Search placeholder text',
                  placeholder: 'Search...',
                  disabled: !this.properties.enableSearch,
                }),
                PropertyPaneDropdown('searchBarStyle', {
                  label: 'Search bar style',
                  options: [
                    { key: 'minimal', text: 'Minimal — clean top bar' },
                    { key: 'elevated', text: 'Elevated — card with shadow' },
                    { key: 'toolbar', text: 'Toolbar — integrated with sort controls' },
                  ],
                  selectedKey: this.properties.searchBarStyle || 'minimal',
                  disabled: !this.properties.enableSearch,
                }),
                PropertyPaneDropdown('searchBarPosition', {
                  label: 'Search bar position',
                  options: [
                    { key: 'top-full', text: 'Top — full width' },
                    { key: 'top-right', text: 'Top — right aligned' },
                    { key: 'toolbar', text: 'Toolbar — with controls' },
                  ],
                  selectedKey: this.properties.searchBarPosition || 'top-full',
                  disabled: !this.properties.enableSearch,
                }),
                PropertyPaneSlider('searchDebounceMs', {
                  label: 'Search debounce (ms)',
                  min: 100,
                  max: 1000,
                  step: 50,
                  value: this.properties.searchDebounceMs || 300,
                  showValue: true,
                  disabled: !this.properties.enableSearch,
                }),
              ],
            },
            {
              groupName: '🏷️ Filter / Category Settings',
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('filterInfo', {
                  text: 'ℹ️ To use category filtering, your list should include a Choice or Single line of text column (e.g. Category, Status, Department, Location, Document Type). If no suitable field exists, leave filters off.',
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('enableFilters', {
                  label: 'Enable category filters',
                  onText: 'On',
                  offText: 'Off',
                }),
                PropertyPaneDropdown('filterFieldInternalName', {
                  label: 'Filter field',
                  options: filterableFieldOptions,
                  selectedKey: this.properties.filterFieldInternalName || '',
                  disabled: !this.properties.enableFilters || !this.properties.listId,
                }),
                PropertyPaneDropdown('filterStyle', {
                  label: 'Filter display style',
                  options: [
                    { key: 'pills', text: 'Horizontal pills / tabs' },
                    { key: 'vertical-rail', text: 'Vertical side navigation' },
                    { key: 'cards', text: 'Category cards' },
                    { key: 'compact-buttons', text: 'Compact button row' },
                  ],
                  selectedKey: this.properties.filterStyle || 'pills',
                  disabled: !this.properties.enableFilters,
                }),
                PropertyPaneDropdown('filterPosition', {
                  label: 'Filter position',
                  options: [
                    { key: 'top', text: 'Top (above content)' },
                    { key: 'left', text: 'Left side panel' },
                    { key: 'right', text: 'Right side panel' },
                  ],
                  selectedKey: this.properties.filterPosition || 'top',
                  disabled: !this.properties.enableFilters,
                }),
                PropertyPaneToggle('showAllOption', {
                  label: 'Show "All" option',
                  onText: 'Yes',
                  offText: 'No',
                  disabled: !this.properties.enableFilters,
                }),
                PropertyPaneTextField('allOptionLabel', {
                  label: '"All" option label',
                  placeholder: 'All',
                  disabled: !this.properties.enableFilters || !this.properties.showAllOption,
                }),
                PropertyPaneToggle('showCategoryCounts', {
                  label: 'Show item counts',
                  onText: 'Yes',
                  offText: 'No',
                  disabled: !this.properties.enableFilters,
                }),
                PropertyPaneDropdown('categorySortOrder', {
                  label: 'Sort categories',
                  options: [
                    { key: 'alpha', text: 'Alphabetically' },
                    { key: 'count', text: 'By item count (highest first)' },
                  ],
                  selectedKey: this.properties.categorySortOrder || 'alpha',
                  disabled: !this.properties.enableFilters,
                }),
                PropertyPaneSlider('maxVisibleCategories', {
                  label: 'Max visible categories',
                  min: 3,
                  max: 30,
                  step: 1,
                  value: this.properties.maxVisibleCategories || 10,
                  showValue: true,
                  disabled: !this.properties.enableFilters,
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel('colorCodingInfo', {
                  text: '🎨 Category Colour Coding',
                }),
                PropertyPaneLabel('colorCodingDesc', {
                  text: 'Tints cards and tiles with each category\'s colour. Text and icon colours adjust automatically for contrast.',
                }),
                PropertyPaneToggle('enableCategoryColors', {
                  label: 'Enable category colour coding',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !this.properties.enableFilters || !this.properties.filterFieldInternalName,
                }),
                ...(this.properties.enableCategoryColors && this.properties.enableFilters && this.properties.filterFieldInternalName
                  ? [
                      PropertyPaneLabel('colorPickerLabel', {
                        text: 'Click a swatch to open the colour picker, or choose from the quick palette dots. Changes apply instantly.',
                      }),
                      colorPickerField,
                    ]
                  : [
                      PropertyPaneLabel('colorCodingNote', {
                        text: this.properties.enableFilters && this.properties.filterFieldInternalName
                          ? '⬆️ Enable colour coding above to customise category colours.'
                          : '⬆️ Enable filters and select a filter field above to use colour coding.',
                      }),
                    ]
                ),
              ],
            },
            {
              groupName: '📄 Visible Fields',
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('showFileTypeIcon', {
                  label: 'Show file type icon',
                  onText: 'Yes',
                  offText: 'No',
                }),
                PropertyPaneToggle('showDescription', {
                  label: 'Show description (card mode)',
                  onText: 'Yes',
                  offText: 'No',
                }),
                PropertyPaneToggle('showColumnHeaders', {
                  label: 'Show column headers (table mode)',
                  onText: 'Yes',
                  offText: 'No',
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel('cardMetaLabel', {
                  text: '📋 Card / Tile detail lines',
                }),
                PropertyPaneLabel('cardMetaDesc', {
                  text: 'Choose which fields appear as the two detail lines on cards and tiles. Select "None" to hide a line.',
                }),
                PropertyPaneLabel('cardMetaNote', {
                  text: `ℹ️ A list and view must be selected on page 1 before custom columns appear in these dropdowns. Built-in options (Modified date, Created date, Modified by, Created by) are always available.`,
                }),
                PropertyPaneDropdown('cardMeta1Field', {
                  label: 'Detail line 1',
                  options: [
                    { key: '', text: '— None —' },
                    { key: 'Modified', text: 'Modified date' },
                    { key: 'Created', text: 'Created date' },
                    { key: 'Editor', text: 'Modified by' },
                    { key: 'Author', text: 'Created by' },
                    ...this._fields
                      .filter(f => ['Text', 'Note', 'DateTime', 'User', 'Choice', 'Number', 'Currency', 'Boolean', 'URL', 'Lookup'].indexOf(f.fieldType) !== -1)
                      .map(f => ({ key: f.internalName, text: f.displayName })),
                  ],
                  selectedKey: this.properties.cardMeta1Field !== undefined ? this.properties.cardMeta1Field : 'Modified',
                  disabled: !this.properties.listId,
                }),
                PropertyPaneDropdown('cardMeta2Field', {
                  label: 'Detail line 2',
                  options: [
                    { key: '', text: '— None —' },
                    { key: 'Modified', text: 'Modified date' },
                    { key: 'Created', text: 'Created date' },
                    { key: 'Editor', text: 'Modified by' },
                    { key: 'Author', text: 'Created by' },
                    ...this._fields
                      .filter(f => ['Text', 'Note', 'DateTime', 'User', 'Choice', 'Number', 'Currency', 'Boolean', 'URL', 'Lookup'].indexOf(f.fieldType) !== -1)
                      .map(f => ({ key: f.internalName, text: f.displayName })),
                  ],
                  selectedKey: this.properties.cardMeta2Field !== undefined ? this.properties.cardMeta2Field : 'Editor',
                  disabled: !this.properties.listId,
                }),
                PropertyPaneDropdown('cardMeta1Icon', {
                  label: 'Detail line 1 icon',
                  options: this._metaIconOptions,
                  selectedKey: this.properties.cardMeta1Icon || 'Clock',
                }),
                PropertyPaneDropdown('cardMeta2Icon', {
                  label: 'Detail line 2 icon',
                  options: this._metaIconOptions,
                  selectedKey: this.properties.cardMeta2Icon || 'Contact',
                }),
                PropertyPaneToggle('showChoicePillsOnCards', {
                  label: 'Show Choice column values as colored badges on item cards',
                  onText: 'Yes',
                  offText: 'No',
                }),
              ],
            },
            {
              groupName: '⚙️ Advanced',
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('sortHeading', {
                  text: 'Sort options',
                }),
                PropertyPaneDropdown('defaultSortPreset', {
                  label: 'Default sort order',
                  options: [
                    { key: 'alphaAsc', text: 'Alphabetical (A → Z)' },
                    { key: 'alphaDesc', text: 'Alphabetical (Z → A)' },
                    { key: 'createdAsc', text: 'Created date (earliest → latest)' },
                    { key: 'createdDesc', text: 'Created date (latest → earliest)' },
                    { key: 'modifiedAsc', text: 'Modified date (oldest → newest)' },
                    { key: 'modifiedDesc', text: 'Modified date (newest → oldest)' },
                    { key: 'idAsc', text: 'Item ID (lowest → highest)' },
                    { key: 'idDesc', text: 'Item ID (highest → lowest)' },
                  ],
                  selectedKey: this.properties.defaultSortPreset || 'modifiedDesc',
                }),
                PropertyPaneToggle('allowUserSort', {
                  label: 'Allow users to change sorting',
                  onText: 'Yes',
                  offText: 'No',
                }),
                PropertyPaneToggle('enableSortControl', {
                  label: 'Show sort controls in toolbar',
                  onText: 'Yes',
                  offText: 'No',
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneDropdown('linkTarget', {
                  label: 'Open documents in',
                  options: [
                    { key: '_blank', text: 'New tab' },
                    { key: '_self', text: 'Same tab' },
                  ],
                  selectedKey: this.properties.linkTarget || '_blank',
                }),
                PropertyPaneTextField('emptyStateMessage', {
                  label: 'Empty state message',
                  placeholder: 'No items found.',
                }),
                PropertyPaneTextField('customCssClass', {
                  label: 'Custom CSS class',
                  placeholder: 'my-custom-class',
                  description: 'Add a custom CSS class to the web part root for targeted styling.',
                }),
                PropertyPaneLabel('advancedNote', {
                  text: 'ℹ️ For advanced customisation, target the custom CSS class in a site-level style sheet.',
                }),
              ],
            },
          ],
        },

      ],
    };
  }
}
