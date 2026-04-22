import * as React from 'react';
import { useState, useCallback, useEffect } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { FilePicker } from '@pnp/spfx-controls-react/lib/FilePicker';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/controls/filePicker/FilePicker.types';
import { FilePickerTab } from '@pnp/spfx-controls-react/lib/controls/filePicker/FilePickerTab';
import { IListItem } from '../../models/IListItem';
import { IColumnDef } from '../../services/ViewMapper';
import { IItemIconOverride, LinkTarget } from '../../models/IWebPartConfig';
import { formatDate, formatFieldValue } from '../../helpers/fieldFormatting';
import { getFileIconInfo, getFileIconColorHex } from '../../helpers/fileIconMapping';
import { getContrastTextColor, tintColor } from '../../helpers/colorUtils';
import { isItemNew } from '../../helpers/dateUtils';
import styles from '../../styles/ContentLibrary.module.scss';

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IDocumentPreviewGridProps {
  items: IListItem[];
  columns: IColumnDef[];
  gridColumns: number;
  isDocumentLibrary: boolean;
  linkTarget: LinkTarget;
  cardCornerRadius: number;
  isEditMode: boolean;
  /** SPFx context — required by ImagePicker for site/tenant access */
  context: WebPartContext;
  itemIconOverrides: Record<string, IItemIconOverride>;
  onEditItemIcon: (itemId: string, defaultIconName: string, defaultIconColor: string, title: string) => void;
  /** Called when user saves a thumbnail override from inside this component */
  onSaveThumbnail: (itemId: string, thumbnailUrl: string | undefined) => void;
  onItemClick: (item: IListItem) => void;
  /** Internal field name for first meta line (empty = hide) */
  cardMeta1Field: string;
  /** Internal field name for second meta line (empty = hide) */
  cardMeta2Field: string;
  allColumns: IColumnDef[];
  enableCategoryColors: boolean;
  categoryColors: Record<string, string>;
  categoryTextOverrides: Record<string, 'light' | 'dark'>;
  filterFieldInternalName: string;
  itemFontSize: number;
  itemIconSize: number;
}

// ─── Upload helper ────────────────────────────────────────────────────────────

/**
 * Uploads a File to the current site's Site Assets library under a
 * "ContentLibraryThumbnails" subfolder and returns the absolute URL.
 *
 * Uses PnPjs v3 with the SPFx context so authentication is handled
 * automatically by the existing credential chain.
 */
async function uploadToSiteAssets(file: File, context: WebPartContext): Promise<string> {
  const sp = spfi().using(SPFx(context));
  const folderUrl = `${context.pageContext.web.serverRelativeUrl}/SiteAssets/ContentLibraryThumbnails`;

  // Ensure the folder exists (no-op if it already does)
  try {
    await sp.web.folders.addUsingPath(folderUrl, true);
  } catch {
    // folder may already exist — continue
  }

  // Upload the file, overwriting if a file with the same name exists
  const uploadResult = await sp.web
    .getFolderByServerRelativePath(folderUrl)
    .files.addUsingPath(file.name, file, { Overwrite: true });

  // Build the absolute URL from the server-relative path returned by PnPjs
  const serverRelativeUrl: string = (uploadResult as unknown as { ServerRelativeUrl: string }).ServerRelativeUrl
    ?? uploadResult.data?.ServerRelativeUrl
    ?? `${folderUrl}/${file.name}`;

  return `${window.location.origin}${serverRelativeUrl}`;
}

// ─── Thumbnail resolution ─────────────────────────────────────────────────────

/**
 * Returns the best available thumbnail URL for an item in priority order:
 *  1. customThumbnailUrl stored in the icon override
 *  2. SharePoint file preview URL (only available for doc libs with a fileRef)
 *  3. undefined → component falls back to file-type icon
 *
 * SharePoint preview URL pattern (works for most Office files and images in SPO):
 *   /_layouts/15/getpreview.ashx?path=<encoded server-relative path>
 */
function resolveThumbnailUrl(item: IListItem, override: IItemIconOverride | undefined): string | undefined {
  if (override?.customThumbnailUrl) return override.customThumbnailUrl;

  if (item.fileRef) {
    // SharePoint's built-in preview thumbnail endpoint
    return `/_layouts/15/getpreview.ashx?resolution=3&path=${encodeURIComponent(item.fileRef)}`;
  }

  return undefined;
}

// ─── Thumbnail editor panel ───────────────────────────────────────────────────

interface IThumbnailEditorProps {
  isOpen: boolean;
  itemTitle: string;
  itemId: string;
  currentUrl: string | undefined;
  context: WebPartContext;
  onSave: (itemId: string, url: string | undefined) => void;
  onDismiss: () => void;
}

const ThumbnailEditor: React.FC<IThumbnailEditorProps> = ({
  isOpen, itemTitle, itemId, currentUrl, context, onSave, onDismiss,
}) => {
  // URL staged within this panel session — not committed until Apply is clicked
  const [pendingUrl, setPendingUrl] = useState<string>(currentUrl ?? '');
  // Whether the FilePicker panel is open
  const [filePickerOpen, setFilePickerOpen] = useState(false);
  // Whether the staged image preview failed to load
  const [previewError, setPreviewError] = useState(false);
  // Uploading state — true while a local file is being uploaded to Site Assets
  const [uploading, setUploading] = useState(false);
  // Upload error message
  const [uploadError, setUploadError] = useState<string | undefined>();

  // Reset staged state each time the panel opens for a (possibly different) item
  useEffect(() => {
    if (isOpen) {
      setPendingUrl(currentUrl ?? '');
      setPreviewError(false);
      setFilePickerOpen(false);
      setUploading(false);
      setUploadError(undefined);
    }
  }, [isOpen, itemId, currentUrl]);

  /**
   * FilePicker onSave handler.
   *
   * Two cases:
   *  A) fileAbsoluteUrl is present  → image already lives in SharePoint (Site
   *     Files tab selection). Use the URL directly.
   *  B) fileAbsoluteUrl is undefined → local file upload. The FilePicker has
   *     already staged the file; we call downloadFileContent() to get the File
   *     object, upload it to Site Assets via PnPjs, then store the resulting
   *     absolute URL.
   */
  const handleFilePickerSave = useCallback(async (results: IFilePickerResult[]) => {
    setFilePickerOpen(false);
    const first = results && results[0];
    if (!first) return;

    // Case A — already has an absolute URL (selected from Site Files)
    if (first.fileAbsoluteUrl) {
      setPendingUrl(first.fileAbsoluteUrl);
      setPreviewError(false);
      return;
    }

    // Case B — local file upload; fileAbsoluteUrl is undefined per the type docs
    if (typeof first.downloadFileContent === 'function') {
      setUploading(true);
      setUploadError(undefined);
      try {
        const file: File = await first.downloadFileContent();
        const absoluteUrl = await uploadToSiteAssets(file, context);
        setPendingUrl(absoluteUrl);
        setPreviewError(false);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        setUploadError(`Upload failed: ${msg}`);
      } finally {
        setUploading(false);
      }
    }
  }, [context]);

  const handleSave = useCallback(() => {
    onSave(itemId, pendingUrl.trim() || undefined);
    onDismiss();
  }, [itemId, pendingUrl, onSave, onDismiss]);

  const handleRemove = useCallback(() => {
    onSave(itemId, undefined);
    onDismiss();
  }, [itemId, onSave, onDismiss]);

  const footer = (): React.ReactElement => (
    <div style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '8px 0' }}>
      <PrimaryButton text="Apply" onClick={handleSave} disabled={uploading} />
      <DefaultButton text="Cancel" onClick={onDismiss} disabled={uploading} />
      {currentUrl && (
        <DefaultButton
          text="Remove"
          iconProps={{ iconName: 'Delete' }}
          onClick={handleRemove}
          disabled={uploading}
          style={{ marginLeft: 'auto' }}
        />
      )}
    </div>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.smallFixedFar}
      headerText={`Thumbnail — ${itemTitle}`}
      isFooterAtBottom
      onRenderFooterContent={footer}
      closeButtonAriaLabel="Close thumbnail editor"
      styles={{ main: { maxWidth: 360 } }}
    >
      <div style={{ padding: '4px 0 16px' }}>
        {/* Help note */}
        <p className={styles.thumbnailHint}>
          Recommended size: <strong>640 &times; 360 px</strong> (16:9). Use a lightweight
          image for best performance. Leave unset to use the SharePoint-generated preview.
        </p>

        {/* Upload progress */}
        {uploading && (
          <div className={styles.thumbnailUploading}>
            <Spinner size={SpinnerSize.small} label="Uploading to Site Assets…" />
          </div>
        )}

        {/* Upload error */}
        {uploadError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setUploadError(undefined)}
            styles={{ root: { marginBottom: 8 } }}
          >
            {uploadError}
          </MessageBar>
        )}

        {/* Staged thumbnail preview */}
        {!uploading && pendingUrl && !previewError && (
          <div className={styles.thumbnailPreviewWrapper}>
            <img
              src={pendingUrl}
              alt="Thumbnail preview"
              className={styles.thumbnailPreviewImg}
              onError={() => setPreviewError(true)}
            />
          </div>
        )}
        {!uploading && pendingUrl && previewError && (
          <div className={styles.thumbnailPreviewError}>
            Could not load image
          </div>
        )}

        {/* Select / change image button */}
        <DefaultButton
          text={pendingUrl ? 'Change image' : 'Select image'}
          iconProps={{ iconName: 'Photo2' }}
          onClick={() => setFilePickerOpen(true)}
          disabled={uploading}
          style={{ marginTop: 10 }}
        />

        {/* FilePicker — button hidden; isPanelOpen driven by local state.
            Tabs restricted to Upload + Site files to avoid Graph 404 errors. */}
        <FilePicker
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          context={context as unknown as any}
          accepts={['.gif', '.jpg', '.jpeg', '.png', '.svg', '.webp']}
          isPanelOpen={filePickerOpen}
          hidden
          hideStockImages
          hideWebSearchTab
          hideOrganisationalAssetTab
          hideRecentTab
          hideOneDriveTab
          hideLocalMultipleUploadTab
          hideLinkUploadTab
          defaultSelectedTab={FilePickerTab.Upload}
          onSave={handleFilePickerSave}
          onCancel={() => setFilePickerOpen(false)}
          buttonLabel="Select image"
        />
      </div>
    </Panel>
  );
};

// ─── Thumbnail image sub-component ───────────────────────────────────────────
// Defined before DocumentPreviewGrid to satisfy no-use-before-define.

interface IThumbnailImageProps {
  thumbnailUrl: string | undefined;
  hasCustomThumbnail: boolean;
  title: string;
  iconName: string;
  iconColor: string;
  iconSize: number;
}

const ThumbnailImage: React.FC<IThumbnailImageProps> = ({
  thumbnailUrl, title, iconName, iconColor, iconSize,
}) => {
  const [failed, setFailed] = useState(false);

  useEffect(() => { setFailed(false); }, [thumbnailUrl]);

  const showImage = thumbnailUrl && !failed;

  return (
    <div className={styles.previewThumbnail}>
      {showImage ? (
        <img
          src={thumbnailUrl}
          alt={title}
          onError={() => setFailed(true)}
        />
      ) : (
        <div className={styles.previewThumbnailFallback} aria-hidden="true">
          <Icon iconName={iconName} style={{ fontSize: iconSize, color: iconColor }} />
        </div>
      )}
    </div>
  );
};

// ─── Main component ───────────────────────────────────────────────────────────

interface IThumbnailEditorState {
  itemId: string;
  itemTitle: string;
  currentUrl: string | undefined;
}

const DocumentPreviewGrid: React.FC<IDocumentPreviewGridProps> = ({
  items,
  gridColumns,
  isDocumentLibrary,
  cardCornerRadius,
  isEditMode,
  context,
  itemIconOverrides,
  onEditItemIcon,
  onSaveThumbnail,
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
  const [thumbnailEditor, setThumbnailEditor] = useState<IThumbnailEditorState | undefined>();

  // Build column lookup once
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

  const metaIcon = (fieldName: string): string => {
    if (fieldName === 'Modified' || fieldName === 'Created') return 'Clock';
    if (fieldName === 'Editor' || fieldName === 'Author') return 'Contact';
    const col = colMap[fieldName];
    if (!col) return 'Info';
    if (col.fieldType === 'DateTime') return 'Clock';
    if (col.fieldType === 'User') return 'Contact';
    return 'Tag';
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const openThumbnailEditor = useCallback((item: IListItem, currentUrl: string | undefined, e: any) => {
    e.preventDefault();
    e.stopPropagation();
    setThumbnailEditor({ itemId: item.id, itemTitle: item.title ?? item.fileLeafRef ?? '', currentUrl });
  }, []);

  return (
    <>
      <div
        className={styles.previewGrid}
        style={{ '--preview-cols': gridColumns } as React.CSSProperties}
        role="list"
        aria-label="Preview cards"
      >
        {items.map(item => {
          const override = itemIconOverrides[item.id];
          const defaultIconInfo = getFileIconInfo(item.fileType, item.isFolder);
          const defaultIconColor = getFileIconColorHex(item.fileType, item.isFolder);
          const iconName = override?.iconName ?? defaultIconInfo.iconName;
          const iconColor = override?.iconColor ?? defaultIconColor;

          const displayTitle = isDocumentLibrary
            ? (item.fileLeafRef ?? item.title)
            : item.title;

          const thumbnailUrl = resolveThumbnailUrl(item, override);

          const catValue = filterFieldInternalName ? String(item[filterFieldInternalName] ?? '') : '';
          const catBgHex = enableCategoryColors && catValue && categoryColors[catValue] ? categoryColors[catValue] : undefined;
          const textOverride = catValue ? categoryTextOverrides?.[catValue] : undefined;
          const cardTextColor = catBgHex
            ? (textOverride === 'light' ? '#ffffff' : textOverride === 'dark' ? '#1b1b1b' : getContrastTextColor(tintColor(catBgHex, 0.18)))
            : undefined;

          const isNew = isItemNew(item);

          const card = (
            <div
              role="button"
              tabIndex={0}
              className={styles.previewCard}
              style={{ borderRadius: cardCornerRadius }}
              aria-label={displayTitle}
              onClick={() => onItemClick(item)}
              onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && onItemClick(item)}
            >
              {/* NEW badge */}
              {isNew && (
                <span className={styles.newBadge} aria-label="New item">NEW</span>
              )}

              {/* Category accent bar */}
              {catBgHex && (
                <div
                  className={styles.previewCategoryBar}
                  style={{ background: catBgHex }}
                  aria-hidden="true"
                />
              )}

              {/* Thumbnail */}
              <ThumbnailImage
                thumbnailUrl={thumbnailUrl}
                hasCustomThumbnail={!!override?.customThumbnailUrl}
                title={displayTitle ?? ''}
                iconName={iconName}
                iconColor={iconColor}
                iconSize={itemIconSize}
              />

              {/* Body */}
              <div className={styles.previewBody}>
                <div
                  className={styles.previewTitle}
                  title={displayTitle}
                  style={{ fontSize: itemFontSize, ...(cardTextColor ? { color: cardTextColor } : {}) }}
                >
                  {displayTitle}
                </div>

                {[cardMeta1Field, cardMeta2Field].map((fieldName, idx) => {
                  const val = resolveMetaValue(item, fieldName);
                  if (!val) return null;
                  return (
                    <div
                      key={idx}
                      className={styles.previewMeta}
                      style={cardTextColor ? { color: cardTextColor, opacity: 0.75 } : undefined}
                    >
                      <Icon iconName={metaIcon(fieldName)} className={styles.previewMetaIcon} aria-hidden="true" />
                      <span className={styles.previewMetaText}>{val}</span>
                    </div>
                  );
                })}
              </div>

              {/* Edit thumbnail button (preview-mode edit; always visible in edit mode, hover in view mode) */}
              {isEditMode && (
                <div className={styles.previewCustomizeBtn}>
                  <IconButton
                    iconProps={{ iconName: 'Photo2' }}
                    title="Set custom thumbnail"
                    ariaLabel={`Set custom thumbnail for ${displayTitle}`}
                    onClick={e => openThumbnailEditor(item, override?.customThumbnailUrl, e)}
                    styles={{
                      root: { background: 'rgba(255,255,255,0.9)', borderRadius: '50%', width: 28, height: 28 },
                      icon: { fontSize: 14 },
                    }}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Color' }}
                    title="Edit icon"
                    ariaLabel={`Edit icon for ${displayTitle}`}
                    onClick={e => {
                      e.preventDefault();
                      e.stopPropagation();
                      onEditItemIcon(item.id, defaultIconInfo.iconName, defaultIconColor, displayTitle ?? '');
                    }}
                    styles={{
                      root: { background: 'rgba(255,255,255,0.9)', borderRadius: '50%', width: 28, height: 28, marginLeft: 4 },
                      icon: { fontSize: 14 },
                    }}
                  />
                </div>
              )}
            </div>
          );

          return (
            <div key={item.id} role="listitem">
              {card}
            </div>
          );
        })}
      </div>

      {/* Thumbnail editor panel — rendered once outside the list */}
      <ThumbnailEditor
        isOpen={!!thumbnailEditor}
        itemId={thumbnailEditor?.itemId ?? ''}
        itemTitle={thumbnailEditor?.itemTitle ?? ''}
        currentUrl={thumbnailEditor?.currentUrl}
        context={context}
        onSave={onSaveThumbnail}
        onDismiss={() => setThumbnailEditor(undefined)}
      />
    </>
  );
};

export default DocumentPreviewGrid;
