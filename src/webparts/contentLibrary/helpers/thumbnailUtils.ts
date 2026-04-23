import { IListItem } from '../models/IListItem';
import { IItemIconOverride } from '../models/IWebPartConfig';

/**
 * Best thumbnail URL for an item (matches Preview grid logic):
 * 1. customThumbnailUrl from web part icon overrides
 * 2. SharePoint getpreview.ashx when fileRef exists (document libraries)
 */
export function resolveItemThumbnailUrl(item: IListItem, override: IItemIconOverride | undefined): string | undefined {
  if (override?.customThumbnailUrl) {
    return override.customThumbnailUrl;
  }
  if (item.fileRef) {
    return `/_layouts/15/getpreview.ashx?resolution=3&path=${encodeURIComponent(item.fileRef)}`;
  }
  return undefined;
}
