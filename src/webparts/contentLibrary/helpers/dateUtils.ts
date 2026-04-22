import { IListItem } from '../models/IListItem';

/** Number of days after creation that an item is considered "new". */
const NEW_ITEM_THRESHOLD_DAYS = 7;

/**
 * Returns true if the item was created (or, as fallback, modified) within the
 * last NEW_ITEM_THRESHOLD_DAYS days.
 *
 * Priority: item.created → item.modified → false
 */
export function isItemNew(item: IListItem): boolean {
  const dateStr = item.created ?? item.modified;
  if (!dateStr) return false;

  try {
    const created = new Date(dateStr);
    if (isNaN(created.getTime())) return false;

    const diffMs = Date.now() - created.getTime();
    const diffDays = diffMs / (1000 * 60 * 60 * 24);
    return diffDays >= 0 && diffDays <= NEW_ITEM_THRESHOLD_DAYS;
  } catch {
    return false;
  }
}
