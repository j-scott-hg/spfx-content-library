import { IListItem } from '../models/IListItem';

/**
 * Filters items by a search query across specified fields.
 * Case-insensitive substring match.
 */
export function filterItemsBySearch(
  items: IListItem[],
  query: string,
  searchFields: string[]
): IListItem[] {
  if (!query || !query.trim()) return items;
  const q = query.trim().toLowerCase();

  return items.filter(item => {
    for (const field of searchFields) {
      const value = item[field];
      if (value !== null && value !== undefined) {
        const str = String(value).toLowerCase();
        if (str.indexOf(q) !== -1) return true;
      }
    }
    return false;
  });
}

/**
 * Sorts items by a given field in ascending or descending order.
 */
export function sortItems(
  items: IListItem[],
  sortField: string,
  sortAsc: boolean
): IListItem[] {
  if (!sortField) return items;

  const toDateValue = (value: unknown): number | null => {
    if (typeof value !== 'string') return null;
    const timestamp = Date.parse(value);
    return Number.isNaN(timestamp) ? null : timestamp;
  };

  return [...items].sort((a, b) => {
    const aVal = a[sortField];
    const bVal = b[sortField];

    if (aVal === null || aVal === undefined) return sortAsc ? 1 : -1;
    if (bVal === null || bVal === undefined) return sortAsc ? -1 : 1;

    if (typeof aVal === 'number' && typeof bVal === 'number') {
      return sortAsc ? aVal - bVal : bVal - aVal;
    }

    const aDate = toDateValue(aVal);
    const bDate = toDateValue(bVal);
    if (aDate !== null && bDate !== null) {
      return sortAsc ? aDate - bDate : bDate - aDate;
    }

    const aStr = String(aVal).toLowerCase();
    const bStr = String(bVal).toLowerCase();
    const cmp = aStr.localeCompare(bStr);
    return sortAsc ? cmp : -cmp;
  });
}

/**
 * Creates a debounced version of a function.
 */
export function debounce<T extends (...args: unknown[]) => void>(
  fn: T,
  delayMs: number
): (...args: Parameters<T>) => void {
  let timer: ReturnType<typeof setTimeout>;
  return (...args: Parameters<T>) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delayMs);
  };
}
