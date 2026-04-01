import { IListItem } from '../models/IListItem';
import { CategorySortOrder } from '../models/IWebPartConfig';

export interface ICategoryOption {
  key: string;
  label: string;
  count: number;
}

/**
 * Extracts unique category values from a field on a list of items.
 * Returns sorted options with counts.
 */
export function extractCategories(
  items: IListItem[],
  fieldInternalName: string,
  sortOrder: CategorySortOrder = 'alpha',
  allLabel = 'All'
): ICategoryOption[] {
  if (!fieldInternalName || !items.length) return [];

  const countMap = new Map<string, number>();

  for (const item of items) {
    const rawValue = item[fieldInternalName];
    const values = normalizeFieldValue(rawValue);

    for (const val of values) {
      const trimmed = val.trim();
      if (trimmed) {
        countMap.set(trimmed, (countMap.get(trimmed) ?? 0) + 1);
      }
    }
  }

  const options: ICategoryOption[] = [];
  countMap.forEach((count, key) => {
    options.push({ key, label: key, count });
  });

  if (sortOrder === 'alpha') {
    options.sort((a, b) => a.label.localeCompare(b.label));
  } else {
    options.sort((a, b) => b.count - a.count);
  }

  return options;
}

/**
 * Normalizes a field value to an array of strings.
 * Handles single values, arrays (multi-choice), and lookup objects.
 */
function normalizeFieldValue(value: unknown): string[] {
  if (value === null || value === undefined) return [];

  if (Array.isArray(value)) {
    const result: string[] = [];
    value.forEach(v => {
      normalizeFieldValue(v).forEach(s => result.push(s));
    });
    return result;
  }

  if (typeof value === 'object') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const obj = value as any;
    // Lookup field
    if ('LookupValue' in obj) return [String(obj.LookupValue)];
    // User field
    if ('Title' in obj) return [String(obj.Title)];
    return [];
  }

  return [String(value)];
}

/**
 * Checks whether an item matches a given category value for a field.
 */
export function itemMatchesCategory(
  item: IListItem,
  fieldInternalName: string,
  categoryKey: string
): boolean {
  if (!fieldInternalName || !categoryKey) return true;
  const rawValue = item[fieldInternalName];
  const values = normalizeFieldValue(rawValue);
  return values.some(v => v.trim() === categoryKey);
}
