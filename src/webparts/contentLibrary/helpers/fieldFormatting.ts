import { format, parseISO, isValid } from 'date-fns';

/**
 * Formats a SharePoint date string to a human-readable relative or absolute string.
 */
export function formatDate(dateStr?: string, useRelative = true): string {
  if (!dateStr) return '';
  try {
    const date = parseISO(dateStr);
    if (!isValid(date)) return dateStr;

    if (useRelative) {
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffSec = Math.floor(diffMs / 1000);
      const diffMin = Math.floor(diffSec / 60);
      const diffHr = Math.floor(diffMin / 60);
      const diffDay = Math.floor(diffHr / 24);

      if (diffSec < 60) return 'Just now';
      if (diffMin < 60) return `${diffMin} minute${diffMin === 1 ? '' : 's'} ago`;
      if (diffHr < 24) return `${diffHr} hour${diffHr === 1 ? '' : 's'} ago`;
      if (diffDay === 1) return 'Yesterday';
      if (diffDay < 7) return `${diffDay} days ago`;
      if (diffDay < 30) return `${Math.floor(diffDay / 7)} week${Math.floor(diffDay / 7) === 1 ? '' : 's'} ago`;
    }

    return format(date, 'dd/MM/yyyy');
  } catch {
    return dateStr;
  }
}

/**
 * Formats a file size in bytes to a human-readable string.
 */
export function formatFileSize(bytes?: number): string {
  if (bytes === undefined || bytes === null) return '';
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
}

/**
 * Formats a field value for display based on its type.
 */
export function formatFieldValue(value: unknown, fieldType: string): string {
  if (value === null || value === undefined) return '';

  switch (fieldType) {
    case 'DateTime':
      return formatDate(String(value));
    case 'Boolean':
      return value ? 'Yes' : 'No';
    case 'User':
    case 'UserMulti': {
      if (typeof value === 'object' && value !== null) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const userObj = value as any;
        return String(userObj.Title ?? userObj.Name ?? '');
      }
      return String(value);
    }
    case 'Lookup':
    case 'LookupMulti': {
      if (typeof value === 'object' && value !== null) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const lookupObj = value as any;
        return String(lookupObj.LookupValue ?? '');
      }
      return String(value);
    }
    case 'Currency':
      return typeof value === 'number' ? `$${value.toFixed(2)}` : String(value);
    case 'Number':
      return typeof value === 'number' ? value.toLocaleString() : String(value);
    case 'URL': {
      if (typeof value === 'object' && value !== null) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const urlObj = value as any;
        return String(urlObj.Description ?? urlObj.Url ?? '');
      }
      return String(value);
    }
    default:
      return String(value);
  }
}

/**
 * Truncates a string to a maximum length with ellipsis.
 */
export function truncate(str: string, maxLength: number): string {
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength - 3) + '...';
}
