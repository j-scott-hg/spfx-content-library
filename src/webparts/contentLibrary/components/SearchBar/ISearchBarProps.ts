import * as React from 'react';
import { SearchBarStyle, SearchBarPosition } from '../../models/IWebPartConfig';

export interface ISearchBarProps {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  style: SearchBarStyle;
  position: SearchBarPosition;
  debounceMs?: number;
  /** Inline style applied to the outermost wrapper (used for top-right sizing) */
  wrapperStyle?: React.CSSProperties;
  /** Optional label for the sort dropdown shown in toolbar mode */
  sortLabel?: string;
  onSortChange?: (field: string, asc: boolean) => void;
  sortFields?: Array<{ key: string; text: string }>;
  currentSortField?: string;
  currentSortAsc?: boolean;
  showSortControl?: boolean;
}
