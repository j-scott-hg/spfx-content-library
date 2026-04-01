import { FilterStyle } from '../../models/IWebPartConfig';
import { ICategoryOption } from '../../helpers/categoryExtraction';

export interface IFilterBarProps {
  categories: ICategoryOption[];
  selectedCategory: string;
  onCategoryChange: (key: string) => void;
  style: FilterStyle;
  showAllOption: boolean;
  allOptionLabel: string;
  showCounts: boolean;
  maxVisible?: number;
  /** Optional map of category key → hex colour for colour-coded filter tabs */
  categoryColors?: Record<string, string>;
  enableCategoryColors?: boolean;
  /** Optional per-category text colour override: 'light' = white, 'dark' = near-black */
  categoryTextOverrides?: Record<string, 'light' | 'dark'>;
}
