import * as React from 'react';
import { useMemo } from 'react';
import { IFilterBarProps } from './IFilterBarProps';
import { ICategoryOption } from '../../helpers/categoryExtraction';
import { getContrastTextColor } from '../../helpers/colorUtils';
import styles from '../../styles/ContentLibrary.module.scss';

const ALL_KEY = '__all__';

const FilterBar: React.FC<IFilterBarProps> = ({
  categories,
  selectedCategory,
  onCategoryChange,
  style,
  showAllOption,
  allOptionLabel,
  showCounts,
  maxVisible = 20,
  categoryColors,
  enableCategoryColors,
  categoryTextOverrides,
}) => {
  const visibleCategories = useMemo((): ICategoryOption[] => {
    return categories.slice(0, maxVisible);
  }, [categories, maxVisible]);

  const allOption: ICategoryOption = {
    key: ALL_KEY,
    label: allOptionLabel || 'All',
    count: categories.reduce((sum, c) => sum + c.count, 0),
  };

  const options = showAllOption ? [allOption, ...visibleCategories] : visibleCategories;
  const activeKey = selectedCategory || ALL_KEY;

  const handleClick = (key: string): void => {
    onCategoryChange(key === ALL_KEY ? '' : key);
  };

  // "All" is always light grey — never inherits the theme-blue active state
  const ALL_STYLE_INACTIVE: React.CSSProperties = {
    background: '#f3f2f1',
    borderColor: '#e1dfdd',
    color: '#323130',
  };
  const ALL_STYLE_ACTIVE: React.CSSProperties = {
    background: '#e1dfdd',
    borderColor: '#c8c6c4',
    color: '#201f1e',
    fontWeight: 600,
  };

  /** Resolves the text colour for a category, respecting manual overrides */
  const resolveFg = (key: string, bg: string): string => {
    const override = categoryTextOverrides?.[key];
    if (override === 'light') return '#ffffff';
    if (override === 'dark') return '#201f1e';
    return getContrastTextColor(bg);
  };

  /** Returns inline style overrides for a category button when colour coding is on */
  const getCategoryStyle = (key: string, isActive: boolean): React.CSSProperties => {
    if (key === ALL_KEY) return isActive ? ALL_STYLE_ACTIVE : ALL_STYLE_INACTIVE;
    if (!enableCategoryColors || !categoryColors) return {};
    const bg = categoryColors[key];
    if (!bg) return {};
    const fg = resolveFg(key, bg);
    if (isActive) {
      return { background: bg, borderColor: bg, color: fg };
    }
    // Inactive: colour dot accent + apply text override if set
    const style: React.CSSProperties = { borderLeftColor: bg, borderLeftWidth: 3 };
    if (categoryTextOverrides?.[key]) style.color = fg;
    return style;
  };

  /** Returns a small colour dot for the rail/compact styles */
  const ColorDot = ({ catKey }: { catKey: string }): React.ReactElement | null => {
    if (!enableCategoryColors || !categoryColors || catKey === ALL_KEY) return null;
    const bg = categoryColors[catKey];
    if (!bg) return null;
    return (
      <span
        style={{
          display: 'inline-block',
          width: 8,
          height: 8,
          borderRadius: '50%',
          background: bg,
          flexShrink: 0,
          marginRight: 4,
        }}
        aria-hidden="true"
      />
    );
  };

  if (style === 'pills') {
    return (
      <nav aria-label="Filter by category">
        <div className={styles.filterPills} role="list">
          {options.map(opt => {
            const isActive = activeKey === opt.key;
            const catStyle = getCategoryStyle(opt.key, isActive);
            return (
              <button
                key={opt.key}
                role="listitem"
                className={`${styles.filterPill} ${isActive ? styles.active : ''}`}
                style={catStyle}
                onClick={() => handleClick(opt.key)}
                aria-pressed={isActive}
                aria-label={`${opt.label}${showCounts ? `, ${opt.count} items` : ''}`}
              >
                <ColorDot catKey={opt.key} />
                {opt.label}
                {showCounts && (
                  <span className={styles.filterCount} aria-hidden="true">
                    {opt.count}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      </nav>
    );
  }

  if (style === 'compact-buttons') {
    return (
      <nav aria-label="Filter by category">
        <div className={styles.filterCompact} role="list">
          {options.map(opt => {
            const isActive = activeKey === opt.key;
            const catStyle = getCategoryStyle(opt.key, isActive);
            return (
              <button
                key={opt.key}
                role="listitem"
                className={`${styles.filterCompactBtn} ${isActive ? styles.active : ''}`}
                style={catStyle}
                onClick={() => handleClick(opt.key)}
                aria-pressed={isActive}
                aria-label={`${opt.label}${showCounts ? `, ${opt.count} items` : ''}`}
              >
                <ColorDot catKey={opt.key} />
                {opt.label}
                {showCounts && (
                  <span className={styles.filterCount} aria-hidden="true">
                    {opt.count}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      </nav>
    );
  }

  if (style === 'cards') {
    const defaultColors = ['#0078d4', '#107c41', '#c43e1c', '#7719aa', '#038387', '#881798'];
    return (
      <nav aria-label="Filter by category">
        <div className={styles.categoryCards} role="list">
          {options.map((opt, idx) => {
            const isActive = activeKey === opt.key;
            // "All" is always grey
            if (opt.key === ALL_KEY) {
              const allCardStyle = isActive ? ALL_STYLE_ACTIVE : ALL_STYLE_INACTIVE;
              return (
                <button
                  key={opt.key}
                  role="listitem"
                  className={`${styles.categoryCard} ${isActive ? styles.active : ''}`}
                  style={allCardStyle}
                  onClick={() => handleClick(opt.key)}
                  aria-pressed={isActive}
                  aria-label={`${opt.label}${showCounts ? `, ${opt.count} items` : ''}`}
                >
                  <div className={styles.categoryCardLabel}>{opt.label}</div>
                  {showCounts && <div className={styles.categoryCardCount}>{opt.count} items</div>}
                </button>
              );
            }
            // Use category colour map if enabled, otherwise fall back to default palette
            const bg = (enableCategoryColors && categoryColors && categoryColors[opt.key])
              ? categoryColors[opt.key]
              : defaultColors[(idx - 1) % defaultColors.length]; // -1 because All takes idx 0
            const fg = resolveFg(opt.key, bg);
            return (
              <button
                key={opt.key}
                role="listitem"
                className={`${styles.categoryCard} ${isActive ? styles.active : ''}`}
                style={{ background: bg, color: fg }}
                onClick={() => handleClick(opt.key)}
                aria-pressed={isActive}
                aria-label={`${opt.label}${showCounts ? `, ${opt.count} items` : ''}`}
              >
                <div className={styles.categoryCardLabel}>{opt.label}</div>
                {showCounts && (
                  <div className={styles.categoryCardCount}>{opt.count} items</div>
                )}
              </button>
            );
          })}
        </div>
      </nav>
    );
  }

  // vertical-rail
  return (
    <nav aria-label="Filter by category">
      <div className={styles.categoryRail}>
        <div className={styles.categoryRailTitle} aria-hidden="true">Categories</div>
        {options.map(opt => {
          const isActive = activeKey === opt.key;
          const railStyle: React.CSSProperties = opt.key === ALL_KEY
            ? (isActive ? ALL_STYLE_ACTIVE : ALL_STYLE_INACTIVE)
            : (enableCategoryColors && categoryColors && categoryColors[opt.key])
              ? isActive
                ? { background: categoryColors[opt.key], color: resolveFg(opt.key, categoryColors[opt.key]) }
                : { borderLeftColor: categoryColors[opt.key], borderLeftWidth: 3, borderLeftStyle: 'solid' }
              : {};
          return (
            <button
              key={opt.key}
              className={`${styles.categoryRailItem} ${isActive ? styles.active : ''}`}
              style={railStyle}
              onClick={() => handleClick(opt.key)}
              aria-pressed={isActive}
              aria-label={`${opt.label}${showCounts ? `, ${opt.count} items` : ''}`}
            >
              <ColorDot catKey={opt.key} />
              <span>{opt.label}</span>
              {showCounts && (
                <span className={styles.categoryRailCount} aria-hidden="true">{opt.count}</span>
              )}
            </button>
          );
        })}
      </div>
    </nav>
  );
};

export default FilterBar;
