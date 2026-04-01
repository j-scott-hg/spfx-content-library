/**
 * Colour utility helpers for category colour coding.
 */

/**
 * Parses a hex colour string (#rrggbb or #rgb) into [r, g, b] 0-255.
 */
function hexToRgb(hex: string): [number, number, number] {
  const clean = hex.replace('#', '');
  if (clean.length === 3) {
    const r = parseInt(clean[0] + clean[0], 16);
    const g = parseInt(clean[1] + clean[1], 16);
    const b = parseInt(clean[2] + clean[2], 16);
    return [r, g, b];
  }
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return [r, g, b];
}

/**
 * Calculates the relative luminance of an sRGB colour (WCAG 2.1 formula).
 */
function relativeLuminance(r: number, g: number, b: number): number {
  const toLinear = (c: number): number => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * toLinear(r) + 0.7152 * toLinear(g) + 0.0722 * toLinear(b);
}

/**
 * Returns '#ffffff' or '#1b1b1b' depending on which gives better contrast
 * against the given background hex colour.
 */
export function getContrastTextColor(bgHex: string): string {
  try {
    const [r, g, b] = hexToRgb(bgHex);
    const lum = relativeLuminance(r, g, b);
    // WCAG contrast ratio: (L1 + 0.05) / (L2 + 0.05)
    // White luminance = 1.0, dark luminance ≈ 0.0
    const contrastWithWhite = (1.05) / (lum + 0.05);
    const contrastWithDark  = (lum + 0.05) / (0.05);
    return contrastWithWhite >= contrastWithDark ? '#ffffff' : '#1b1b1b';
  } catch {
    return '#1b1b1b';
  }
}

/**
 * Creates a lighter tinted version of a hex colour for use as a card background.
 * Blends the colour with white at the given opacity (0–1).
 */
export function tintColor(hex: string, opacity: number): string {
  try {
    const [r, g, b] = hexToRgb(hex);
    const tr = Math.round(r + (255 - r) * (1 - opacity));
    const tg = Math.round(g + (255 - g) * (1 - opacity));
    const tb = Math.round(b + (255 - b) * (1 - opacity));
    return `rgb(${tr}, ${tg}, ${tb})`;
  } catch {
    return '#ffffff';
  }
}

/**
 * A palette of visually distinct colours for auto-assigning to categories.
 */
export const CATEGORY_COLOR_PALETTE = [
  '#0078d4', '#107c41', '#c43e1c', '#7719aa', '#038387',
  '#881798', '#ffb900', '#e3008c', '#00b294', '#d83b01',
  '#004b50', '#8764b8', '#69797e', '#0e7a0d', '#a4262c',
];

/**
 * Auto-assigns palette colours to an array of category keys.
 * Returns a Record<categoryKey, hexColor>.
 */
export function autoAssignCategoryColors(categoryKeys: string[]): Record<string, string> {
  const result: Record<string, string> = {};
  categoryKeys.forEach((key, i) => {
    result[key] = CATEGORY_COLOR_PALETTE[i % CATEGORY_COLOR_PALETTE.length];
  });
  return result;
}
