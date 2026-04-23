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

function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (n: number): string => {
    const x = Math.max(0, Math.min(255, Math.round(n)));
    const h = x.toString(16);
    return h.length === 1 ? `0${h}` : h;
  };
  return `#${clamp(r)}${clamp(g)}${clamp(b)}`;
}

/** Blend hex toward black; strength 0–1 (e.g. 0.18). */
function mixTowardBlack(hex: string, strength: number): string {
  try {
    const [r, g, b] = hexToRgb(hex);
    const t = Math.max(0, Math.min(1, strength));
    return rgbToHex(r * (1 - t * 0.92), g * (1 - t * 0.92), b * (1 - t * 0.92));
  } catch {
    return '#2d5080';
  }
}

const DEFAULT_ACCENT_FALLBACK = '#3c6aa7';

/**
 * Normalises author-entered hex (#rgb, #rrggbb, or rrggbb) for web part accent colour.
 */
export function normalizeAccentHex(input: string | undefined, fallback: string = DEFAULT_ACCENT_FALLBACK): string {
  let t = (input || '').trim();
  if (!t) return fallback;
  t = t.replace(/^#/, '');
  if (/^[0-9A-Fa-f]{3}$/.test(t)) {
    t = `${t[0]}${t[0]}${t[1]}${t[1]}${t[2]}${t[2]}`;
  }
  if (!/^[0-9A-Fa-f]{6}$/.test(t)) return fallback;
  const full = `#${t.toLowerCase()}`;
  try {
    const [r, g, b] = hexToRgb(full);
    if ([r, g, b].some(n => Number.isNaN(n))) return fallback;
  } catch {
    return fallback;
  }
  return full;
}

/**
 * CSS custom properties for the web part root (links, active filters, focus, primary buttons).
 */
export function getAccentCssProperties(accentInput: string | undefined): Record<string, string> {
  const primary = normalizeAccentHex(accentInput);
  const primaryDark = mixTowardBlack(primary, 0.22);
  const primaryLight = tintColor(primary, 0.22);
  return {
    '--cl-primary': primary,
    '--cl-primary-dark': primaryDark,
    '--cl-primary-light': primaryLight,
    '--cl-focus': primary,
  };
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
