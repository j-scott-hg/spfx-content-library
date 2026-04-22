const BADGE_PALETTE: Array<{ bg: string; fg: string; border: string }> = [
  { bg: '#ede6ff', fg: '#4f2a86', border: '#d1c0f6' },
  { bg: '#e8f2ff', fg: '#0f548c', border: '#bfdcff' },
  { bg: '#e7f9ed', fg: '#107c10', border: '#bfe7c8' },
  { bg: '#fff4df', fg: '#8a5d00', border: '#f5deaa' },
  { bg: '#fde7e9', fg: '#a4262c', border: '#f4c7cb' },
  { bg: '#e7f7f8', fg: '#006c70', border: '#b9e4e7' },
  { bg: '#f2ecff', fg: '#5c2e91', border: '#ddd0ff' },
  { bg: '#f3f2f1', fg: '#323130', border: '#d2d0ce' },
];

function hashString(input: string): number {
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = ((hash << 5) - hash) + input.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash);
}

export function normalizeChoiceValues(raw: unknown): string[] {
  if (raw === null || raw === undefined || raw === '') return [];
  if (Array.isArray(raw)) {
    return raw.map(v => String(v).trim()).filter(Boolean);
  }
  if (typeof raw === 'object') {
    const obj = raw as { results?: unknown[]; value?: unknown; Label?: unknown };
    if (Array.isArray(obj.results)) {
      return obj.results.map(v => String(v).trim()).filter(Boolean);
    }
    if (obj.value !== undefined && obj.value !== null) {
      return [String(obj.value).trim()].filter(Boolean);
    }
    if (obj.Label !== undefined && obj.Label !== null) {
      return [String(obj.Label).trim()].filter(Boolean);
    }
  }

  const str = String(raw).trim();
  if (!str) return [];
  if (str.indexOf(';#') !== -1) {
    return str.split(';#').map(v => v.trim()).filter(Boolean);
  }
  if (str.indexOf(',') !== -1) {
    return str.split(',').map(v => v.trim()).filter(Boolean);
  }
  return [str];
}

export function getChoiceBadgeStyle(value: string): { background: string; color: string; borderColor: string } {
  const normalized = value.trim().toLowerCase();
  const palette = BADGE_PALETTE[hashString(normalized) % BADGE_PALETTE.length];
  return {
    background: palette.bg,
    color: palette.fg,
    borderColor: palette.border,
  };
}
