import * as React from 'react';
import { useState, useCallback, useRef, useEffect } from 'react';
import { autoAssignCategoryColors, getContrastTextColor } from '../../helpers/colorUtils';

// ─── Colour math helpers ──────────────────────────────────────────────────────

function hexToHsv(hex: string): [number, number, number] {
  const clean = hex.replace('#', '');
  const r = parseInt(clean.substring(0, 2), 16) / 255;
  const g = parseInt(clean.substring(2, 4), 16) / 255;
  const b = parseInt(clean.substring(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const d = max - min;
  let h = 0;
  const s = max === 0 ? 0 : d / max;
  const v = max;
  if (d !== 0) {
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  return [h * 360, s, v];
}

function pad2(s: string): string { return s.length < 2 ? `0${s}` : s; }

function hsvToHex(h: number, s: number, v: number): string {
  const hi = Math.floor(h / 60) % 6;
  const f = h / 60 - Math.floor(h / 60);
  const p = v * (1 - s);
  const q = v * (1 - f * s);
  const t = v * (1 - (1 - f) * s);
  let r = 0, g = 0, b = 0;
  if (hi === 0) { r = v; g = t; b = p; }
  else if (hi === 1) { r = q; g = v; b = p; }
  else if (hi === 2) { r = p; g = v; b = t; }
  else if (hi === 3) { r = p; g = q; b = v; }
  else if (hi === 4) { r = t; g = p; b = v; }
  else { r = v; g = p; b = q; }
  const toHex = (n: number): string => pad2(Math.round(n * 255).toString(16));
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function hexToRgb(hex: string): [number, number, number] {
  const clean = hex.replace('#', '');
  return [
    parseInt(clean.substring(0, 2), 16),
    parseInt(clean.substring(2, 4), 16),
    parseInt(clean.substring(4, 6), 16),
  ];
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${pad2(r.toString(16))}${pad2(g.toString(16))}${pad2(b.toString(16))}`;
}

function isValidHex(s: string): boolean {
  return /^#?[0-9a-fA-F]{6}$/.test(s);
}

// ─── Colour swatches matching the screenshots ─────────────────────────────────

const STANDARD_COLORS: string[] = [
  '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0',
  '#0070c0', '#002060', '#7030a0',
];

// ─── Custom HSV picker sub-component ─────────────────────────────────────────

interface IHsvPickerProps {
  hex: string;
  onCommit: (hex: string) => void;
  onCancel: () => void;
}

const HsvPicker: React.FC<IHsvPickerProps> = ({ hex, onCommit, onCancel }) => {
  const [hsv, setHsv] = useState<[number, number, number]>(() => hexToHsv(hex));
  const [hexInput, setHexInput] = useState(hex.replace('#', ''));
  const [rgb, setRgb] = useState<[number, number, number]>(() => hexToRgb(hex));

  const canvasRef = useRef<HTMLDivElement>(null);
  const hueRef = useRef<HTMLDivElement>(null);
  const draggingCanvas = useRef(false);
  const draggingHue = useRef(false);

  const currentHex = hsvToHex(hsv[0], hsv[1], hsv[2]);

  // Sync hex/rgb inputs when hsv changes
  useEffect(() => {
    const h = hsvToHex(hsv[0], hsv[1], hsv[2]);
    setHexInput(h.replace('#', ''));
    setRgb(hexToRgb(h));
  }, [hsv]);

  const getSvFromEvent = (e: MouseEvent | React.MouseEvent): [number, number] => {
    if (!canvasRef.current) return [hsv[1], hsv[2]];
    const rect = canvasRef.current.getBoundingClientRect();
    const x = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    const y = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));
    return [x, 1 - y];
  };

  const getHueFromEvent = (e: MouseEvent | React.MouseEvent): number => {
    if (!hueRef.current) return hsv[0];
    const rect = hueRef.current.getBoundingClientRect();
    const x = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
    return x * 360;
  };

  useEffect(() => {
    const onMove = (e: MouseEvent): void => {
      if (draggingCanvas.current) {
        const [s, v] = getSvFromEvent(e);
        setHsv(prev => [prev[0], s, v]);
      }
      if (draggingHue.current) {
        const h = getHueFromEvent(e);
        setHsv(prev => [h, prev[1], prev[2]]);
      }
    };
    const onUp = (): void => { draggingCanvas.current = false; draggingHue.current = false; };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hsv]);

  const hueColor = hsvToHex(hsv[0], 1, 1);
  const thumbLeft = `${hsv[1] * 100}%`;
  const thumbTop = `${(1 - hsv[2]) * 100}%`;
  const hueLeft = `${(hsv[0] / 360) * 100}%`;

  const handleHexInput = (val: string): void => {
    setHexInput(val);
    const full = val.startsWith('#') ? val : `#${val}`;
    if (isValidHex(full)) {
      const newHsv = hexToHsv(full);
      setHsv(newHsv);
    }
  };

  const handleRgbInput = (channel: 0 | 1 | 2, val: string): void => {
    const n = Math.max(0, Math.min(255, parseInt(val) || 0));
    const newRgb: [number, number, number] = [...rgb] as [number, number, number];
    newRgb[channel] = n;
    setRgb(newRgb);
    setHsv(hexToHsv(rgbToHex(newRgb[0], newRgb[1], newRgb[2])));
  };

  return (
    <div style={{ padding: '12px 0 4px', display: 'flex', flexDirection: 'column', gap: 10 }}>
      {/* HSV canvas */}
      <div
        ref={canvasRef}
        onMouseDown={e => { draggingCanvas.current = true; const [s, v] = getSvFromEvent(e); setHsv(prev => [prev[0], s, v]); }}
        style={{
          position: 'relative',
          width: '100%',
          height: 160,
          borderRadius: 4,
          background: hueColor,
          cursor: 'crosshair',
          flexShrink: 0,
          overflow: 'hidden',
        }}
      >
        {/* White → transparent gradient (left to right) */}
        <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(to right, #fff, transparent)' }} />
        {/* Transparent → black gradient (top to bottom) */}
        <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(to bottom, transparent, #000)' }} />
        {/* Crosshair thumb */}
        <div style={{
          position: 'absolute',
          left: thumbLeft,
          top: thumbTop,
          width: 12,
          height: 12,
          borderRadius: '50%',
          border: '2px solid #fff',
          boxShadow: '0 0 0 1px rgba(0,0,0,0.4)',
          transform: 'translate(-50%, -50%)',
          pointerEvents: 'none',
        }} />
      </div>

      {/* Hue slider + preview swatch */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <div
          ref={hueRef}
          onMouseDown={e => { draggingHue.current = true; setHsv(prev => [getHueFromEvent(e), prev[1], prev[2]]); }}
          style={{
            flex: 1,
            height: 14,
            borderRadius: 7,
            background: 'linear-gradient(to right, #f00, #ff0, #0f0, #0ff, #00f, #f0f, #f00)',
            position: 'relative',
            cursor: 'pointer',
          }}
        >
          <div style={{
            position: 'absolute',
            left: hueLeft,
            top: '50%',
            width: 14,
            height: 14,
            borderRadius: '50%',
            border: '2px solid #fff',
            boxShadow: '0 0 0 1px rgba(0,0,0,0.4)',
            transform: 'translate(-50%, -50%)',
            background: hueColor,
            pointerEvents: 'none',
          }} />
        </div>
        {/* Preview swatch */}
        <div style={{
          width: 40,
          height: 28,
          borderRadius: 4,
          background: currentHex,
          border: '1px solid rgba(0,0,0,0.2)',
          flexShrink: 0,
        }} />
      </div>

      {/* Hex + RGB inputs */}
      <div style={{ display: 'flex', gap: 6, alignItems: 'flex-end' }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: '0 0 90px' }}>
          <label style={{ fontSize: 11, color: '#605e5c' }}>Hex</label>
          <input
            value={hexInput}
            onChange={e => handleHexInput(e.target.value)}
            maxLength={7}
            style={{ width: '100%', padding: '3px 6px', fontSize: 12, border: '1px solid #c8c6c4', borderRadius: 3, fontFamily: 'monospace', boxSizing: 'border-box' }}
          />
        </div>
        {(['Red', 'Green', 'Blue'] as const).map((label, i) => (
          <div key={label} style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1 }}>
            <label style={{ fontSize: 11, color: '#605e5c' }}>{label}</label>
            <input
              type="number"
              min={0}
              max={255}
              value={rgb[i]}
              onChange={e => handleRgbInput(i as 0 | 1 | 2, e.target.value)}
              style={{ width: '100%', padding: '3px 4px', fontSize: 12, border: '1px solid #c8c6c4', borderRadius: 3, boxSizing: 'border-box' }}
            />
          </div>
        ))}
      </div>

      {/* OK / Cancel */}
      <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
        <button
          onClick={onCancel}
          style={{ padding: '5px 14px', fontSize: 13, background: '#fff', border: '1px solid #c8c6c4', borderRadius: 3, cursor: 'pointer', color: '#201f1e' }}
        >
          Cancel
        </button>
        <button
          onClick={() => onCommit(currentHex)}
          style={{ padding: '5px 14px', fontSize: 13, background: '#0078d4', border: 'none', borderRadius: 3, cursor: 'pointer', color: '#fff', fontWeight: 600 }}
        >
          OK
        </button>
      </div>
    </div>
  );
};

// ─── Swatch grid picker (first level) ────────────────────────────────────────

interface ISwatchPickerProps {
  currentHex: string;
  onSelect: (hex: string) => void;
  onMoreColors: () => void;
}

const SwatchPicker: React.FC<ISwatchPickerProps> = ({ currentHex, onSelect, onMoreColors }) => {
  const themeRow1 = ['#000000', '#404040', '#595959', '#7f7f7f', '#a6a6a6', '#d9d9d9', '#1f3864'];
  const themeRow2 = ['#1f497d', '#2e75b6', '#4472c4', '#9dc3e6', '#bdd7ee', '#dae3f3', '#f2f2f2'];

  return (
    <div style={{ padding: '8px 0 4px', display: 'flex', flexDirection: 'column', gap: 8 }}>
      {/* Theme colours */}
      <div style={{ fontSize: 12, fontWeight: 600, color: '#201f1e' }}>Theme colors</div>
      <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
        {[...themeRow1, ...themeRow2].map(c => (
          <button
            key={c}
            onClick={() => onSelect(c)}
            title={c}
            style={{
              width: 22, height: 22, borderRadius: '50%', background: c, padding: 0, cursor: 'pointer',
              border: currentHex.toLowerCase() === c.toLowerCase() ? '2px solid #0078d4' : '1px solid rgba(0,0,0,0.15)',
              boxSizing: 'border-box',
            }}
          />
        ))}
      </div>

      {/* Standard colours */}
      <div style={{ fontSize: 12, fontWeight: 600, color: '#201f1e', marginTop: 2 }}>Standard colors</div>
      <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
        {STANDARD_COLORS.map(c => (
          <button
            key={c}
            onClick={() => onSelect(c)}
            title={c}
            style={{
              width: 22, height: 22, borderRadius: '50%', background: c, padding: 0, cursor: 'pointer',
              border: currentHex.toLowerCase() === c.toLowerCase() ? '2px solid #0078d4' : '1px solid rgba(0,0,0,0.15)',
              boxSizing: 'border-box',
            }}
          />
        ))}
      </div>

      {/* More custom colors */}
      <button
        onClick={onMoreColors}
        style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          padding: '7px 10px', marginTop: 2,
          background: '#fff', border: '1px solid #c8c6c4', borderRadius: 3,
          cursor: 'pointer', fontSize: 13, color: '#201f1e', width: '100%',
        }}
      >
        <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ fontSize: 16 }}>🎨</span>
          More custom colors
        </span>
        <span style={{ fontSize: 12, color: '#605e5c' }}>›</span>
      </button>
    </div>
  );
};

// ─── Main CategoryColorPicker ─────────────────────────────────────────────────

export interface ICategoryColorPickerProps {
  colorMap: Record<string, string>;
  textOverrides: Record<string, 'light' | 'dark'>;
  categoryKeys: string[];
  onChange: (updatedColors: Record<string, string>, updatedTextOverrides: Record<string, 'light' | 'dark'>) => void;
  disabled?: boolean;
}

const CategoryColorPicker: React.FC<ICategoryColorPickerProps> = ({
  colorMap,
  textOverrides,
  categoryKeys,
  onChange,
  disabled,
}) => {
  const autoDefaults = autoAssignCategoryColors(categoryKeys);
  const resolveColors = (): Record<string, string> => {
    const r: Record<string, string> = {};
    categoryKeys.forEach(k => { r[k] = colorMap[k] || autoDefaults[k]; });
    return r;
  };

  const [localColors, setLocalColors] = useState<Record<string, string>>(resolveColors);
  const [localText, setLocalText] = useState<Record<string, 'light' | 'dark'>>(textOverrides);
  // Which category's picker is open, and which level ('swatches' | 'custom')
  const [openKey, setOpenKey] = useState<string | null>(null);
  const [pickerLevel, setPickerLevel] = useState<'swatches' | 'custom'>('swatches');

  const currentKeys = categoryKeys.join(',');
  useEffect(() => {
    setLocalColors(resolveColors());
    setLocalText({ ...textOverrides });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentKeys, JSON.stringify(colorMap), JSON.stringify(textOverrides)]);

  const handleSwatchSelect = useCallback((key: string, hex: string) => {
    const updatedColors = { ...localColors, [key]: hex };
    setLocalColors(updatedColors);
    onChange(updatedColors, localText);
    setOpenKey(null);
  }, [localColors, localText, onChange]);

  const handleCustomCommit = useCallback((key: string, hex: string) => {
    const updatedColors = { ...localColors, [key]: hex };
    setLocalColors(updatedColors);
    onChange(updatedColors, localText);
    setOpenKey(null);
  }, [localColors, localText, onChange]);

  const handleTextToggle = useCallback((key: string) => {
    const currentOverride = localText[key];
    let next: 'light' | 'dark' | undefined;
    if (currentOverride === undefined) {
      next = 'light';
    } else if (currentOverride === 'light') {
      next = 'dark';
    } else {
      next = undefined;
    }
    const updatedText = { ...localText };
    if (next === undefined) {
      delete updatedText[key];
    } else {
      updatedText[key] = next;
    }
    setLocalText(updatedText);
    onChange(localColors, updatedText);
  }, [localColors, localText, onChange]);

  const handleReset = useCallback(() => {
    const reset = autoAssignCategoryColors(categoryKeys);
    setLocalColors(reset);
    setLocalText({});
    setOpenKey(null);
    onChange(reset, {});
  }, [categoryKeys, onChange]);

  const togglePicker = (key: string): void => {
    if (openKey === key) {
      setOpenKey(null);
    } else {
      setOpenKey(key);
      setPickerLevel('swatches');
    }
  };

  if (categoryKeys.length === 0) {
    return (
      <div style={{ fontSize: 12, color: '#605e5c', padding: '4px 0' }}>
        No categories loaded yet. Select a filter field and ensure items are loaded.
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
      {categoryKeys.map(key => {
        const hex = localColors[key] || '#0078d4';
        const swatchFg = getContrastTextColor(hex);
        const textOverride = localText[key];
        const autoText = getContrastTextColor(hex);
        const isOpen = openKey === key;

        // Text toggle appearance — A = auto, ☀ = forced white, 🌙 = forced dark
        let toggleLabel: string;
        let toggleTitle: string;
        let toggleBg: string;
        let toggleFg: string;
        if (textOverride === 'light') {
          toggleLabel = '☀';
          toggleTitle = 'Text: forced white — click for forced dark';
          toggleBg = '#323130';
          toggleFg = '#ffffff';
        } else if (textOverride === 'dark') {
          toggleLabel = '🌙';
          toggleTitle = 'Text: forced dark — click to reset to auto';
          toggleBg = '#f3f2f1';
          toggleFg = '#201f1e';
        } else {
          toggleLabel = 'A';
          toggleTitle = `Text: auto (${autoText === '#ffffff' ? 'white' : 'dark'}) — click to force white`;
          toggleBg = autoText === '#ffffff' ? '#323130' : '#f3f2f1';
          toggleFg = autoText === '#ffffff' ? '#ffffff' : '#201f1e';
        }

        return (
          <div
            key={key}
            style={{
              opacity: disabled ? 0.5 : 1,
              pointerEvents: disabled ? 'none' : 'auto',
            }}
          >
            {/* Row: swatch | text toggle | label | hex */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '3px 0' }}>
              {/* Colour swatch button — opens picker */}
              <button
                onClick={() => togglePicker(key)}
                title={`Pick colour for "${key}"`}
                aria-label={`Colour for ${key}: ${hex}. Click to open picker.`}
                aria-expanded={isOpen}
                style={{
                  width: 28,
                  height: 22,
                  borderRadius: 4,
                  background: hex,
                  border: isOpen ? '2px solid #0078d4' : '1px solid rgba(0,0,0,0.25)',
                  cursor: 'pointer',
                  flexShrink: 0,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: 0,
                }}
              >
                <span style={{ fontSize: 9, color: swatchFg, userSelect: 'none', lineHeight: 1 }}>▼</span>
              </button>

              {/* Light/dark text toggle */}
              <button
                onClick={() => handleTextToggle(key)}
                title={toggleTitle}
                aria-label={toggleTitle}
                style={{
                  width: 22,
                  height: 22,
                  borderRadius: 3,
                  background: toggleBg,
                  border: '1px solid rgba(0,0,0,0.2)',
                  cursor: 'pointer',
                  padding: 0,
                  flexShrink: 0,
                  fontSize: 11,
                  fontWeight: 700,
                  color: toggleFg,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  lineHeight: 1,
                }}
              >
                {toggleLabel}
              </button>

              {/* Category label */}
              <span style={{ fontSize: 13, color: '#201f1e', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {key}
              </span>

              {/* Hex value */}
              <span style={{ fontSize: 11, color: '#605e5c', fontFamily: 'monospace', flexShrink: 0 }}>
                {hex.toUpperCase()}
              </span>
            </div>

            {/* Inline picker panel */}
            {isOpen && (
              <div style={{
                margin: '0 0 6px 34px',
                padding: '0 10px 10px',
                background: '#fff',
                border: '1px solid #c8c6c4',
                borderRadius: 4,
                boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
              }}>
                {pickerLevel === 'swatches' ? (
                  <SwatchPicker
                    currentHex={hex}
                    onSelect={h => handleSwatchSelect(key, h)}
                    onMoreColors={() => setPickerLevel('custom')}
                  />
                ) : (
                  <HsvPicker
                    hex={hex}
                    onCommit={h => handleCustomCommit(key, h)}
                    onCancel={() => setPickerLevel('swatches')}
                  />
                )}
              </div>
            )}
          </div>
        );
      })}

      {/* Legend + Reset */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: 4 }}>
        <span style={{ fontSize: 11, color: '#605e5c' }}>
          A = auto &nbsp;☀ = white &nbsp;🌙 = dark
        </span>
        <button
          onClick={handleReset}
          disabled={disabled}
          style={{
            padding: '3px 10px', fontSize: 12, background: 'transparent',
            border: '1px solid #c8c6c4', borderRadius: 3, cursor: 'pointer', color: '#605e5c',
          }}
        >
          Reset
        </button>
      </div>
    </div>
  );
};

export default CategoryColorPicker;
