import { CellCoord, CellData, CellRange } from '../types/workbook';

export function colIndexToLetter(col: number): string {
  let result = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export function colLetterToIndex(letters: string): number {
  const upper = letters.toUpperCase().trim();
  let res = 0;
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    res = res * 26 + (code - 64);
  }
  return res;
}

export function coordToAddress(coord: CellCoord): string {
  return `${colIndexToLetter(coord.col)}${coord.row}`;
}

export function rangeToAddress(range: CellRange): string {
  const normalized = normalizeRange(range);
  const startStr = coordToAddress(normalized.start);
  const endStr = coordToAddress(normalized.end);
  return startStr === endStr ? startStr : `${startStr}:${endStr}`;
}

export function parseCellAddress(raw: string): CellCoord | null {
  const cleaned = raw.replace(/\$/g, '').trim().toUpperCase();
  const match = /^([A-Z]{1,3})(\d{1,5})$/.exec(cleaned);
  if (!match) return null;
  const col = colLetterToIndex(match[1]);
  const row = parseInt(match[2], 10);
  if (col < 1 || row < 1 || row > 1000) return null;
  return { row, col };
}

export function parseRangeAddress(raw: string): CellRange | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const parts = trimmed.split(':');
  if (parts.length === 1) {
    const c = parseCellAddress(parts[0]);
    if (!c) return null;
    return { start: c, end: c };
  }
  if (parts.length === 2) {
    const c1 = parseCellAddress(parts[0]);
    const c2 = parseCellAddress(parts[1]);
    if (!c1 || !c2) return null;
    return normalizeRange({ start: c1, end: c2 });
  }
  return null;
}

export function normalizeRange(range: CellRange): CellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
}

export function iterateRangeCoords(range: CellRange): CellCoord[] {
  const norm = normalizeRange(range);
  const coords: CellCoord[] = [];
  for (let r = norm.start.row; r <= norm.end.row; r++) {
    for (let c = norm.start.col; c <= norm.end.col; c++) {
      coords.push({ row: r, col: c });
    }
  }
  return coords;
}

export function isCoordInRange(coord: CellCoord, range: CellRange): boolean {
  const norm = normalizeRange(range);
  return (
    coord.row >= norm.start.row &&
    coord.row <= norm.end.row &&
    coord.col >= norm.start.col &&
    coord.col <= norm.end.col
  );
}

export function formatCellDisplayValue(
  value: string | number | boolean | null | undefined,
  numberFormat?: string
): string {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') {
    if (value.startsWith('#')) return value;
    const maybeNum = Number(value);
    if (value.trim() === '' || Number.isNaN(maybeNum)) {
      return value;
    }
    if (!numberFormat || numberFormat === 'General' || numberFormat === '@') {
      return value;
    }
    return formatNumberWithPattern(maybeNum, numberFormat);
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return '#NUM!';
    if (!numberFormat || numberFormat === 'General') {
      return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(6)));
    }
    return formatNumberWithPattern(value, numberFormat);
  }
  return String(value);
}

function formatNumberWithPattern(num: number, rawPattern: string): string {
  const pattern = rawPattern.split(';')[0].trim();
  if (!pattern || pattern === 'General') {
    return String(num);
  }

  // Percentage formats: 0%, 0.0%, 0.00%
  if (pattern.includes('%')) {
    const pct = num * 100;
    const dotIdx = pattern.indexOf('.');
    let decimals = 0;
    if (dotIdx !== -1) {
      const afterDot = pattern.slice(dotIdx + 1).replace(/[^0#]/g, '');
      decimals = afterDot.length;
    }
    return `${pct.toFixed(decimals)}%`;
  }

  // Currency or comma-grouped number formats: $#,##0, $#,##0.00, #,##0, 0.00
  const hasDollar = pattern.includes('$');
  const hasComma = pattern.includes(',');
  const dotIdx = pattern.indexOf('.');
  let decimals = 0;
  if (dotIdx !== -1) {
    const afterDot = pattern.slice(dotIdx + 1).replace(/[^0#]/g, '');
    decimals = afterDot.length;
  }

  const absNum = Math.abs(num);
  const fixed = absNum.toFixed(decimals);
  const [intPart, decPart] = fixed.split('.');

  const formattedInt = hasComma
    ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
    : intPart;

  const numberText = decPart !== undefined ? `${formattedInt}.${decPart}` : formattedInt;
  const sign = num < 0 ? '-' : '';
  return hasDollar ? `${sign}$${numberText}` : `${sign}${numberText}`;
}

export function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (n: number) => Math.max(0, Math.min(255, Math.round(n)));
  const toHex = (n: number) => clamp(n).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

export function vbaColorToCss(val: unknown): string | undefined {
  if (val === null || val === undefined || val === '' || val === -4142 || val === 'xlNone') {
    return undefined;
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed.startsWith('#')) return trimmed;
    const named: Record<string, string> = {
      vbblack: '#000000',
      vbred: '#DC2626',
      vbgreen: '#16A34A',
      vbblue: '#2563EB',
      vbyellow: '#FACC15',
      vbmagenta: '#DB2777',
      vbcyan: '#06B6D4',
      vbwhite: '#FFFFFF',
    };
    const lower = trimmed.toLowerCase();
    if (named[lower]) return named[lower];
    const num = Number(trimmed);
    if (!Number.isNaN(num)) return vbaLongToCss(num);
    return trimmed;
  }
  if (typeof val === 'number') {
    return vbaLongToCss(val);
  }
  return undefined;
}

// VBA RGB(r, g, b) packs R in lowest byte, G in middle, B in high byte: r + g*256 + b*65536
export function vbaLongToCss(colorLong: number): string {
  const n = Math.max(0, Math.floor(colorLong)) & 0xffffff;
  const r = n & 0xff;
  const g = (n >> 8) & 0xff;
  const b = (n >> 16) & 0xff;
  return rgbToHex(r, g, b);
}

export function isNumericCell(cell: CellData | undefined): boolean {
  if (!cell || cell.value === null || cell.value === undefined || cell.value === '') return false;
  if (typeof cell.value === 'number') return Number.isFinite(cell.value);
  if (typeof cell.value === 'string') {
    const n = Number(cell.value);
    return cell.value.trim() !== '' && Number.isFinite(n);
  }
  return false;
}
