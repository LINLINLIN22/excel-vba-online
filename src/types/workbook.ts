export type HorizontalAlign = 'left' | 'center' | 'right';

export interface CellStyle {
  backgroundColor?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  horizontalAlign?: HorizontalAlign;
  wrapText?: boolean;
  numberFormat?: string;
  borderColor?: string;
  borderStyle?: 'thin' | 'medium' | 'none';
}

export interface CellData {
  value: string | number | boolean | null;
  formula?: string;
  style?: CellStyle;
}

export interface SheetData {
  id: string;
  name: string;
  rows: number;
  cols: number;
  cells: Record<string, CellData>;
  colWidths: Record<number, number>;
}

export interface CellCoord {
  row: number; // 1-indexed
  col: number; // 1-indexed
}

export interface CellRange {
  start: CellCoord;
  end: CellCoord;
}

export interface VbaModule {
  id: string;
  name: string;
  code: string;
}

export interface VbaDiagnostic {
  line: number; // 1-indexed
  col: number;
  severity: 'error' | 'warning';
  message: string;
}

export interface ExecutionLogEntry {
  id: string;
  type: 'print' | 'msgbox' | 'info' | 'error';
  message: string;
  line?: number;
}

export interface ExecutionResult {
  status: 'idle' | 'success' | 'error';
  headline: string;
  detail: string;
  durationMs: number;
  logs: ExecutionLogEntry[];
  changedCellsCount: number;
  errorLine?: number;
}

export interface ExamplePreset {
  id: string;
  title: string;
  category: string;
  difficulty: '初階' | '中階';
  description: string;
  tipTitle: string;
  tipBody: string;
  code: string;
  sheets: SheetData[];
  selectedRange?: CellRange;
}

export interface LearnLesson {
  id: string;
  number: string;
  title: string;
  summary: string;
  concept: string;
  code: string;
  keyTakeaways: string[];
  presetId?: string;
}
