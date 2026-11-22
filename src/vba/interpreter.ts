import {
  colLetterToIndex,
  coordToAddress,
  iterateRangeCoords,
  normalizeRange,
  parseRangeAddress,
  rangeToAddress,
  vbaColorToCss,
} from '../spreadsheet/addressing';
import { recalculateSheets } from '../spreadsheet/formulaEngine';
import {
  CellData,
  CellRange,
  CellStyle,
  ExecutionLogEntry,
  ExecutionResult,
  HorizontalAlign,
  SheetData,
} from '../types/workbook';
import { tokenizeVbaLine, VbaToken } from './lexer';
import { lintVbaCode } from './linter';

interface RangeRef {
  kind: 'range';
  sheetName: string;
  range: CellRange;
}

interface SheetRef {
  kind: 'sheet';
  sheetName: string;
}

interface SubProxy {
  kind: 'proxy';
  target: RangeRef;
  subName: 'interior' | 'font' | 'borders';
}

type VbaValue =
  | string
  | number
  | boolean
  | null
  | undefined
  | RangeRef
  | SheetRef
  | SubProxy
  | VbaValue[];

interface ProcedureDef {
  name: string;
  kind: 'sub' | 'function';
  params: string[];
  startLine: number; // index in lines[]
  endLine: number;
}

class VbaRuntimeError extends Error {
  constructor(
    message: string,
    public line: number
  ) {
    super(message);
  }
}

class ExitSignal {
  constructor(public kind: 'sub' | 'function' | 'for' | 'do') {}
}

interface ExecLine {
  lineNum: number;
  tokens: VbaToken[];
}

export interface RunMacroInput {
  code: string;
  sheets: SheetData[];
  activeSheetId: string;
  selectedRange: CellRange;
}

export interface RunMacroOutput {
  sheets: SheetData[];
  activeSheetId: string;
  selectedRange: CellRange;
  result: ExecutionResult;
}

const MAX_STEPS = 50000;

export function executeVbaMacro(input: RunMacroInput): RunMacroOutput {
  const startTime = performance.now();

  // Deep clone sheets so failed runs don't corrupt state unless desired; actually VBA commits until error,
  // but cloning cleanly prevents React mutation bugs
  let sheets: SheetData[] = input.sheets.map((s) => ({
    ...s,
    colWidths: { ...s.colWidths },
    cells: Object.fromEntries(
      Object.entries(s.cells).map(([k, v]) => [
        k,
        { ...v, style: v.style ? { ...v.style } : undefined },
      ])
    ),
  }));

  let activeSheetId = input.activeSheetId;
  let selectedRange: CellRange = {
    start: { ...input.selectedRange.start },
    end: { ...input.selectedRange.end },
  };

  // 1. Run static linter first
  const diagnostics = lintVbaCode(input.code);
  const firstError = diagnostics.find((d) => d.severity === 'error');
  if (firstError) {
    return {
      sheets: input.sheets,
      activeSheetId: input.activeSheetId,
      selectedRange: input.selectedRange,
      result: {
        status: 'error',
        headline: `Syntax error on line ${firstError.line}`,
        detail: firstError.message,
        durationMs: Math.round(performance.now() - startTime),
        logs: [
          {
            id: `err-${Date.now()}`,
            type: 'error',
            line: firstError.line,
            message: `Line ${firstError.line}: ${firstError.message}`,
          },
        ],
        changedCellsCount: 0,
        errorLine: firstError.line,
      },
    };
  }

  // 2. Parse procedures
  const rawLines = input.code.split('\n');
  const parsedLines: ExecLine[] = rawLines.map((raw, idx) => ({
    lineNum: idx + 1,
    tokens: tokenizeVbaLine(raw, idx + 1).filter(
      (t) => t.type !== 'whitespace' && t.type !== 'comment'
    ),
  }));

  const procedures = new Map<string, ProcedureDef>();
  let firstSub: ProcedureDef | null = null;
  let currentProc: ProcedureDef | null = null;

  for (let i = 0; i < parsedLines.length; i++) {
    const toks = parsedLines[i].tokens;
    if (toks.length === 0) continue;
    const w0 = toks[0].value.toLowerCase();
    const w1 = toks[1]?.value.toLowerCase() ?? '';

    const isDecl =
      w0 === 'sub' ||
      w0 === 'function' ||
      ((w0 === 'public' || w0 === 'private') && (w1 === 'sub' || w1 === 'function'));

    if (isDecl) {
      const idx = w0 === 'public' || w0 === 'private' ? 1 : 0;
      const kind = toks[idx].value.toLowerCase() as 'sub' | 'function';
      const nameTok = toks[idx + 1];
      if (!nameTok) {
        return makeErrorResult(input, startTime, parsedLines[i].lineNum, `Expected procedure name after ${kind}.`);
      }
      const params: string[] = [];
      const openIdx = toks.findIndex((t) => t.value === '(');
      const closeIdx = toks.findIndex((t) => t.value === ')');
      if (openIdx !== -1 && closeIdx > openIdx + 1) {
        let p = openIdx + 1;
        while (p < closeIdx) {
          const val = toks[p].value.toLowerCase();
          if (val === 'byval' || val === 'byref') {
            p++;
            continue;
          }
          if (toks[p].type === 'ident' || toks[p].type === 'keyword') {
            params.push(toks[p].value.toLowerCase());
            // skip "As Type" and comma
            while (p < closeIdx && toks[p].value !== ',') p++;
          }
          p++;
        }
      }
      currentProc = {
        name: nameTok.value.toLowerCase(),
        kind,
        params,
        startLine: i + 1,
        endLine: parsedLines.length - 1,
      };
      if (kind === 'sub' && !firstSub) {
        firstSub = currentProc;
      }
      procedures.set(currentProc.name, currentProc);
      continue;
    }

    if (currentProc && w0 === 'end' && (w1 === 'sub' || w1 === 'function')) {
      currentProc.endLine = i - 1;
      currentProc = null;
    }
  }

  if (!firstSub) {
    return makeErrorResult(
      input,
      startTime,
      1,
      'No runnable Sub procedure found. Wrap your code in Sub MyMacro() ... End Sub.'
    );
  }

  const logs: ExecutionLogEntry[] = [];
  const changedCells = new Set<string>();
  let stepCount = 0;

  const getActiveSheet = (): SheetData => {
    return sheets.find((s) => s.id === activeSheetId) ?? sheets[0];
  };

  const getSheetByNameOrIndex = (arg: VbaValue, lineNum: number): SheetData => {
    if (typeof arg === 'number') {
      const idx = Math.floor(arg) - 1;
      if (idx < 0 || idx >= sheets.length) {
        throw new VbaRuntimeError(`Subscript out of range: Sheets(${arg}) does not exist.`, lineNum);
      }
      return sheets[idx];
    }
    const name = String(arg ?? '').trim();
    const found = sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
    if (!found) {
      throw new VbaRuntimeError(`Subscript out of range: Sheet "${name}" not found.`, lineNum);
    }
    return found;
  };

  const ensureSheetSize = (sheet: SheetData, row: number, col: number) => {
    if (row > sheet.rows) sheet.rows = Math.min(200, row);
    if (col > sheet.cols) sheet.cols = Math.min(52, col);
  };

  const readCell = (sheetName: string, row: number, col: number): CellData | undefined => {
    const sheet = sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
    if (!sheet) return undefined;
    const addr = coordToAddress({ row, col });
    return sheet.cells[addr];
  };

  const writeCell = (
    sheetName: string,
    row: number,
    col: number,
    updater: (prev: CellData) => CellData
  ) => {
    const sheet = sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
    if (!sheet) return;
    ensureSheetSize(sheet, row, col);
    const addr = coordToAddress({ row, col });
    const prev: CellData = sheet.cells[addr] ?? { value: null };
    sheet.cells[addr] = updater({
      ...prev,
      style: prev.style ? { ...prev.style } : undefined,
    });
    changedCells.add(`${sheet.name}!${addr}`);
  };

  const resolveRangeValue = (ref: RangeRef): VbaValue => {
    const cell = readCell(ref.sheetName, ref.range.start.row, ref.range.start.col);
    return cell?.value ?? null;
  };

  const coercePrimitive = (val: VbaValue): string | number | boolean | null => {
    if (val && typeof val === 'object' && !Array.isArray(val) && val.kind === 'range') {
      const cell = readCell(val.sheetName, val.range.start.row, val.range.start.col);
      return cell?.value ?? null;
    }
    if (val === undefined) return null;
    if (Array.isArray(val) || (val && typeof val === 'object')) return String(val);
    return val;
  };

  const toNumber = (val: VbaValue): number => {
    const prim = coercePrimitive(val);
    if (typeof prim === 'number') return prim;
    if (typeof prim === 'boolean') return prim ? 1 : 0;
    if (typeof prim === 'string') {
      const cleaned = prim.replace(/[$,]/g, '').trim();
      if (!cleaned) return 0;
      const n = Number(cleaned);
      return Number.isFinite(n) ? n : 0;
    }
    return 0;
  };

  const toStr = (val: VbaValue): string => {
    const prim = coercePrimitive(val);
    if (prim === null || prim === undefined) return '';
    return String(prim);
  };

  const toBool = (val: VbaValue): boolean => {
    const prim = coercePrimitive(val);
    return Boolean(prim);
  };

  const getProperty = (
    target: VbaValue,
    prop: string,
    args: VbaValue[] | null,
    lineNum: number
  ): VbaValue => {
    const p = prop.toLowerCase();

    // Debug.Print is handled as statement; WorksheetFunction
    if (target && typeof target === 'object' && !Array.isArray(target)) {
      if (target.kind === 'sheet') {
        const sheet = getSheetByNameOrIndex(target.sheetName, lineNum);
        if (p === 'range') {
          return buildRangeFromArgs(sheet.name, args ?? [], lineNum);
        }
        if (p === 'cells') {
          return buildCellsFromArgs(sheet, args ?? [], lineNum);
        }
        if (p === 'rows') {
          const r = args && args.length > 0 ? Math.max(1, Math.floor(toNumber(args[0]))) : 1;
          return {
            kind: 'range',
            sheetName: sheet.name,
            range: { start: { row: r, col: 1 }, end: { row: r, col: sheet.cols } },
          };
        }
        if (p === 'columns') {
          const c = args && args.length > 0 ? resolveColArg(args[0], lineNum) : 1;
          return {
            kind: 'range',
            sheetName: sheet.name,
            range: { start: { row: 1, col: c }, end: { row: sheet.rows, col: c } },
          };
        }
        if (p === 'name') return sheet.name;
        if (p === 'activate' || p === 'select') {
          activeSheetId = sheet.id;
          return null;
        }
      }

      if (target.kind === 'range') {
        if (p === 'value' || p === 'value2' || p === 'text') {
          return resolveRangeValue(target);
        }
        if (p === 'formula') {
          const cell = readCell(target.sheetName, target.range.start.row, target.range.start.col);
          return cell?.formula ?? cell?.value ?? '';
        }
        if (p === 'numberformat') {
          const cell = readCell(target.sheetName, target.range.start.row, target.range.start.col);
          return cell?.style?.numberFormat ?? 'General';
        }
        if (p === 'interior') return { kind: 'proxy', target, subName: 'interior' };
        if (p === 'font') return { kind: 'proxy', target, subName: 'font' };
        if (p === 'borders') return { kind: 'proxy', target, subName: 'borders' };
        if (p === 'row') return target.range.start.row;
        if (p === 'column') return target.range.start.col;
        if (p === 'address') return rangeToAddress(target.range);
        if (p === 'count') {
          const norm = normalizeRange(target.range);
          return (norm.end.row - norm.start.row + 1) * (norm.end.col - norm.start.col + 1);
        }
        if (p === 'offset') {
          const rOff = Math.trunc(toNumber(args?.[0] ?? 0));
          const cOff = Math.trunc(toNumber(args?.[1] ?? 0));
          const nr: CellRange = {
            start: {
              row: Math.max(1, target.range.start.row + rOff),
              col: Math.max(1, target.range.start.col + cOff),
            },
            end: {
              row: Math.max(1, target.range.end.row + rOff),
              col: Math.max(1, target.range.end.col + cOff),
            },
          };
          return { kind: 'range', sheetName: target.sheetName, range: nr };
        }
        if (p === 'resize') {
          const rSize = Math.max(1, Math.trunc(toNumber(args?.[0] ?? 1)));
          const cSize = Math.max(1, Math.trunc(toNumber(args?.[1] ?? 1)));
          return {
            kind: 'range',
            sheetName: target.sheetName,
            range: {
              start: { ...target.range.start },
              end: {
                row: target.range.start.row + rSize - 1,
                col: target.range.start.col + cSize - 1,
              },
            },
          };
        }
        if (p === 'cells') {
          const r = Math.max(1, Math.trunc(toNumber(args?.[0] ?? 1)));
          const c = Math.max(1, Math.trunc(toNumber(args?.[1] ?? 1)));
          const absRow = target.range.start.row + r - 1;
          const absCol = target.range.start.col + c - 1;
          return {
            kind: 'range',
            sheetName: target.sheetName,
            range: { start: { row: absRow, col: absCol }, end: { row: absRow, col: absCol } },
          };
        }
        if (p === 'end') {
          const dir = String(args?.[0] ?? 'xldown').toLowerCase();
          const sheet = getSheetByNameOrIndex(target.sheetName, lineNum);
          let r = target.range.start.row;
          let c = target.range.start.col;
          const hasVal = (rr: number, cc: number) => {
            const cell = sheet.cells[coordToAddress({ row: rr, col: cc })];
            return cell !== undefined && cell.value !== null && cell.value !== '';
          };
          if (dir === 'xlup' || dir === '-4162') {
            let scan = Math.min(r, sheet.rows);
            while (scan > 1 && !hasVal(scan, c)) scan--;
            r = scan;
          } else if (dir === 'xldown' || dir === '-4121') {
            let scan = r;
            while (scan < sheet.rows && hasVal(scan + 1, c)) scan++;
            r = scan;
          } else if (dir === 'xltoleft' || dir === '-4159') {
            let scan = Math.min(c, sheet.cols);
            while (scan > 1 && !hasVal(r, scan)) scan--;
            c = scan;
          } else if (dir === 'xltoright' || dir === '-4161') {
            let scan = c;
            while (scan < sheet.cols && hasVal(r, scan + 1)) scan++;
            c = scan;
          }
          return {
            kind: 'range',
            sheetName: target.sheetName,
            range: { start: { row: r, col: c }, end: { row: r, col: c } },
          };
        }
        if (p === 'clear' || p === 'clearcontents' || p === 'clearformats') {
          applyRangeMethod(target, p);
          return null;
        }
        if (p === 'select') {
          const targetSheet = getSheetByNameOrIndex(target.sheetName, lineNum);
          activeSheetId = targetSheet.id;
          selectedRange = normalizeRange(target.range);
          return null;
        }
      }

      if (target.kind === 'proxy') {
        if (target.subName === 'interior' && p === 'color') {
          const cell = readCell(target.target.sheetName, target.target.range.start.row, target.target.range.start.col);
          return cell?.style?.backgroundColor ?? null;
        }
        if (target.subName === 'font') {
          const cell = readCell(target.target.sheetName, target.target.range.start.row, target.target.range.start.col);
          if (p === 'color') return cell?.style?.color ?? null;
          if (p === 'bold') return Boolean(cell?.style?.bold);
          if (p === 'italic') return Boolean(cell?.style?.italic);
          if (p === 'underline') return Boolean(cell?.style?.underline);
          if (p === 'size') return cell?.style?.fontSize ?? 11;
          if (p === 'name') return cell?.style?.fontFamily ?? 'Calibri';
        }
      }
    }

    if (target === '__SHEETS_COLLECTION__') {
      if (p === 'add') {
        let idx = sheets.length + 1;
        let newName = `Sheet${idx}`;
        while (sheets.some((s) => s.name.toLowerCase() === newName.toLowerCase())) {
          idx++;
          newName = `Sheet${idx}`;
        }
        const newSheet: SheetData = {
          id: `sheet-${Date.now()}-${sheets.length + 1}`,
          name: newName,
          rows: 15,
          cols: 7,
          cells: {},
          colWidths: { 1: 102, 2: 76, 3: 84, 4: 74, 5: 74, 6: 74, 7: 74 },
        };
        sheets.push(newSheet);
        activeSheetId = newSheet.id;
        return { kind: 'sheet', sheetName: newSheet.name };
      }
      if (p === 'count') {
        return sheets.length;
      }
    }

    throw new VbaRuntimeError(`Unsupported property or method ".${prop}".`, lineNum);
  };

  const applyRangeMethod = (target: RangeRef, method: string) => {
    const coords = iterateRangeCoords(target.range);
    for (const c of coords) {
      writeCell(target.sheetName, c.row, c.col, (prev) => {
        if (method === 'clear') return { value: null };
        if (method === 'clearcontents') return { ...prev, value: null, formula: undefined };
        if (method === 'clearformats') return { value: prev.value, formula: prev.formula };
        return prev;
      });
    }
  };

  const setProperty = (
    target: VbaValue,
    propPath: string[],
    value: VbaValue,
    lineNum: number
  ) => {
    if (!target || typeof target !== 'object' || Array.isArray(target)) {
      throw new VbaRuntimeError(`Object required when assigning ".${propPath.join('.')}".`, lineNum);
    }

    let rangeTarget: RangeRef | null = null;
    let normalizedPath = propPath.map((s) => s.toLowerCase());

    if (target.kind === 'sheet') {
      const sheet = getSheetByNameOrIndex(target.sheetName, lineNum);
      const sheetProp = propPath.map((s) => s.toLowerCase()).join('.');
      if (sheetProp === 'name') {
        const nextName = String(coercePrimitive(value) ?? '').trim();
        if (!nextName) {
          throw new VbaRuntimeError('Worksheet name cannot be empty.', lineNum);
        }
        const conflict = sheets.find(
          (s) => s.id !== sheet.id && s.name.toLowerCase() === nextName.toLowerCase()
        );
        if (conflict) {
          throw new VbaRuntimeError(`Worksheet "${nextName}" already exists.`, lineNum);
        }
        sheet.name = nextName;
        return;
      }
      throw new VbaRuntimeError(`Cannot assign ".${propPath.join('.')}" on Worksheet.`, lineNum);
    }

    if (target.kind === 'range') {
      rangeTarget = target;
    } else if (target.kind === 'proxy') {
      rangeTarget = target.target;
      normalizedPath = [target.subName, ...normalizedPath];
    }

    if (!rangeTarget) {
      throw new VbaRuntimeError(`Cannot assign ".${propPath.join('.')}" on this object.`, lineNum);
    }

    const joined = normalizedPath.join('.');
    const coords = iterateRangeCoords(rangeTarget.range);
    const prim = coercePrimitive(value);

    for (const coord of coords) {
      writeCell(rangeTarget.sheetName, coord.row, coord.col, (prev) => {
        const nextStyle: CellStyle = { ...(prev.style ?? {}) };
        if (joined === 'value' || joined === 'value2') {
          if (typeof prim === 'string' && prim.startsWith('=')) {
            return { ...prev, formula: prim, value: prim, style: nextStyle };
          }
          return { ...prev, value: prim, formula: undefined, style: nextStyle };
        }
        if (joined === 'formula') {
          const f = prim === null ? '' : String(prim);
          return {
            ...prev,
            formula: f.startsWith('=') ? f : `=${f}`,
            value: f,
            style: nextStyle,
          };
        }
        if (joined === 'numberformat') {
          nextStyle.numberFormat = prim ? String(prim) : 'General';
          return { ...prev, style: nextStyle };
        }
        if (joined === 'interior.color' || joined === 'interior.colorindex') {
          nextStyle.backgroundColor = vbaColorToCss(prim);
          return { ...prev, style: nextStyle };
        }
        if (joined === 'font.color' || joined === 'font.colorindex') {
          nextStyle.color = vbaColorToCss(prim);
          return { ...prev, style: nextStyle };
        }
        if (joined === 'font.bold') {
          nextStyle.bold = Boolean(prim);
          return { ...prev, style: nextStyle };
        }
        if (joined === 'font.italic') {
          nextStyle.italic = Boolean(prim);
          return { ...prev, style: nextStyle };
        }
        if (joined === 'font.underline') {
          nextStyle.underline = Boolean(prim);
          return { ...prev, style: nextStyle };
        }
        if (joined === 'font.size') {
          const sz = Number(prim);
          if (Number.isFinite(sz) && sz >= 6 && sz <= 48) {
            nextStyle.fontSize = sz;
          }
          return { ...prev, style: nextStyle };
        }
        if (joined === 'font.name') {
          nextStyle.fontFamily = prim ? String(prim) : 'Calibri';
          return { ...prev, style: nextStyle };
        }
        if (joined === 'horizontalalignment') {
          const s = String(prim ?? '').toLowerCase();
          let align: HorizontalAlign = 'left';
          if (s === 'xlcenter' || s === 'center' || s === '-4108') align = 'center';
          else if (s === 'xlright' || s === 'right' || s === '-4152') align = 'right';
          nextStyle.horizontalAlign = align;
          return { ...prev, style: nextStyle };
        }
        if (joined === 'wraptext') {
          nextStyle.wrapText = Boolean(prim);
          return { ...prev, style: nextStyle };
        }
        if (joined === 'borders.linestyle') {
          const s = String(prim ?? '').toLowerCase();
          nextStyle.borderStyle = s === 'xlnone' || s === '-4142' || s === '0' ? 'none' : 'thin';
          return { ...prev, style: nextStyle };
        }
        if (joined === 'borders.color') {
          nextStyle.borderColor = vbaColorToCss(prim);
          nextStyle.borderStyle = nextStyle.borderStyle ?? 'thin';
          return { ...prev, style: nextStyle };
        }
        if (joined === 'columnwidth') {
          const sheet = sheets.find(
            (sh) => sh.name.toLowerCase() === rangeTarget!.sheetName.toLowerCase()
          );
          const w = Number(prim);
          if (sheet && Number.isFinite(w) && w > 0) {
            sheet.colWidths[coord.col] = Math.max(48, Math.min(320, Math.round(w * 8)));
          }
          return prev;
        }
        throw new VbaRuntimeError(`Unsupported property assignment ".${propPath.join('.')}".`, lineNum);
      });
    }

    // Immediately recalculate formulas if any value or formula changed
    if (joined === 'value' || joined === 'value2' || joined === 'formula') {
      sheets = recalculateSheets(sheets);
    }
  };

  const resolveColArg = (arg: VbaValue, lineNum: number): number => {
    if (typeof arg === 'number') return Math.max(1, Math.trunc(arg));
    const s = String(arg ?? '').trim();
    const num = Number(s);
    if (s !== '' && Number.isFinite(num)) return Math.max(1, Math.trunc(num));
    const idx = colLetterToIndex(s);
    if (idx < 1) {
      throw new VbaRuntimeError(`Invalid column reference "${s}".`, lineNum);
    }
    return idx;
  };

  const buildRangeFromArgs = (sheetName: string, args: VbaValue[], lineNum: number): RangeRef => {
    if (args.length === 1) {
      const addrStr = String(coercePrimitive(args[0]) ?? '').trim();
      const parsed = parseRangeAddress(addrStr);
      if (!parsed) {
        throw new VbaRuntimeError(`Method 'Range' failed: invalid address "${addrStr}".`, lineNum);
      }
      return { kind: 'range', sheetName, range: parsed };
    }
    if (args.length === 2) {
      const r1 = args[0];
      const r2 = args[1];
      const addr1 =
        r1 && typeof r1 === 'object' && !Array.isArray(r1) && r1.kind === 'range'
          ? r1.range.start
          : parseRangeAddress(String(coercePrimitive(r1) ?? ''))?.start;
      const addr2 =
        r2 && typeof r2 === 'object' && !Array.isArray(r2) && r2.kind === 'range'
          ? r2.range.end
          : parseRangeAddress(String(coercePrimitive(r2) ?? ''))?.end;
      if (!addr1 || !addr2) {
        throw new VbaRuntimeError(`Invalid Range(cell1, cell2) arguments.`, lineNum);
      }
      return {
        kind: 'range',
        sheetName,
        range: normalizeRange({ start: addr1, end: addr2 }),
      };
    }
    throw new VbaRuntimeError(`Range() expects 1 or 2 arguments.`, lineNum);
  };

  const buildCellsFromArgs = (sheet: SheetData, args: VbaValue[], lineNum: number): RangeRef => {
    if (args.length === 0) {
      return {
        kind: 'range',
        sheetName: sheet.name,
        range: { start: { row: 1, col: 1 }, end: { row: sheet.rows, col: sheet.cols } },
      };
    }
    if (args.length !== 2) {
      throw new VbaRuntimeError(`Cells(row, col) expects 2 arguments.`, lineNum);
    }
    const row = Math.trunc(toNumber(args[0]));
    const col = resolveColArg(coercePrimitive(args[1]), lineNum);
    if (row < 1 || col < 1) {
      throw new VbaRuntimeError(`Application-defined error: Cells(${row}, ${col}) row and column must be >= 1.`, lineNum);
    }
    return {
      kind: 'range',
      sheetName: sheet.name,
      range: { start: { row, col }, end: { row, col } },
    };
  };

  const evalBuiltinCall = (
    name: string,
    args: VbaValue[],
    lineNum: number
  ): VbaValue => {
    const fn = name.toLowerCase();
    switch (fn) {
      case 'range':
        return buildRangeFromArgs(getActiveSheet().name, args, lineNum);
      case 'cells':
        return buildCellsFromArgs(getActiveSheet(), args, lineNum);
      case 'sheets':
      case 'worksheets': {
        const sh = getSheetByNameOrIndex(coercePrimitive(args[0]), lineNum);
        return { kind: 'sheet', sheetName: sh.name };
      }
      case 'rows': {
        if (args.length === 0) {
          return {
            kind: 'range',
            sheetName: getActiveSheet().name,
            range: { start: { row: 1, col: 1 }, end: { row: getActiveSheet().rows, col: 1 } },
          };
        }
        const r = Math.max(1, Math.trunc(toNumber(args[0])));
        const sh = getActiveSheet();
        return {
          kind: 'range',
          sheetName: sh.name,
          range: { start: { row: r, col: 1 }, end: { row: r, col: sh.cols } },
        };
      }
      case 'columns': {
        const c = args.length > 0 ? resolveColArg(coercePrimitive(args[0]), lineNum) : 1;
        const sh = getActiveSheet();
        return {
          kind: 'range',
          sheetName: sh.name,
          range: { start: { row: 1, col: c }, end: { row: sh.rows, col: c } },
        };
      }
      case 'rgb': {
        const r = Math.max(0, Math.min(255, Math.trunc(toNumber(args[0] ?? 0))));
        const g = Math.max(0, Math.min(255, Math.trunc(toNumber(args[1] ?? 0))));
        const b = Math.max(0, Math.min(255, Math.trunc(toNumber(args[2] ?? 0))));
        return r + g * 256 + b * 65536;
      }
      case 'msgbox': {
        const msg = toStr(args[0] ?? '');
        logs.push({
          id: `msg-${logs.length}-${Date.now()}`,
          type: 'msgbox',
          line: lineNum,
          message: msg,
        });
        return 1;
      }
      case 'len':
        return toStr(args[0]).length;
      case 'left':
        return toStr(args[0]).slice(0, Math.max(0, Math.trunc(toNumber(args[1] ?? 0))));
      case 'right': {
        const s = toStr(args[0]);
        const n = Math.max(0, Math.trunc(toNumber(args[1] ?? 0)));
        return n === 0 ? '' : s.slice(-n);
      }
      case 'mid': {
        const s = toStr(args[0]);
        const start = Math.max(1, Math.trunc(toNumber(args[1] ?? 1))) - 1;
        const len = args[2] !== undefined ? Math.max(0, Math.trunc(toNumber(args[2]))) : undefined;
        return len !== undefined ? s.slice(start, start + len) : s.slice(start);
      }
      case 'trim':
        return toStr(args[0]).trim();
      case 'ltrim':
        return toStr(args[0]).replace(/^\s+/, '');
      case 'rtrim':
        return toStr(args[0]).replace(/\s+$/, '');
      case 'ucase':
        return toStr(args[0]).toUpperCase();
      case 'lcase':
        return toStr(args[0]).toLowerCase();
      case 'replace': {
        const s = toStr(args[0]);
        const find = toStr(args[1]);
        const rep = toStr(args[2]);
        return find ? s.split(find).join(rep) : s;
      }
      case 'instr': {
        const hasStart = args.length >= 3;
        const start = hasStart ? Math.max(1, Math.trunc(toNumber(args[0]))) : 1;
        const haystack = toStr(hasStart ? args[1] : args[0]);
        const needle = toStr(hasStart ? args[2] : args[1]);
        const idx = haystack.indexOf(needle, start - 1);
        return idx === -1 ? 0 : idx + 1;
      }
      case 'split':
        return toStr(args[0]).split(args[1] !== undefined ? toStr(args[1]) : ' ');
      case 'join':
        return Array.isArray(args[0])
          ? args[0].map((v) => toStr(v)).join(args[1] !== undefined ? toStr(args[1]) : ' ')
          : toStr(args[0]);
      case 'cstr':
        return toStr(args[0]);
      case 'cint':
      case 'clng':
        return Math.round(toNumber(args[0]));
      case 'cdbl':
      case 'val':
        return toNumber(args[0]);
      case 'int':
        return Math.floor(toNumber(args[0]));
      case 'round': {
        const n = toNumber(args[0]);
        const d = Math.trunc(toNumber(args[1] ?? 0));
        const f = Math.pow(10, d);
        return Math.round(n * f) / f;
      }
      case 'abs':
        return Math.abs(toNumber(args[0]));
      case 'sqr':
        return Math.sqrt(Math.max(0, toNumber(args[0])));
      case 'isnumeric': {
        const p = coercePrimitive(args[0]);
        if (typeof p === 'number') return Number.isFinite(p);
        if (typeof p === 'string') return p.trim() !== '' && Number.isFinite(Number(p));
        return false;
      }
      case 'isempty': {
        const p = coercePrimitive(args[0]);
        return p === null || p === undefined || p === '';
      }
      case 'iif':
        return toBool(args[0]) ? args[1] : args[2];
      case 'format': {
        const val = coercePrimitive(args[0]);
        const fmt = toStr(args[1]);
        if (typeof val === 'number' && fmt.includes('%')) {
          return `${(val * 100).toFixed(1)}%`;
        }
        if (typeof val === 'number' && fmt.includes('$')) {
          return `$${val.toLocaleString('en-US')}`;
        }
        return toStr(val);
      }
      case 'now':
        return new Date().toLocaleString();
      case 'date':
        return new Date().toLocaleDateString();
      case 'year':
        return new Date().getFullYear();
      case 'month':
        return new Date().getMonth() + 1;
      case 'day':
        return new Date().getDate();
      default:
        return undefined;
    }
  };

  const evalWorksheetFunction = (fnName: string, args: VbaValue[], lineNum: number): number => {
    const nums: number[] = [];
    for (const a of args) {
      if (a && typeof a === 'object' && !Array.isArray(a) && a.kind === 'range') {
        const coords = iterateRangeCoords(a.range);
        for (const c of coords) {
          const cell = readCell(a.sheetName, c.row, c.col);
          if (cell && cell.value !== null && cell.value !== '' && typeof cell.value !== 'boolean') {
            const n = Number(cell.value);
            if (Number.isFinite(n)) nums.push(n);
          }
        }
      } else {
        const n = Number(coercePrimitive(a));
        if (Number.isFinite(n)) nums.push(n);
      }
    }

    switch (fnName.toLowerCase()) {
      case 'sum':
        return nums.reduce((acc, n) => acc + n, 0);
      case 'average':
        return nums.length ? nums.reduce((acc, n) => acc + n, 0) / nums.length : 0;
      case 'count':
        return nums.length;
      case 'min':
        return nums.length ? Math.min(...nums) : 0;
      case 'max':
        return nums.length ? Math.max(...nums) : 0;
      case 'round': {
        const n = toNumber(args[0]);
        const d = Math.trunc(toNumber(args[1] ?? 0));
        const f = Math.pow(10, d);
        return Math.round(n * f) / f;
      }
      default:
        throw new VbaRuntimeError(`Unsupported WorksheetFunction.${fnName}.`, lineNum);
    }
  };

  class ExprEvaluator {
    private pos = 0;

    constructor(
      private tokens: VbaToken[],
      private lineNum: number,
      private env: Map<string, VbaValue>,
      private withStack: VbaValue[]
    ) {}

    peek(offset = 0): VbaToken | undefined {
      return this.tokens[this.pos + offset];
    }

    consume(): VbaToken | undefined {
      return this.tokens[this.pos++];
    }

    parseExpression(): VbaValue {
      return this.parseOr();
    }

    private parseOr(): VbaValue {
      let left = this.parseAnd();
      while (this.peek()?.value.toLowerCase() === 'or') {
        this.consume();
        const right = this.parseAnd();
        left = toBool(left) || toBool(right);
      }
      return left;
    }

    private parseAnd(): VbaValue {
      let left = this.parseNot();
      while (this.peek()?.value.toLowerCase() === 'and') {
        this.consume();
        const right = this.parseNot();
        left = toBool(left) && toBool(right);
      }
      return left;
    }

    private parseNot(): VbaValue {
      if (this.peek()?.value.toLowerCase() === 'not') {
        this.consume();
        return !toBool(this.parseNot());
      }
      return this.parseComparison();
    }

    private parseComparison(): VbaValue {
      let left = this.parseConcat();
      while (true) {
        const tok = this.peek();
        if (tok && ['=', '<>', '<', '<=', '>', '>='].includes(tok.value)) {
          this.consume();
          const right = this.parseConcat();
          left = compareVba(coercePrimitive(left), coercePrimitive(right), tok.value);
        } else {
          break;
        }
      }
      return left;
    }

    private parseConcat(): VbaValue {
      let left = this.parseAddSub();
      while (this.peek()?.value === '&') {
        this.consume();
        const right = this.parseAddSub();
        left = `${toStr(left)}${toStr(right)}`;
      }
      return left;
    }

    private parseAddSub(): VbaValue {
      let left = this.parseMod();
      while (this.peek()?.value === '+' || this.peek()?.value === '-') {
        const op = this.consume()!.value;
        const right = this.parseMod();
        const lPrim = coercePrimitive(left);
        const rPrim = coercePrimitive(right);
        if (op === '+' && (typeof lPrim === 'string' || typeof rPrim === 'string')) {
          const lNum = Number(lPrim);
          const rNum = Number(rPrim);
          if (Number.isFinite(lNum) && Number.isFinite(rNum) && lPrim !== '' && rPrim !== '') {
            left = lNum + rNum;
          } else {
            left = `${toStr(lPrim)}${toStr(rPrim)}`;
          }
        } else {
          left = op === '+' ? toNumber(lPrim) + toNumber(rPrim) : toNumber(lPrim) - toNumber(rPrim);
        }
      }
      return left;
    }

    private parseMod(): VbaValue {
      let left = this.parseMulDiv();
      while (this.peek()?.value.toLowerCase() === 'mod') {
        this.consume();
        const right = this.parseMulDiv();
        const d = toNumber(right);
        if (d === 0) throw new VbaRuntimeError('Division by zero.', this.lineNum);
        left = toNumber(left) % d;
      }
      return left;
    }

    private parseMulDiv(): VbaValue {
      let left = this.parsePower();
      while (this.peek()?.value === '*' || this.peek()?.value === '/' || this.peek()?.value === '\\') {
        const op = this.consume()!.value;
        const right = this.parsePower();
        const l = toNumber(left);
        const r = toNumber(right);
        if ((op === '/' || op === '\\') && r === 0) {
          throw new VbaRuntimeError('Division by zero.', this.lineNum);
        }
        if (op === '*') left = l * r;
        else if (op === '/') left = l / r;
        else left = Math.trunc(l / r);
      }
      return left;
    }

    private parsePower(): VbaValue {
      let left = this.parseUnary();
      while (this.peek()?.value === '^') {
        this.consume();
        const right = this.parseUnary();
        left = Math.pow(toNumber(left), toNumber(right));
      }
      return left;
    }

    private parseUnary(): VbaValue {
      const tok = this.peek();
      if (tok?.value === '-' || tok?.value === '+') {
        this.consume();
        const val = toNumber(this.parseUnary());
        return tok.value === '-' ? -val : val;
      }
      return this.parsePostfix();
    }

    private parsePostfix(): VbaValue {
      let val = this.parsePrimary();
      while (this.peek()?.value === '.') {
        this.consume();
        const memberTok = this.consume();
        if (!memberTok) {
          throw new VbaRuntimeError('Expected property or method name after ".".', this.lineNum);
        }
        let args: VbaValue[] | null = null;
        if (this.peek()?.value === '(') {
          args = this.parseArgList();
        }
        // Handle Application.WorksheetFunction or WorksheetFunction.<Fn>
        if (val === '__APPLICATION__' && memberTok.value.toLowerCase() === 'worksheetfunction') {
          val = '__WORKSHEETFUNCTION__';
          continue;
        }
        if (val === '__WORKSHEETFUNCTION__') {
          val = evalWorksheetFunction(memberTok.value, args ?? [], this.lineNum);
          continue;
        }
        val = getProperty(val, memberTok.value, args, this.lineNum);
      }
      return val;
    }

    private parseArgList(): VbaValue[] {
      this.consume(); // '('
      const args: VbaValue[] = [];
      if (this.peek()?.value !== ')') {
        while (this.pos < this.tokens.length) {
          args.push(this.parseExpression());
          if (this.peek()?.value === ',') {
            this.consume();
          } else {
            break;
          }
        }
      }
      if (this.peek()?.value === ')') {
        this.consume();
      }
      return args;
    }

    private parsePrimary(): VbaValue {
      const tok = this.peek();
      if (!tok) return null;

      if (tok.type === 'number') {
        this.consume();
        return parseFloat(tok.value);
      }
      if (tok.type === 'string') {
        this.consume();
        return tok.value.slice(1, -1).replace(/""/g, '"');
      }
      if (tok.value === '(') {
        this.consume();
        const val = this.parseExpression();
        if (this.peek()?.value === ')') this.consume();
        return val;
      }
      if (tok.value === '.') {
        if (this.withStack.length === 0) {
          throw new VbaRuntimeError('Invalid "." reference outside of a With block.', this.lineNum);
        }
        return this.withStack[this.withStack.length - 1];
      }

      this.consume();
      const lower = tok.value.toLowerCase();

      // Constants
      if (lower === 'true') return true;
      if (lower === 'false') return false;
      if (lower === 'nothing' || lower === 'empty' || lower === 'null') return null;
      if (lower === 'vbcrlf' || lower === 'vbnewline') return '\n';
      if (lower === 'vbtab') return '\t';
      if (
        [
          'vbred',
          'vbgreen',
          'vbblue',
          'vbyellow',
          'vbwhite',
          'vbblack',
          'vbcyan',
          'vbmagenta',
          'xlleft',
          'xlcenter',
          'xlright',
          'xlup',
          'xldown',
          'xltoleft',
          'xltoright',
          'xlnone',
          'xlcontinuous',
        ].includes(lower)
      ) {
        return lower;
      }

      if (lower === 'activesheet') {
        return { kind: 'sheet', sheetName: getActiveSheet().name };
      }
      if (lower === 'thisworkbook' || lower === 'activeworkbook') {
        return { kind: 'sheet', sheetName: getActiveSheet().name };
      }
      if (lower === 'activecell' || lower === 'selection') {
        return {
          kind: 'range',
          sheetName: getActiveSheet().name,
          range: normalizeRange(selectedRange),
        };
      }
      if (lower === 'application') return '__APPLICATION__';
      if (lower === 'worksheetfunction') return '__WORKSHEETFUNCTION__';

      // Call or array indexing
      if (this.peek()?.value === '(') {
        const args = this.parseArgList();
        const userProc = procedures.get(lower);
        if (userProc) {
          return runProcedure(userProc, args);
        }
        const envVal = this.env.get(lower);
        if (Array.isArray(envVal)) {
          const idx = Math.trunc(toNumber(args[0] ?? 0));
          return envVal[idx] ?? null;
        }
        const builtinRes = evalBuiltinCall(lower, args, this.lineNum);
        if (builtinRes !== undefined) return builtinRes;
        throw new VbaRuntimeError(`Sub or Function "${tok.value}" is not defined.`, this.lineNum);
      }

      if (lower === 'sheets' || lower === 'worksheets') {
        return '__SHEETS_COLLECTION__';
      }

      if (lower === 'rows' || lower === 'columns') {
        return evalBuiltinCall(lower, [], this.lineNum);
      }

      if (this.env.has(lower)) {
        return this.env.get(lower);
      }

      return null;
    }
  }

  const evalExpr = (
    tokens: VbaToken[],
    lineNum: number,
    env: Map<string, VbaValue>,
    withStack: VbaValue[]
  ): VbaValue => {
    if (tokens.length === 0) return null;
    const evaluator = new ExprEvaluator(tokens, lineNum, env, withStack);
    return evaluator.parseExpression();
  };

  // Resolve assignment LHS: either variable name or target object + property path
  const executeAssignment = (
    lhsTokens: VbaToken[],
    rhsTokens: VbaToken[],
    lineNum: number,
    env: Map<string, VbaValue>,
    withStack: VbaValue[]
  ) => {
    const val = evalExpr(rhsTokens, lineNum, env, withStack);

    // Simple variable: x = expr
    if (lhsTokens.length === 1 && lhsTokens[0].value !== '.') {
      const varName = lhsTokens[0].value.toLowerCase();
      env.set(varName, val);
      return;
    }

    // Split LHS at property chain after last expression or inside With
    // Find dot chains at top level (parenDepth === 0)
    const dotPositions: number[] = [];
    let depth = 0;
    for (let i = 0; i < lhsTokens.length; i++) {
      if (lhsTokens[i].value === '(') depth++;
      else if (lhsTokens[i].value === ')') depth--;
      else if (lhsTokens[i].value === '.' && depth === 0) {
        dotPositions.push(i);
      }
    }

    if (dotPositions.length === 0) {
      throw new VbaRuntimeError('Invalid assignment target.', lineNum);
    }

    // If starts with '.', base is top of withStack
    if (dotPositions[0] === 0) {
      if (withStack.length === 0) {
        throw new VbaRuntimeError('Invalid "." assignment outside With block.', lineNum);
      }
      let base: VbaValue = withStack[withStack.length - 1];
      // Walk intermediate method/range calls like .Cells(1,1).Value or .Interior.Color
      const segments = splitByDotsAtTopLevel(lhsTokens.slice(1));
      let idx = 0;
      while (idx < segments.length - 1) {
        const seg = segments[idx];
        const segName = seg[0]?.value.toLowerCase() ?? '';
        if (['interior', 'font', 'borders'].includes(segName)) {
          break;
        }
        const args =
          seg.length > 1 && seg[1].value === '('
            ? evalArgTokens(seg.slice(2, -1), lineNum, env, withStack)
            : null;
        base = getProperty(base, seg[0].value, args, lineNum);
        idx++;
      }
      const propPath = segments.slice(idx).map((s) => s[0].value);
      setProperty(base, propPath, val, lineNum);
      return;
    }

    // Otherwise evaluate base before first property that isn't Sheets/Range/Cells/Offset/etc.
    const segments = splitByDotsAtTopLevel(lhsTokens);
    let base: VbaValue = evalExpr(segments[0], lineNum, env, withStack);
    let idx = 1;
    while (idx < segments.length - 1) {
      const seg = segments[idx];
      const segName = seg[0]?.value.toLowerCase() ?? '';
      if (['interior', 'font', 'borders'].includes(segName)) {
        break;
      }
      const args =
        seg.length > 1 && seg[1].value === '('
          ? evalArgTokens(seg.slice(2, -1), lineNum, env, withStack)
          : null;
      base = getProperty(base, seg[0].value, args, lineNum);
      idx++;
    }
    const propPath = segments.slice(idx).map((s) => s[0].value);
    setProperty(base, propPath, val, lineNum);
  };

  const evalArgTokens = (
    tokens: VbaToken[],
    lineNum: number,
    env: Map<string, VbaValue>,
    withStack: VbaValue[]
  ): VbaValue[] => {
    const args: VbaValue[] = [];
    let start = 0;
    let depth = 0;
    for (let i = 0; i <= tokens.length; i++) {
      if (i === tokens.length || (tokens[i].value === ',' && depth === 0)) {
        const slice = tokens.slice(start, i);
        if (slice.length > 0) {
          args.push(evalExpr(slice, lineNum, env, withStack));
        }
        start = i + 1;
      } else if (tokens[i].value === '(') depth++;
      else if (tokens[i].value === ')') depth--;
    }
    return args;
  };

  const runProcedure = (proc: ProcedureDef, args: VbaValue[]): VbaValue => {
    const env = new Map<string, VbaValue>();
    proc.params.forEach((p, idx) => {
      env.set(p, args[idx] ?? null);
    });
    if (proc.kind === 'function') {
      env.set(proc.name, null);
    }

    const withStack: VbaValue[] = [];
    try {
      executeBlock(proc.startLine, proc.endLine, env, withStack);
    } catch (sig) {
      if (sig instanceof ExitSignal && (sig.kind === 'sub' || sig.kind === 'function')) {
        // Normal exit
      } else {
        throw sig;
      }
    }

    return proc.kind === 'function' ? env.get(proc.name) : null;
  };

  const executeBlock = (
    startIdx: number,
    endIdx: number,
    env: Map<string, VbaValue>,
    withStack: VbaValue[]
  ) => {
    let i = startIdx;
    while (i <= endIdx) {
      stepCount++;
      if (stepCount > MAX_STEPS) {
        throw new VbaRuntimeError(
          `Execution halted after ${MAX_STEPS.toLocaleString()} steps to prevent an infinite loop.`,
          parsedLines[i].lineNum
        );
      }

      const { lineNum, tokens } = parsedLines[i];
      if (tokens.length === 0) {
        i++;
        continue;
      }

      const w0 = tokens[0].value.toLowerCase();
      const w1 = tokens[1]?.value.toLowerCase() ?? '';

      // Dim / Const / Option
      if (w0 === 'option') {
        i++;
        continue;
      }
      if (w0 === 'dim' || w0 === 'const') {
        // Initialize variables
        const varTok = tokens[1];
        if (varTok) {
          const eqIdx = findTopLevelToken(tokens, '=');
          if (eqIdx !== -1) {
            env.set(
              varTok.value.toLowerCase(),
              evalExpr(tokens.slice(eqIdx + 1), lineNum, env, withStack)
            );
          } else if (!env.has(varTok.value.toLowerCase())) {
            env.set(varTok.value.toLowerCase(), null);
          }
        }
        i++;
        continue;
      }

      // Exit statements
      if (w0 === 'exit') {
        if (w1 === 'for') throw new ExitSignal('for');
        if (w1 === 'do') throw new ExitSignal('do');
        if (w1 === 'sub') throw new ExitSignal('sub');
        if (w1 === 'function') throw new ExitSignal('function');
      }

      // With ... End With
      if (w0 === 'with') {
        const endWithIdx = findMatchingEnd(parsedLines, i, endIdx, 'with', ['end', 'with']);
        const targetObj = evalExpr(tokens.slice(1), lineNum, env, withStack);
        withStack.push(targetObj);
        try {
          executeBlock(i + 1, endWithIdx - 1, env, withStack);
        } finally {
          withStack.pop();
        }
        i = endWithIdx + 1;
        continue;
      }

      // For Each ... In ... Next
      if (w0 === 'for' && w1 === 'each') {
        const nextIdx = findMatchingForNext(parsedLines, i, endIdx);
        const inIdx = tokens.findIndex((t) => t.value.toLowerCase() === 'in');
        const varName = tokens[2]?.value.toLowerCase();
        const collectionVal = evalExpr(tokens.slice(inIdx + 1), lineNum, env, withStack);

        let items: VbaValue[] = [];
        if (
          collectionVal &&
          typeof collectionVal === 'object' &&
          !Array.isArray(collectionVal) &&
          collectionVal.kind === 'range'
        ) {
          items = iterateRangeCoords(collectionVal.range).map((c) => ({
            kind: 'range',
            sheetName: collectionVal.sheetName,
            range: { start: c, end: c },
          }));
        } else if (Array.isArray(collectionVal)) {
          items = collectionVal;
        }

        for (const item of items) {
          env.set(varName, item);
          try {
            executeBlock(i + 1, nextIdx - 1, env, withStack);
          } catch (sig) {
            if (sig instanceof ExitSignal && sig.kind === 'for') break;
            throw sig;
          }
        }
        i = nextIdx + 1;
        continue;
      }

      // For i = start To end [Step step] ... Next
      if (w0 === 'for') {
        const nextIdx = findMatchingForNext(parsedLines, i, endIdx);
        const eqIdx = tokens.findIndex((t) => t.value === '=');
        const toIdx = tokens.findIndex((t) => t.value.toLowerCase() === 'to');
        const stepIdx = tokens.findIndex((t) => t.value.toLowerCase() === 'step');
        const varName = tokens[1].value.toLowerCase();

        const startVal = toNumber(evalExpr(tokens.slice(eqIdx + 1, toIdx), lineNum, env, withStack));
        const endVal = toNumber(
          evalExpr(tokens.slice(toIdx + 1, stepIdx !== -1 ? stepIdx : tokens.length), lineNum, env, withStack)
        );
        const stepVal =
          stepIdx !== -1
            ? toNumber(evalExpr(tokens.slice(stepIdx + 1), lineNum, env, withStack))
            : 1;

        if (stepVal === 0) {
          throw new VbaRuntimeError('For loop Step cannot be 0.', lineNum);
        }

        for (
          let counter = startVal;
          stepVal > 0 ? counter <= endVal : counter >= endVal;
          counter += stepVal
        ) {
          env.set(varName, counter);
          try {
            executeBlock(i + 1, nextIdx - 1, env, withStack);
          } catch (sig) {
            if (sig instanceof ExitSignal && sig.kind === 'for') break;
            throw sig;
          }
        }
        i = nextIdx + 1;
        continue;
      }

      // Do [While|Until cond] ... Loop [While|Until cond]
      if (w0 === 'do') {
        const loopIdx = findMatchingDoLoop(parsedLines, i, endIdx);
        const loopTokens = parsedLines[loopIdx].tokens;
        const hasPreCond = w1 === 'while' || w1 === 'until';
        const l1 = loopTokens[1]?.value.toLowerCase() ?? '';
        const hasPostCond = l1 === 'while' || l1 === 'until';

        while (true) {
          if (hasPreCond) {
            const cond = toBool(evalExpr(tokens.slice(2), lineNum, env, withStack));
            if (w1 === 'while' && !cond) break;
            if (w1 === 'until' && cond) break;
          }
          try {
            executeBlock(i + 1, loopIdx - 1, env, withStack);
          } catch (sig) {
            if (sig instanceof ExitSignal && sig.kind === 'do') break;
            throw sig;
          }
          if (hasPostCond) {
            const cond = toBool(evalExpr(loopTokens.slice(2), parsedLines[loopIdx].lineNum, env, withStack));
            if (l1 === 'while' && !cond) break;
            if (l1 === 'until' && cond) break;
          }
        }
        i = loopIdx + 1;
        continue;
      }

      // If ... Then
      if (w0 === 'if') {
        const thenIdx = tokens.findIndex((t) => t.value.toLowerCase() === 'then');
        // Single-line If
        if (thenIdx !== -1 && thenIdx < tokens.length - 1) {
          const cond = toBool(evalExpr(tokens.slice(1, thenIdx), lineNum, env, withStack));
          const elseIdx = findTopLevelKeyword(tokens, 'else', thenIdx + 1);
          if (cond) {
            const stmtTokens = elseIdx !== -1 ? tokens.slice(thenIdx + 1, elseIdx) : tokens.slice(thenIdx + 1);
            executeSingleStatement(stmtTokens, lineNum, env, withStack);
          } else if (elseIdx !== -1) {
            executeSingleStatement(tokens.slice(elseIdx + 1), lineNum, env, withStack);
          }
          i++;
          continue;
        }

        // Multi-line If
        const branches = parseIfBranches(parsedLines, i, endIdx);
        let executed = false;
        for (const br of branches.cases) {
          const cond =
            br.condTokens === null
              ? true
              : toBool(evalExpr(br.condTokens, parsedLines[br.headerLineIdx].lineNum, env, withStack));
          if (cond && !executed) {
            executeBlock(br.bodyStart, br.bodyEnd, env, withStack);
            executed = true;
            break;
          }
        }
        i = branches.endIfIdx + 1;
        continue;
      }

      // Select Case ... End Select
      if (w0 === 'select' && w1 === 'case') {
        const endSelectIdx = findMatchingEnd(parsedLines, i, endIdx, 'select', ['end', 'select']);
        const testVal = coercePrimitive(evalExpr(tokens.slice(2), lineNum, env, withStack));
        const cases = parseSelectCases(parsedLines, i, endSelectIdx);
        for (const c of cases) {
          if (c.isElse) {
            executeBlock(c.bodyStart, c.bodyEnd, env, withStack);
            break;
          }
          const matched = c.matchExprs.some((exprToks) => {
            const val = coercePrimitive(evalExpr(exprToks, parsedLines[c.headerIdx].lineNum, env, withStack));
            return compareVba(testVal, val, '=');
          });
          if (matched) {
            executeBlock(c.bodyStart, c.bodyEnd, env, withStack);
            break;
          }
        }
        i = endSelectIdx + 1;
        continue;
      }

      // Normal statement(s), optionally split by ':'
      executeSingleStatement(tokens, lineNum, env, withStack);
      i++;
    }
  };

  const executeSingleStatement = (
    tokens: VbaToken[],
    lineNum: number,
    env: Map<string, VbaValue>,
    withStack: VbaValue[]
  ) => {
    if (tokens.length === 0) return;

    // Split on top-level ':'
    const colIdx = findTopLevelToken(tokens, ':');
    if (colIdx !== -1) {
      executeSingleStatement(tokens.slice(0, colIdx), lineNum, env, withStack);
      executeSingleStatement(tokens.slice(colIdx + 1), lineNum, env, withStack);
      return;
    }

    const w0 = tokens[0].value.toLowerCase();

    // Debug.Print ...
    if (w0 === 'debug' && tokens[1]?.value === '.' && tokens[2]?.value.toLowerCase() === 'print') {
      const val = evalExpr(tokens.slice(3), lineNum, env, withStack);
      logs.push({
        id: `print-${logs.length}-${Date.now()}`,
        type: 'print',
        line: lineNum,
        message: toStr(val),
      });
      return;
    }

    // MsgBox "text" (without parens)
    if (w0 === 'msgbox' && tokens[1]?.value !== '(') {
      const commaIdx = findTopLevelToken(tokens, ',');
      const exprSlice = commaIdx !== -1 ? tokens.slice(1, commaIdx) : tokens.slice(1);
      const val = evalExpr(exprSlice, lineNum, env, withStack);
      logs.push({
        id: `msg-${logs.length}-${Date.now()}`,
        type: 'msgbox',
        line: lineNum,
        message: toStr(val),
      });
      return;
    }

    // Set x = expr / Let x = expr
    const cleanTokens = w0 === 'set' || w0 === 'let' ? tokens.slice(1) : tokens;

    // Assignment: LHS = RHS (top-level '=')
    const eqIdx = findTopLevelToken(cleanTokens, '=');
    if (eqIdx !== -1) {
      executeAssignment(
        cleanTokens.slice(0, eqIdx),
        cleanTokens.slice(eqIdx + 1),
        lineNum,
        env,
        withStack
      );
      return;
    }

    // Call SubName arg1, arg2
    const stmtToks = w0 === 'call' ? tokens.slice(1) : tokens;
    const procName = stmtToks[0]?.value.toLowerCase();
    const userProc = procedures.get(procName);
    if (userProc) {
      const argSlice =
        stmtToks[1]?.value === '(' ? stmtToks.slice(2, -1) : stmtToks.slice(1);
      const args = evalArgTokens(argSlice, lineNum, env, withStack);
      runProcedure(userProc, args);
      return;
    }

    // Expression call (e.g. Range("A1:C1").ClearContents or .Select)
    evalExpr(stmtToks, lineNum, env, withStack);
  };

  try {
    runProcedure(firstSub, []);
    sheets = recalculateSheets(sheets);

    const durationMs = Math.max(1, Math.round(performance.now() - startTime));
    return {
      sheets,
      activeSheetId,
      selectedRange,
      result: {
        status: 'success',
        headline: 'Macro ran successfully',
        detail:
          changedCells.size > 0
            ? `Your spreadsheet is up to date (${changedCells.size} cell${changedCells.size === 1 ? '' : 's'} updated).`
            : 'Your spreadsheet is up to date.',
        durationMs,
        logs,
        changedCellsCount: changedCells.size,
      },
    };
  } catch (err) {
    const runtimeErr =
      err instanceof VbaRuntimeError
        ? err
        : new VbaRuntimeError(err instanceof Error ? err.message : 'Unexpected runtime error.', 1);

    return makeErrorResult(input, startTime, runtimeErr.line, runtimeErr.message, logs);
  }
}

function makeErrorResult(
  input: RunMacroInput,
  startTime: number,
  line: number,
  message: string,
  priorLogs: ExecutionLogEntry[] = []
): RunMacroOutput {
  return {
    sheets: input.sheets,
    activeSheetId: input.activeSheetId,
    selectedRange: input.selectedRange,
    result: {
      status: 'error',
      headline: `Error on line ${line}`,
      detail: message,
      durationMs: Math.max(1, Math.round(performance.now() - startTime)),
      logs: [
        ...priorLogs,
        {
          id: `err-${Date.now()}`,
          type: 'error',
          line,
          message: `Line ${line}: ${message}`,
        },
      ],
      changedCellsCount: 0,
      errorLine: line,
    },
  };
}

function findTopLevelToken(tokens: VbaToken[], target: string): number {
  let depth = 0;
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i].value === '(') depth++;
    else if (tokens[i].value === ')') depth--;
    else if (depth === 0 && tokens[i].value === target) return i;
  }
  return -1;
}

function findTopLevelKeyword(tokens: VbaToken[], kw: string, fromIdx = 0): number {
  let depth = 0;
  for (let i = fromIdx; i < tokens.length; i++) {
    if (tokens[i].value === '(') depth++;
    else if (tokens[i].value === ')') depth--;
    else if (depth === 0 && tokens[i].value.toLowerCase() === kw) return i;
  }
  return -1;
}

function splitByDotsAtTopLevel(tokens: VbaToken[]): VbaToken[][] {
  const parts: VbaToken[][] = [];
  let start = 0;
  let depth = 0;
  for (let i = 0; i <= tokens.length; i++) {
    if (i === tokens.length || (tokens[i].value === '.' && depth === 0)) {
      parts.push(tokens.slice(start, i));
      start = i + 1;
    } else if (tokens[i].value === '(') depth++;
    else if (tokens[i].value === ')') depth--;
  }
  return parts;
}

function findMatchingEnd(
  lines: ExecLine[],
  startIdx: number,
  maxIdx: number,
  openWord: string,
  closePair: [string, string]
): number {
  let depth = 1;
  for (let i = startIdx + 1; i <= maxIdx; i++) {
    const toks = lines[i].tokens;
    if (toks.length === 0) continue;
    const w0 = toks[0].value.toLowerCase();
    const w1 = toks[1]?.value.toLowerCase() ?? '';
    if (w0 === openWord) depth++;
    else if (w0 === closePair[0] && w1 === closePair[1]) {
      depth--;
      if (depth === 0) return i;
    }
  }
  throw new VbaRuntimeError(
    `Missing "${closePair[0]} ${closePair[1]}" for block on line ${lines[startIdx].lineNum}.`,
    lines[startIdx].lineNum
  );
}

function findMatchingForNext(lines: ExecLine[], startIdx: number, maxIdx: number): number {
  let depth = 1;
  for (let i = startIdx + 1; i <= maxIdx; i++) {
    const toks = lines[i].tokens;
    if (toks.length === 0) continue;
    const w0 = toks[0].value.toLowerCase();
    if (w0 === 'for') depth++;
    else if (w0 === 'next') {
      depth--;
      if (depth === 0) return i;
    }
  }
  throw new VbaRuntimeError(`Missing "Next" for For loop on line ${lines[startIdx].lineNum}.`, lines[startIdx].lineNum);
}

function findMatchingDoLoop(lines: ExecLine[], startIdx: number, maxIdx: number): number {
  let depth = 1;
  for (let i = startIdx + 1; i <= maxIdx; i++) {
    const toks = lines[i].tokens;
    if (toks.length === 0) continue;
    const w0 = toks[0].value.toLowerCase();
    if (w0 === 'do') depth++;
    else if (w0 === 'loop') {
      depth--;
      if (depth === 0) return i;
    }
  }
  throw new VbaRuntimeError(`Missing "Loop" for Do loop on line ${lines[startIdx].lineNum}.`, lines[startIdx].lineNum);
}

function parseIfBranches(lines: ExecLine[], startIdx: number, maxIdx: number) {
  let depth = 1;
  let curHeader = startIdx;
  let curCond: VbaToken[] | null = lines[startIdx].tokens.slice(
    1,
    lines[startIdx].tokens.findIndex((t) => t.value.toLowerCase() === 'then')
  );
  const cases: Array<{
    headerLineIdx: number;
    condTokens: VbaToken[] | null;
    bodyStart: number;
    bodyEnd: number;
  }> = [];

  for (let i = startIdx + 1; i <= maxIdx; i++) {
    const toks = lines[i].tokens;
    if (toks.length === 0) continue;
    const w0 = toks[0].value.toLowerCase();
    const w1 = toks[1]?.value.toLowerCase() ?? '';

    if (w0 === 'if') {
      const thenIdx = toks.findIndex((t) => t.value.toLowerCase() === 'then');
      if (thenIdx === toks.length - 1) depth++;
    } else if ((w0 === 'end' && w1 === 'if') || w0 === 'endif') {
      depth--;
      if (depth === 0) {
        cases.push({
          headerLineIdx: curHeader,
          condTokens: curCond,
          bodyStart: curHeader + 1,
          bodyEnd: i - 1,
        });
        return { cases, endIfIdx: i };
      }
    } else if (depth === 1 && w0 === 'elseif') {
      cases.push({
        headerLineIdx: curHeader,
        condTokens: curCond,
        bodyStart: curHeader + 1,
        bodyEnd: i - 1,
      });
      curHeader = i;
      const thenIdx = toks.findIndex((t) => t.value.toLowerCase() === 'then');
      curCond = toks.slice(1, thenIdx);
    } else if (depth === 1 && w0 === 'else') {
      cases.push({
        headerLineIdx: curHeader,
        condTokens: curCond,
        bodyStart: curHeader + 1,
        bodyEnd: i - 1,
      });
      curHeader = i;
      curCond = null;
    }
  }

  throw new VbaRuntimeError(`Missing "End If" for If block on line ${lines[startIdx].lineNum}.`, lines[startIdx].lineNum);
}

function parseSelectCases(lines: ExecLine[], startIdx: number, endSelectIdx: number) {
  const cases: Array<{
    headerIdx: number;
    isElse: boolean;
    matchExprs: VbaToken[][];
    bodyStart: number;
    bodyEnd: number;
  }> = [];

  let depth = 1;
  let current: {
    headerIdx: number;
    isElse: boolean;
    matchExprs: VbaToken[][];
  } | null = null;

  for (let i = startIdx + 1; i < endSelectIdx; i++) {
    const toks = lines[i].tokens;
    if (toks.length === 0) continue;
    const w0 = toks[0].value.toLowerCase();
    const w1 = toks[1]?.value.toLowerCase() ?? '';

    if (w0 === 'select' && w1 === 'case') depth++;
    else if (w0 === 'end' && w1 === 'select') depth--;
    else if (depth === 1 && w0 === 'case') {
      if (current) {
        cases.push({
          ...current,
          bodyStart: current.headerIdx + 1,
          bodyEnd: i - 1,
        });
      }
      if (w1 === 'else') {
        current = { headerIdx: i, isElse: true, matchExprs: [] };
      } else {
        const exprParts: VbaToken[][] = [];
        let start = 1;
        let pDepth = 0;
        for (let k = 1; k <= toks.length; k++) {
          if (k === toks.length || (toks[k].value === ',' && pDepth === 0)) {
            exprParts.push(toks.slice(start, k));
            start = k + 1;
          } else if (toks[k].value === '(') pDepth++;
          else if (toks[k].value === ')') pDepth--;
        }
        current = { headerIdx: i, isElse: false, matchExprs: exprParts };
      }
    }
  }

  if (current) {
    cases.push({
      ...current,
      bodyStart: current.headerIdx + 1,
      bodyEnd: endSelectIdx - 1,
    });
  }

  return cases;
}

function compareVba(a: string | number | boolean | null, b: string | number | boolean | null, op: string): boolean {
  const numA = typeof a === 'number' ? a : Number(a);
  const numB = typeof b === 'number' ? b : Number(b);
  const bothNum =
    a !== null &&
    b !== null &&
    a !== '' &&
    b !== '' &&
    typeof a !== 'boolean' &&
    typeof b !== 'boolean' &&
    Number.isFinite(numA) &&
    Number.isFinite(numB);

  const l = bothNum ? numA : String(a ?? '').toLowerCase();
  const r = bothNum ? numB : String(b ?? '').toLowerCase();

  switch (op) {
    case '=':
      return l === r;
    case '<>':
      return l !== r;
    case '<':
      return l < r;
    case '<=':
      return l <= r;
    case '>':
      return l > r;
    case '>=':
      return l >= r;
    default:
      return false;
  }
}
