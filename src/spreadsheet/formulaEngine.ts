import { CellData, SheetData } from '../types/workbook';
import { iterateRangeCoords, parseCellAddress, parseRangeAddress } from './addressing';

export function recalculateSheets(sheets: SheetData[]): SheetData[] {
  // Clone cells shallowly per pass and run 2 passes for dependent formulas
  let current = sheets.map((s) => ({
    ...s,
    cells: Object.fromEntries(
      Object.entries(s.cells).map(([k, v]) => [k, { ...v }])
    ),
  }));

  for (let pass = 0; pass < 2; pass++) {
    for (const sheet of current) {
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        if (cell.formula && cell.formula.startsWith('=')) {
          const val = evaluateFormula(cell.formula, sheet, current, new Set([`${sheet.name}!${addr}`]));
          cell.value = val;
        }
      }
    }
  }

  return current;
}

export function evaluateFormula(
  formula: string,
  activeSheet: SheetData,
  allSheets: SheetData[],
  visited: Set<string> = new Set()
): string | number | boolean {
  const expr = formula.trim().replace(/^=/, '').trim();
  if (!expr) return '';

  try {
    const tokens = tokenizeFormula(expr);
    const parser = new FormulaParser(tokens, activeSheet, allSheets, visited);
    const result = parser.parseExpression();
    return result;
  } catch {
    return '#VALUE!';
  }
}

type FToken =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'bool'; value: boolean }
  | { type: 'ident'; value: string }
  | { type: 'op'; value: string }
  | { type: 'paren'; value: '(' | ')' }
  | { type: 'comma' }
  | { type: 'colon' }
  | { type: 'bang' };

function tokenizeFormula(expr: string): FToken[] {
  const tokens: FToken[] = [];
  let i = 0;

  while (i < expr.length) {
    const ch = expr[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let str = '';
      i++;
      while (i < expr.length) {
        if (expr[i] === '"' && expr[i + 1] === '"') {
          str += '"';
          i += 2;
        } else if (expr[i] === '"') {
          i++;
          break;
        } else {
          str += expr[i++];
        }
      }
      tokens.push({ type: 'string', value: str });
      continue;
    }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(expr[i + 1] ?? ''))) {
      let numStr = '';
      while (i < expr.length && /[0-9.]/.test(expr[i])) {
        numStr += expr[i++];
      }
      tokens.push({ type: 'number', value: parseFloat(numStr) });
      continue;
    }
    if (ch === '(' || ch === ')') {
      tokens.push({ type: 'paren', value: ch });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ type: 'comma' });
      i++;
      continue;
    }
    if (ch === ':') {
      tokens.push({ type: 'colon' });
      i++;
      continue;
    }
    if (ch === '!') {
      tokens.push({ type: 'bang' });
      i++;
      continue;
    }
    const twoChar = expr.slice(i, i + 2);
    if (['<=', '>=', '<>'].includes(twoChar)) {
      tokens.push({ type: 'op', value: twoChar });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>'.includes(ch)) {
      tokens.push({ type: 'op', value: ch });
      i++;
      continue;
    }
    if (/[A-Za-z_$]/.test(ch)) {
      let ident = '';
      while (i < expr.length && /[A-Za-z0-9_$]/.test(expr[i])) {
        ident += expr[i++];
      }
      const upper = ident.toUpperCase();
      if (upper === 'TRUE') tokens.push({ type: 'bool', value: true });
      else if (upper === 'FALSE') tokens.push({ type: 'bool', value: false });
      else tokens.push({ type: 'ident', value: ident });
      continue;
    }
    // Single quoted sheet name: 'Sheet 2'!A1
    if (ch === "'") {
      let name = '';
      i++;
      while (i < expr.length && expr[i] !== "'") {
        name += expr[i++];
      }
      i++; // skip closing '
      tokens.push({ type: 'ident', value: name });
      continue;
    }
    i++;
  }

  return tokens;
}

class FormulaParser {
  private pos = 0;

  constructor(
    private tokens: FToken[],
    private activeSheet: SheetData,
    private allSheets: SheetData[],
    private visited: Set<string>
  ) {}

  private peek(offset = 0): FToken | undefined {
    return this.tokens[this.pos + offset];
  }

  private consume(): FToken | undefined {
    return this.tokens[this.pos++];
  }

  parseExpression(): string | number | boolean {
    return this.parseComparison();
  }

  private parseComparison(): string | number | boolean {
    let left = this.parseConcat();
    while (true) {
      const tok = this.peek();
      if (tok?.type === 'op' && ['=', '<>', '<', '<=', '>', '>='].includes(tok.value)) {
        this.consume();
        const right = this.parseConcat();
        left = compareValues(left, right, tok.value);
      } else {
        break;
      }
    }
    return left;
  }

  private parseConcat(): string | number | boolean {
    let left = this.parseAddSub();
    while (true) {
      const tok = this.peek();
      if (tok?.type === 'op' && tok.value === '&') {
        this.consume();
        const right = this.parseAddSub();
        left = `${left ?? ''}${right ?? ''}`;
      } else {
        break;
      }
    }
    return left;
  }

  private parseAddSub(): string | number | boolean {
    let left = this.parseMulDiv();
    while (true) {
      const tok = this.peek();
      if (tok?.type === 'op' && (tok.value === '+' || tok.value === '-')) {
        this.consume();
        const right = this.parseMulDiv();
        const lNum = toNumber(left);
        const rNum = toNumber(right);
        left = tok.value === '+' ? lNum + rNum : lNum - rNum;
      } else {
        break;
      }
    }
    return left;
  }

  private parseMulDiv(): string | number | boolean {
    let left = this.parsePower();
    while (true) {
      const tok = this.peek();
      if (tok?.type === 'op' && (tok.value === '*' || tok.value === '/')) {
        this.consume();
        const right = this.parsePower();
        const lNum = toNumber(left);
        const rNum = toNumber(right);
        if (tok.value === '/' && rNum === 0) return '#DIV/0!';
        left = tok.value === '*' ? lNum * rNum : lNum / rNum;
      } else {
        break;
      }
    }
    return left;
  }

  private parsePower(): string | number | boolean {
    let left = this.parseUnary();
    while (true) {
      const tok = this.peek();
      if (tok?.type === 'op' && tok.value === '^') {
        this.consume();
        const right = this.parseUnary();
        left = Math.pow(toNumber(left), toNumber(right));
      } else {
        break;
      }
    }
    return left;
  }

  private parseUnary(): string | number | boolean {
    const tok = this.peek();
    if (tok?.type === 'op' && (tok.value === '-' || tok.value === '+')) {
      this.consume();
      const val = toNumber(this.parseUnary());
      return tok.value === '-' ? -val : val;
    }
    return this.parsePrimary();
  }

  private parsePrimary(): string | number | boolean {
    const tok = this.peek();
    if (!tok) return 0;

    if (tok.type === 'number') {
      this.consume();
      return tok.value;
    }
    if (tok.type === 'string') {
      this.consume();
      return tok.value;
    }
    if (tok.type === 'bool') {
      this.consume();
      return tok.value;
    }
    if (tok.type === 'paren' && tok.value === '(') {
      this.consume();
      const val = this.parseExpression();
      const closeTok = this.peek();
      if (closeTok?.type === 'paren' && closeTok.value === ')') {
        this.consume();
      }
      return val;
    }

    if (tok.type === 'ident') {
      this.consume();
      // Function call: IDENT(...)
      const openTok = this.peek();
      if (openTok?.type === 'paren' && openTok.value === '(') {
        return this.parseFunctionCall(tok.value.toUpperCase());
      }
      // Cross-sheet reference: SheetName!A1
      let targetSheet = this.activeSheet;
      let cellRefName = tok.value;
      if (this.peek()?.type === 'bang') {
        this.consume();
        const found = this.allSheets.find(
          (s) => s.name.toLowerCase() === tok.value.toLowerCase()
        );
        if (found) targetSheet = found;
        const nextIdent = this.consume();
        if (nextIdent?.type === 'ident') {
          cellRefName = nextIdent.value;
        }
      }

      const coord = parseCellAddress(cellRefName);
      if (coord) {
        const cleanAddr = cellRefName.replace(/\$/g, '').toUpperCase();
        return this.readCellValue(targetSheet, cleanAddr);
      }
      return '#NAME?';
    }

    this.consume();
    return 0;
  }

  private readCellValue(sheet: SheetData, addr: string): string | number | boolean {
    const key = `${sheet.name}!${addr}`;
    const cell: CellData | undefined = sheet.cells[addr];
    if (!cell) return 0;
    if (cell.formula && cell.formula.startsWith('=') && !this.visited.has(key)) {
      const nextVisited = new Set(this.visited);
      nextVisited.add(key);
      return evaluateFormula(cell.formula, sheet, this.allSheets, nextVisited);
    }
    if (cell.value === null || cell.value === undefined || cell.value === '') return 0;
    return cell.value;
  }

  private parseRangeOrValueArg(): Array<string | number | boolean | null> {
    // Check if argument looks like [Sheet!]A1:B5
    const savedPos = this.pos;
    const t0 = this.peek(0);
    if (t0?.type === 'ident') {
      let targetSheet = this.activeSheet;
      let firstRef = t0.value;
      let offset = 1;
      if (this.peek(1)?.type === 'bang' && this.peek(2)?.type === 'ident') {
        const found = this.allSheets.find(
          (s) => s.name.toLowerCase() === t0.value.toLowerCase()
        );
        if (found) targetSheet = found;
        firstRef = (this.peek(2) as { type: 'ident'; value: string }).value;
        offset = 3;
      }
      if (
        this.peek(offset)?.type === 'colon' &&
        this.peek(offset + 1)?.type === 'ident'
      ) {
        const secondRef = (this.peek(offset + 1) as { type: 'ident'; value: string }).value;
        const range = parseRangeAddress(`${firstRef}:${secondRef}`);
        if (range) {
          this.pos += offset + 2;
          // If followed directly by comma or ')', treat as a true range list
          const next = this.peek();
          if (!next || next.type === 'comma' || (next.type === 'paren' && next.value === ')')) {
            const coords = iterateRangeCoords(range);
            return coords.map((c) => {
              const addr = `${colLetter(c.col)}${c.row}`;
              const cell = targetSheet.cells[addr];
              if (!cell || cell.value === '' || cell.value === null || cell.value === undefined) {
                return null;
              }
              return cell.value;
            });
          }
          this.pos = savedPos;
        }
      }
    }
    return [this.parseExpression()];
  }

  private parseFunctionCall(name: string): string | number | boolean {
    this.consume(); // consume '('
    const argLists: Array<Array<string | number | boolean | null>> = [];

    const firstArgTok = this.peek();
    if (!(firstArgTok?.type === 'paren' && firstArgTok.value === ')')) {
      while (this.pos < this.tokens.length) {
        argLists.push(this.parseRangeOrValueArg());
        if (this.peek()?.type === 'comma') {
          this.consume();
        } else {
          break;
        }
      }
    }
    const endArgTok = this.peek();
    if (endArgTok?.type === 'paren' && endArgTok.value === ')') {
      this.consume();
    }

    const flatValues = argLists.flat();
    const numericValues = flatValues
      .filter((v): v is string | number => v !== null && v !== '' && typeof v !== 'boolean')
      .map((v) => (typeof v === 'number' ? v : Number(v)))
      .filter((n) => Number.isFinite(n));

    switch (name) {
      case 'SUM':
        return numericValues.reduce((acc, n) => acc + n, 0);
      case 'AVERAGE':
      case 'AVG':
        return numericValues.length
          ? numericValues.reduce((acc, n) => acc + n, 0) / numericValues.length
          : '#DIV/0!';
      case 'COUNT':
        return numericValues.length;
      case 'COUNTA':
        return flatValues.filter((v) => v !== null && v !== '').length;
      case 'MIN':
        return numericValues.length ? Math.min(...numericValues) : 0;
      case 'MAX':
        return numericValues.length ? Math.max(...numericValues) : 0;
      case 'ROUND': {
        const val = toNumber(argLists[0]?.[0] ?? 0);
        const digits = toNumber(argLists[1]?.[0] ?? 0);
        const factor = Math.pow(10, digits);
        return Math.round(val * factor) / factor;
      }
      case 'ABS':
        return Math.abs(toNumber(argLists[0]?.[0] ?? 0));
      case 'INT':
        return Math.floor(toNumber(argLists[0]?.[0] ?? 0));
      case 'IF': {
        const cond = Boolean(argLists[0]?.[0]);
        return cond
          ? (argLists[1]?.[0] ?? true)
          : (argLists[2]?.[0] ?? false);
      }
      case 'AND':
        return flatValues.every((v) => Boolean(v));
      case 'OR':
        return flatValues.some((v) => Boolean(v));
      case 'LEN':
        return String(argLists[0]?.[0] ?? '').length;
      case 'UPPER':
        return String(argLists[0]?.[0] ?? '').toUpperCase();
      case 'LOWER':
        return String(argLists[0]?.[0] ?? '').toLowerCase();
      case 'CONCAT':
      case 'CONCATENATE':
        return flatValues.map((v) => (v === null ? '' : String(v))).join('');
      case 'SUMIF': {
        const rangeVals = argLists[0] ?? [];
        const criteria = argLists[1]?.[0] ?? '';
        const sumVals = argLists[2] ?? rangeVals;
        let total = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (matchesCriteria(rangeVals[i], criteria)) {
            const n = Number(sumVals[i]);
            if (Number.isFinite(n)) total += n;
          }
        }
        return total;
      }
      case 'COUNTIF': {
        const rangeVals = argLists[0] ?? [];
        const criteria = argLists[1]?.[0] ?? '';
        return rangeVals.filter((v) => matchesCriteria(v, criteria)).length;
      }
      default:
        return '#NAME?';
    }
  }
}

function colLetter(col: number): string {
  let res = '';
  let n = col;
  while (n > 0) {
    const r = (n - 1) % 26;
    res = String.fromCharCode(65 + r) + res;
    n = Math.floor((n - 1) / 26);
  }
  return res;
}

function toNumber(v: unknown): number {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const cleaned = v.replace(/[$,]/g, '').trim();
    if (!cleaned) return 0;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}

function compareValues(a: unknown, b: unknown, op: string): boolean {
  const numA = typeof a === 'number' ? a : Number(a);
  const numB = typeof b === 'number' ? b : Number(b);
  const bothNumeric =
    a !== '' &&
    b !== '' &&
    typeof a !== 'boolean' &&
    typeof b !== 'boolean' &&
    Number.isFinite(numA) &&
    Number.isFinite(numB);

  const l = bothNumeric ? numA : String(a ?? '').toLowerCase();
  const r = bothNumeric ? numB : String(b ?? '').toLowerCase();

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

function matchesCriteria(val: string | number | boolean | null, criteria: unknown): boolean {
  if (val === null) return false;
  const critStr = String(criteria ?? '').trim();
  const opMatch = /^(>=|<=|<>|>|<|=)(.*)$/.exec(critStr);
  if (opMatch) {
    return compareValues(val, opMatch[2].trim(), opMatch[1]);
  }
  return String(val).toLowerCase() === critStr.toLowerCase();
}
