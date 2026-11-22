export type VbaTokenType =
  | 'keyword'
  | 'builtin'
  | 'constant'
  | 'string'
  | 'number'
  | 'comment'
  | 'operator'
  | 'punctuation'
  | 'property'
  | 'ident'
  | 'whitespace';

export interface VbaToken {
  type: VbaTokenType;
  value: string;
  line: number; // 1-indexed
  col: number; // 1-indexed
}

export const VBA_KEYWORDS = new Set([
  'sub',
  'function',
  'end',
  'dim',
  'as',
  'set',
  'let',
  'const',
  'option',
  'explicit',
  'with',
  'for',
  'each',
  'in',
  'to',
  'step',
  'next',
  'if',
  'then',
  'elseif',
  'else',
  'do',
  'while',
  'until',
  'loop',
  'exit',
  'select',
  'case',
  'call',
  'byval',
  'byref',
  'and',
  'or',
  'not',
  'mod',
  'xor',
  'integer',
  'long',
  'double',
  'single',
  'string',
  'boolean',
  'variant',
  'object',
  'range',
  'worksheet',
  'workbook',
]);

export const VBA_BUILTINS = new Set([
  'range',
  'cells',
  'sheets',
  'worksheets',
  'activesheet',
  'activecell',
  'selection',
  'thisworkbook',
  'activeworkbook',
  'rows',
  'columns',
  'rgb',
  'msgbox',
  'inputbox',
  'debug',
  'worksheetfunction',
  'application',
  'len',
  'left',
  'right',
  'mid',
  'trim',
  'ltrim',
  'rtrim',
  'ucase',
  'lcase',
  'replace',
  'instr',
  'split',
  'join',
  'cstr',
  'cint',
  'clng',
  'cdbl',
  'val',
  'int',
  'round',
  'abs',
  'sqr',
  'isnumeric',
  'isempty',
  'iif',
  'format',
  'now',
  'date',
  'year',
  'month',
  'day',
]);

export const VBA_CONSTANTS = new Set([
  'true',
  'false',
  'nothing',
  'empty',
  'null',
  'vbred',
  'vbgreen',
  'vbblue',
  'vbyellow',
  'vbwhite',
  'vbblack',
  'vbcyan',
  'vbmagenta',
  'vbcrlf',
  'vblf',
  'vbnewline',
  'vbtab',
  'vbokonly',
  'vbokcancel',
  'vbinformation',
  'vbexclamation',
  'vbcritical',
  'xlleft',
  'xlcenter',
  'xlright',
  'xlup',
  'xldown',
  'xltoleft',
  'xltoright',
  'xlnone',
  'xlcontinuous',
]);

export function tokenizeVbaLine(lineText: string, lineNum: number): VbaToken[] {
  const tokens: VbaToken[] = [];
  let i = 0;

  while (i < lineText.length) {
    const ch = lineText[i];

    // Whitespace
    if (ch === ' ' || ch === '\t' || ch === '\r') {
      const start = i;
      while (i < lineText.length && (lineText[i] === ' ' || lineText[i] === '\t' || lineText[i] === '\r')) {
        i++;
      }
      tokens.push({
        type: 'whitespace',
        value: lineText.slice(start, i),
        line: lineNum,
        col: start + 1,
      });
      continue;
    }

    // Comment starting with '
    if (ch === "'") {
      tokens.push({
        type: 'comment',
        value: lineText.slice(i),
        line: lineNum,
        col: i + 1,
      });
      break;
    }

    // String literal "..."
    if (ch === '"') {
      const start = i;
      i++;
      while (i < lineText.length) {
        if (lineText[i] === '"' && lineText[i + 1] === '"') {
          i += 2;
        } else if (lineText[i] === '"') {
          i++;
          break;
        } else {
          i++;
        }
      }
      tokens.push({
        type: 'string',
        value: lineText.slice(start, i),
        line: lineNum,
        col: start + 1,
      });
      continue;
    }

    // Number literal
    if (/[0-9]/.test(ch)) {
      const start = i;
      while (i < lineText.length && /[0-9.]/.test(lineText[i])) {
        i++;
      }
      tokens.push({
        type: 'number',
        value: lineText.slice(start, i),
        line: lineNum,
        col: start + 1,
      });
      continue;
    }

    // Multi-char operators
    const twoChar = lineText.slice(i, i + 2);
    if (['<>', '<=', '>=', ':='].includes(twoChar)) {
      tokens.push({
        type: 'operator',
        value: twoChar,
        line: lineNum,
        col: i + 1,
      });
      i += 2;
      continue;
    }

    // Single-char operators
    if ('=<>+-*/\\^&'.includes(ch)) {
      tokens.push({
        type: 'operator',
        value: ch,
        line: lineNum,
        col: i + 1,
      });
      i++;
      continue;
    }

    // Punctuation
    if ('.,():'.includes(ch)) {
      tokens.push({
        type: 'punctuation',
        value: ch,
        line: lineNum,
        col: i + 1,
      });
      i++;
      continue;
    }

    // Identifiers & Keywords
    if (/[A-Za-z_\u4e00-\u9fff]/.test(ch)) {
      const start = i;
      while (i < lineText.length && /[A-Za-z0-9_\u4e00-\u9fff]/.test(lineText[i])) {
        i++;
      }
      const word = lineText.slice(start, i);
      const lower = word.toLowerCase();

      // Check REM comment at statement start
      if (lower === 'rem' && tokens.every((t) => t.type === 'whitespace')) {
        tokens.push({
          type: 'comment',
          value: lineText.slice(start),
          line: lineNum,
          col: start + 1,
        });
        break;
      }

      const prevNonWs = [...tokens].reverse().find((t) => t.type !== 'whitespace');
      const isAfterDot = prevNonWs?.type === 'punctuation' && prevNonWs.value === '.';

      let type: VbaTokenType = 'ident';
      if (isAfterDot) {
        type = 'property';
      } else if (VBA_CONSTANTS.has(lower)) {
        type = 'constant';
      } else if (VBA_KEYWORDS.has(lower)) {
        type = 'keyword';
      } else if (VBA_BUILTINS.has(lower)) {
        type = 'builtin';
      }

      tokens.push({
        type,
        value: word,
        line: lineNum,
        col: start + 1,
      });
      continue;
    }

    // Fallback single character
    tokens.push({
      type: 'punctuation',
      value: ch,
      line: lineNum,
      col: i + 1,
    });
    i++;
  }

  return tokens;
}

export function tokenizeVbaCode(code: string): VbaToken[][] {
  return code.split('\n').map((line, idx) => tokenizeVbaLine(line, idx + 1));
}
