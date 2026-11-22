import { VbaDiagnostic } from '../types/workbook';
import { tokenizeVbaLine, VbaToken } from './lexer';

interface BlockFrame {
  kind: 'sub' | 'function' | 'with' | 'for' | 'if' | 'do' | 'select';
  line: number;
  label: string;
}

export function lintVbaCode(code: string): VbaDiagnostic[] {
  const diagnostics: VbaDiagnostic[] = [];
  const lines = code.split('\n');
  const stack: BlockFrame[] = [];
  let hasSubOrFunc = false;

  for (let i = 0; i < lines.length; i++) {
    const lineNum = i + 1;
    const rawLine = lines[i];

    // Check unclosed string literal on this line
    let inString = false;
    for (let c = 0; c < rawLine.length; c++) {
      const ch = rawLine[c];
      if (!inString && ch === "'") break;
      if (ch === '"') {
        if (inString && rawLine[c + 1] === '"') {
          c++;
        } else {
          inString = !inString;
        }
      }
    }
    if (inString) {
      diagnostics.push({
        line: lineNum,
        col: rawLine.indexOf('"') + 1,
        severity: 'error',
        message: 'Unclosed string literal (missing closing ").',
      });
    }

    const tokens = tokenizeVbaLine(rawLine, lineNum).filter(
      (t) => t.type !== 'whitespace' && t.type !== 'comment'
    );
    if (tokens.length === 0) continue;

    // Check balanced parentheses on the line
    let parenDepth = 0;
    for (const tok of tokens) {
      if (tok.value === '(') parenDepth++;
      if (tok.value === ')') {
        parenDepth--;
        if (parenDepth < 0) {
          diagnostics.push({
            line: lineNum,
            col: tok.col,
            severity: 'error',
            message: 'Unexpected closing parenthesis ")".',
          });
          break;
        }
      }
    }
    if (parenDepth > 0) {
      diagnostics.push({
        line: lineNum,
        col: tokens[tokens.length - 1].col,
        severity: 'error',
        message: 'Missing closing parenthesis ")".',
      });
    }

    const first = tokens[0];
    const firstLower = first.value.toLowerCase();
    const secondLower = tokens[1]?.value.toLowerCase() ?? '';

    // Check leading dot outside With block
    if (first.value === '.' && !stack.some((b) => b.kind === 'with')) {
      diagnostics.push({
        line: lineNum,
        col: first.col,
        severity: 'error',
        message: 'Invalid unqualified reference "." outside of a With ... End With block.',
      });
    }

    // Block opening/closing checks
    if (
      (firstLower === 'sub' || firstLower === 'function') ||
      ((firstLower === 'public' || firstLower === 'private') &&
        (secondLower === 'sub' || secondLower === 'function'))
    ) {
      hasSubOrFunc = true;
      const isFunc = firstLower === 'function' || secondLower === 'function';
      if (stack.some((b) => b.kind === 'sub' || b.kind === 'function')) {
        diagnostics.push({
          line: lineNum,
          col: first.col,
          severity: 'error',
          message: 'Nested Sub/Function is not allowed. Close the previous procedure with End Sub/End Function.',
        });
      }
      stack.push({
        kind: isFunc ? 'function' : 'sub',
        line: lineNum,
        label: isFunc ? 'Function' : 'Sub',
      });
      continue;
    }

    if (firstLower === 'end') {
      if (secondLower === 'sub' || secondLower === 'function') {
        const expected = secondLower as 'sub' | 'function';
        const top = stack[stack.length - 1];
        if (!top) {
          diagnostics.push({
            line: lineNum,
            col: first.col,
            severity: 'error',
            message: `Unexpected "End ${secondLower === 'sub' ? 'Sub' : 'Function'}" without matching opening declaration.`,
          });
        } else if (top.kind !== expected) {
          diagnostics.push({
            line: lineNum,
            col: first.col,
            severity: 'error',
            message: `Expected "End ${top.label}" to close ${top.label} on line ${top.line}, found "End ${secondLower}".`,
          });
        } else {
          stack.pop();
        }
        continue;
      }

      if (secondLower === 'with') {
        popExpected(stack, 'with', 'End With', lineNum, first.col, diagnostics);
        continue;
      }

      if (secondLower === 'if') {
        popExpected(stack, 'if', 'End If', lineNum, first.col, diagnostics);
        continue;
      }

      if (secondLower === 'select') {
        popExpected(stack, 'select', 'End Select', lineNum, first.col, diagnostics);
        continue;
      }
    }

    if (firstLower === 'endif') {
      popExpected(stack, 'if', 'End If', lineNum, first.col, diagnostics);
      continue;
    }

    if (firstLower === 'with') {
      if (tokens.length < 2) {
        diagnostics.push({
          line: lineNum,
          col: first.col,
          severity: 'error',
          message: 'With statement requires an object expression, e.g. With Range("A1:C1").',
        });
      }
      stack.push({ kind: 'with', line: lineNum, label: 'With' });
      continue;
    }

    if (firstLower === 'for') {
      stack.push({ kind: 'for', line: lineNum, label: 'For' });
      continue;
    }

    if (firstLower === 'next') {
      popExpected(stack, 'for', 'Next', lineNum, first.col, diagnostics);
      continue;
    }

    if (firstLower === 'do') {
      stack.push({ kind: 'do', line: lineNum, label: 'Do' });
      continue;
    }

    if (firstLower === 'loop') {
      popExpected(stack, 'do', 'Loop', lineNum, first.col, diagnostics);
      continue;
    }

    if (firstLower === 'select' && secondLower === 'case') {
      stack.push({ kind: 'select', line: lineNum, label: 'Select Case' });
      continue;
    }

    if (firstLower === 'if') {
      const thenIndex = findKeywordIndex(tokens, 'then');
      if (thenIndex === -1) {
        diagnostics.push({
          line: lineNum,
          col: first.col,
          severity: 'error',
          message: 'Expected "Then" after If condition.',
        });
      } else if (thenIndex === tokens.length - 1) {
        // Multi-line If block
        stack.push({ kind: 'if', line: lineNum, label: 'If' });
      }
      continue;
    }

    if (firstLower === 'elseif') {
      const thenIndex = findKeywordIndex(tokens, 'then');
      if (thenIndex === -1) {
        diagnostics.push({
          line: lineNum,
          col: first.col,
          severity: 'error',
          message: 'Expected "Then" after ElseIf condition.',
        });
      }
      if (!stack.some((b) => b.kind === 'if')) {
        diagnostics.push({
          line: lineNum,
          col: first.col,
          severity: 'error',
          message: '"ElseIf" used without a matching "If ... Then" block.',
        });
      }
      continue;
    }

    if (firstLower === 'else') {
      if (!stack.some((b) => b.kind === 'if')) {
        diagnostics.push({
          line: lineNum,
          col: first.col,
          severity: 'error',
          message: '"Else" used without a matching "If ... Then" block.',
        });
      }
    }
  }

  for (const unclosed of stack) {
    const expectedClose: Record<BlockFrame['kind'], string> = {
      sub: 'End Sub',
      function: 'End Function',
      with: 'End With',
      for: 'Next',
      if: 'End If',
      do: 'Loop',
      select: 'End Select',
    };
    diagnostics.push({
      line: unclosed.line,
      col: 1,
      severity: 'error',
      message: `${unclosed.label} block opened on line ${unclosed.line} is missing "${expectedClose[unclosed.kind]}".`,
    });
  }

  if (!hasSubOrFunc && code.trim().length > 0) {
    diagnostics.push({
      line: 1,
      col: 1,
      severity: 'warning',
      message: 'Wrap your macro in a Sub MyMacro() ... End Sub procedure.',
    });
  }

  return diagnostics;
}

function findKeywordIndex(tokens: VbaToken[], kw: string): number {
  return tokens.findIndex((t) => t.value.toLowerCase() === kw);
}

function popExpected(
  stack: BlockFrame[],
  expectedKind: BlockFrame['kind'],
  closingLabel: string,
  lineNum: number,
  col: number,
  diagnostics: VbaDiagnostic[]
) {
  if (stack.length === 0) {
    diagnostics.push({
      line: lineNum,
      col,
      severity: 'error',
      message: `Unexpected "${closingLabel}" without matching opening block.`,
    });
    return;
  }
  const top = stack[stack.length - 1];
  if (top.kind !== expectedKind) {
    diagnostics.push({
      line: lineNum,
      col,
      severity: 'error',
      message: `Mismatched block: expected closing for "${top.label}" (line ${top.line}) before "${closingLabel}".`,
    });
    const idx = [...stack].reverse().findIndex((b) => b.kind === expectedKind);
    if (idx !== -1) {
      stack.splice(stack.length - 1 - idx, 1);
    }
    return;
  }
  stack.pop();
}
