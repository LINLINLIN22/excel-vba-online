export interface VbaSuggestion {
  label: string;
  insertText: string;
  kind: 'object' | 'property' | 'keyword' | 'function' | 'snippet' | 'constant';
  signature?: string;
  description: string;
}

export const VBA_DOT_PROPERTIES: VbaSuggestion[] = [
  {
    label: 'Value',
    insertText: 'Value',
    kind: 'property',
    signature: '.Value = expr',
    description: 'Read or set the value inside a cell or range.',
  },
  {
    label: 'Formula',
    insertText: 'Formula = "=SUM()"',
    kind: 'property',
    signature: '.Formula = "=SUM(A1:A5)"',
    description: 'Assign an Excel formula to the cell.',
  },
  {
    label: 'NumberFormat',
    insertText: 'NumberFormat = "$#,##0"',
    kind: 'property',
    signature: '.NumberFormat = "$#,##0"',
    description: 'Set display format (e.g. "$#,##0", "0.0%", "#,##0.00").',
  },
  {
    label: 'Interior.Color',
    insertText: 'Interior.Color = RGB(33, 115, 70)',
    kind: 'property',
    signature: '.Interior.Color = RGB(r, g, b)',
    description: 'Set the cell background fill color.',
  },
  {
    label: 'Font.Color',
    insertText: 'Font.Color = RGB(255, 255, 255)',
    kind: 'property',
    signature: '.Font.Color = RGB(r, g, b)',
    description: 'Set the text color of the range.',
  },
  {
    label: 'Font.Bold',
    insertText: 'Font.Bold = True',
    kind: 'property',
    signature: '.Font.Bold = True | False',
    description: 'Make cell text bold.',
  },
  {
    label: 'Font.Italic',
    insertText: 'Font.Italic = True',
    kind: 'property',
    signature: '.Font.Italic = True | False',
    description: 'Make cell text italic.',
  },
  {
    label: 'Font.Size',
    insertText: 'Font.Size = 12',
    kind: 'property',
    signature: '.Font.Size = 12',
    description: 'Change font size in points.',
  },
  {
    label: 'HorizontalAlignment',
    insertText: 'HorizontalAlignment = xlCenter',
    kind: 'property',
    signature: '.HorizontalAlignment = xlLeft | xlCenter | xlRight',
    description: 'Align cell content horizontally.',
  },
  {
    label: 'WrapText',
    insertText: 'WrapText = True',
    kind: 'property',
    signature: '.WrapText = True | False',
    description: 'Wrap long lines of text within cells.',
  },
  {
    label: 'Borders.LineStyle',
    insertText: 'Borders.LineStyle = xlContinuous',
    kind: 'property',
    signature: '.Borders.LineStyle = xlContinuous',
    description: 'Apply continuous thin borders around cells.',
  },
  {
    label: 'Offset',
    insertText: 'Offset(1, 0)',
    kind: 'function',
    signature: '.Offset(rowOffset, colOffset)',
    description: 'Shift a range relative to its current position.',
  },
  {
    label: 'ClearContents',
    insertText: 'ClearContents',
    kind: 'function',
    signature: '.ClearContents',
    description: 'Clear cell values while keeping formatting.',
  },
  {
    label: 'Clear',
    insertText: 'Clear',
    kind: 'function',
    signature: '.Clear',
    description: 'Clear both cell values and formatting.',
  },
  {
    label: 'End(xlUp)',
    insertText: 'End(xlUp)',
    kind: 'function',
    signature: '.End(xlUp | xlDown | xlToLeft | xlToRight)',
    description: 'Jump to the last non-empty cell in a direction.',
  },
  {
    label: 'Row',
    insertText: 'Row',
    kind: 'property',
    description: 'Get the 1-based row number of the first cell.',
  },
  {
    label: 'Column',
    insertText: 'Column',
    kind: 'property',
    description: 'Get the 1-based column number of the first cell.',
  },
];

export const VBA_GLOBAL_SUGGESTIONS: VbaSuggestion[] = [
  {
    label: 'Range("A1")',
    insertText: 'Range("A1")',
    kind: 'object',
    signature: 'Range("A1:C5")',
    description: 'Reference a single cell or rectangular range.',
  },
  {
    label: 'Cells(row, col)',
    insertText: 'Cells(1, 1)',
    kind: 'object',
    signature: 'Cells(rowIndex, colIndex)',
    description: 'Reference a cell dynamically using numeric row and column.',
  },
  {
    label: 'With ... End With',
    insertText: 'With Range("A1:C1")\n    .Font.Bold = True\nEnd With',
    kind: 'snippet',
    signature: 'With object ... End With',
    description: 'Apply multiple properties to the same range cleanly.',
  },
  {
    label: 'For ... Next',
    insertText: 'For i = 2 To 6\n    Cells(i, 4).Value = Cells(i, 2).Value * Cells(i, 3).Value\nNext i',
    kind: 'snippet',
    signature: 'For counter = start To end ... Next counter',
    description: 'Loop through a numeric range of rows or columns.',
  },
  {
    label: 'For Each ... In',
    insertText: 'For Each cell In Range("C2:C4")\n    \nNext cell',
    kind: 'snippet',
    signature: 'For Each cell In Range(...) ... Next cell',
    description: 'Iterate over every cell in a selected range.',
  },
  {
    label: 'If ... Then ... End If',
    insertText: 'If Range("C2").Value > 300 Then\n    Range("C2").Font.Bold = True\nEnd If',
    kind: 'snippet',
    signature: 'If condition Then ... Else ... End If',
    description: 'Execute statements conditionally.',
  },
  {
    label: 'RGB(r, g, b)',
    insertText: 'RGB(33, 115, 70)',
    kind: 'function',
    signature: 'RGB(red, green, blue)',
    description: 'Create an Excel color from 0–255 RGB channels.',
  },
  {
    label: 'Sheets("Sheet1")',
    insertText: 'Sheets("Sheet1")',
    kind: 'object',
    signature: 'Sheets(nameOrIndex)',
    description: 'Target a specific worksheet in the workbook.',
  },
  {
    label: 'Debug.Print',
    insertText: 'Debug.Print ',
    kind: 'function',
    signature: 'Debug.Print expression',
    description: 'Write values or diagnostics to the Execution log.',
  },
  {
    label: 'MsgBox',
    insertText: 'MsgBox "Done!"',
    kind: 'function',
    signature: 'MsgBox prompt, [buttons], [title]',
    description: 'Display an alert dialog message to the user.',
  },
  {
    label: 'Dim ... As',
    insertText: 'Dim total As Double',
    kind: 'keyword',
    signature: 'Dim varName As DataType',
    description: 'Declare a typed local variable (Integer, Double, String, Range).',
  },
  {
    label: 'WorksheetFunction.Sum',
    insertText: 'WorksheetFunction.Sum(Range("C2:C4"))',
    kind: 'function',
    signature: 'WorksheetFunction.Sum(range)',
    description: 'Call Excel Sum, Average, Min, Max, or Count inside VBA.',
  },
  {
    label: 'Sub ... End Sub',
    insertText: 'Sub MyMacro()\n    \nEnd Sub',
    kind: 'snippet',
    signature: 'Sub MacroName() ... End Sub',
    description: 'Define a runnable VBA macro procedure.',
  },
  {
    label: 'ActiveSheet',
    insertText: 'ActiveSheet',
    kind: 'object',
    description: 'Refer to the currently active worksheet.',
  },
  {
    label: 'ActiveCell',
    insertText: 'ActiveCell',
    kind: 'object',
    description: 'Refer to the currently selected cell on the sheet.',
  },
  {
    label: 'xlCenter',
    insertText: 'xlCenter',
    kind: 'constant',
    description: 'Center horizontal alignment constant.',
  },
  {
    label: 'xlLeft',
    insertText: 'xlLeft',
    kind: 'constant',
    description: 'Left horizontal alignment constant.',
  },
  {
    label: 'xlRight',
    insertText: 'xlRight',
    kind: 'constant',
    description: 'Right horizontal alignment constant.',
  },
  {
    label: 'vbRed / vbGreen / vbYellow',
    insertText: 'vbRed',
    kind: 'constant',
    description: 'Built-in VBA color constants (vbRed, vbGreen, vbBlue, vbYellow).',
  },
];

export const QUICK_HINT_CHIPS = [
  { label: 'Range()', insert: 'Range("A1")' },
  { label: 'Cells(r, c)', insert: 'Cells(1, 1)' },
  { label: 'With ... End With', insert: 'With Range("A1:C1")\n    .Font.Bold = True\nEnd With' },
  { label: '.Interior.Color', insert: '.Interior.Color = RGB(33, 115, 70)' },
  { label: '.NumberFormat', insert: '.NumberFormat = "$#,##0"' },
  { label: 'For ... Next', insert: 'For i = 2 To 5\n    \nNext i' },
  { label: 'If ... Then', insert: 'If  Then\n    \nEnd If' },
  { label: 'Debug.Print', insert: 'Debug.Print ' },
];

export function getAutocompleteSuggestions(
  code: string,
  cursorPos: number
): { prefix: string; replaceStart: number; items: VbaSuggestion[] } {
  const beforeCursor = code.slice(0, cursorPos);
  const lineStart = beforeCursor.lastIndexOf('\n') + 1;
  const currentLine = beforeCursor.slice(lineStart);

  // Inside comment -> no suggestions
  let inQuotes = false;
  for (let i = 0; i < currentLine.length; i++) {
    if (currentLine[i] === '"') inQuotes = !inQuotes;
    if (!inQuotes && currentLine[i] === "'") {
      return { prefix: '', replaceStart: cursorPos, items: [] };
    }
  }

  // After dot: e.g. ".Int" or "Range("A1").Fon"
  const dotMatch = /\.([A-Za-z0-9_]*)$/.exec(currentLine);
  if (dotMatch) {
    const prefix = dotMatch[1];
    const query = prefix.toLowerCase();
    const items = VBA_DOT_PROPERTIES.filter((s) =>
      s.label.toLowerCase().includes(query)
    );
    return {
      prefix,
      replaceStart: cursorPos - prefix.length,
      items: items.slice(0, 8),
    };
  }

  // Word prefix at cursor
  const wordMatch = /([A-Za-z_][A-Za-z0-9_]*)$/.exec(currentLine);
  if (wordMatch) {
    const prefix = wordMatch[1];
    const query = prefix.toLowerCase();
    if (query.length < 1) {
      return { prefix: '', replaceStart: cursorPos, items: [] };
    }
    const items = VBA_GLOBAL_SUGGESTIONS.filter(
      (s) =>
        s.label.toLowerCase().startsWith(query) ||
        s.label.toLowerCase().includes(query)
    ).sort((a, b) => {
      const aStarts = a.label.toLowerCase().startsWith(query) ? 0 : 1;
      const bStarts = b.label.toLowerCase().startsWith(query) ? 0 : 1;
      return aStarts - bStarts;
    });
    return {
      prefix,
      replaceStart: cursorPos - prefix.length,
      items: items.slice(0, 8),
    };
  }

  return { prefix: '', replaceStart: cursorPos, items: [] };
}
