import { LearnLesson } from '../types/workbook';

export const LEARN_LESSONS: LearnLesson[] = [
  {
    id: 'lesson-ranges',
    number: '01',
    title: 'Targeting Cells with Range & Cells',
    summary:
      'Every Excel macro starts by pointing at cells. Use Range("A1:C1") for fixed addresses and Cells(row, col) when row numbers change in a loop.',
    concept:
      'Range accepts standard A1 notation ("B2", "A1:C5"), while Cells takes numeric (row, column) coordinates so you can iterate with variables.',
    code: `Sub TargetCells()
    Range("A1").Value = "Product"
    Range("B1:C1").Font.Bold = True
    Cells(2, 1).Value = "Notebook"
    Cells(2, 2).Value = 24
End Sub`,
    keyTakeaways: [
      'Range("A1").Value reads or writes a single cell.',
      'Range("A1:C1") applies styles or values across an entire rectangle at once.',
      'Cells(2, 3) points to row 2, column 3 (C2) using numbers.',
    ],
    presetId: 'format-sales',
  },
  {
    id: 'lesson-with',
    number: '02',
    title: 'Writing Clean Styles with With ... End With',
    summary:
      'Instead of repeating Range("A1:C1") on every line, open a With block and start each property with a leading dot.',
    concept:
      'With blocks keep your macros readable and reduce typing mistakes. You can configure fill colors, font colors, borders, and alignment together.',
    code: `With Range("A1:C1")
    .Interior.Color = RGB(33, 115, 70)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With`,
    keyTakeaways: [
      'RGB(red, green, blue) accepts values from 0 to 255.',
      'Always close every With block with End With.',
      'Type a dot "." inside the editor to trigger property hints.',
    ],
    presetId: 'format-sales',
  },
  {
    id: 'lesson-loops',
    number: '03',
    title: 'Automating Rows with For ... Next & For Each',
    summary:
      'Loops let a 5-line macro process 5 rows or 5,000 rows without manual copy-paste.',
    concept:
      'Use For i = start To end when you need row indices, or For Each cell In Range(...) when inspecting every cell in a selection.',
    code: `For r = 2 To 6
    If Cells(r, 2).Value < 15 Then
        Cells(r, 4).Value = "Reorder"
        Cells(r, 4).Font.Bold = True
    End If
Next r`,
    keyTakeaways: [
      'Dim r As Integer declares your loop counter cleanly.',
      'Combine If ... Then ... Else ... End If inside loops to classify rows.',
      'Use Cells(Rows.Count, 1).End(xlUp).Row to detect the last used row dynamically.',
    ],
    presetId: 'highlight-low-stock',
  },
  {
    id: 'lesson-formulas',
    number: '04',
    title: 'Injecting Formulas & Number Formats',
    summary:
      'Assign live Excel formulas via .Formula and control how currencies or percentages render via .NumberFormat.',
    concept:
      'Setting .Formula = "=SUM(C2:C4)" keeps the spreadsheet reactive—if you edit C2 later, C6 recalculates automatically.',
    code: `Range("C2:C6").NumberFormat = "$#,##0"
Range("C6").Formula = "=SUM(C2:C4)"
Range("A6").Value = "Total"`,
    keyTakeaways: [
      '.Value writes a static result; .Formula writes a live formula.',
      'Use "$#,##0" for currency and "0.0%" for percentages.',
      'WorksheetFunction.Sum(...) lets VBA compute totals in memory.',
    ],
    presetId: 'commission-tiers',
  },
  {
    id: 'lesson-debug',
    number: '05',
    title: 'Inspecting & Debugging with Debug.Print & MsgBox',
    summary:
      'Verify variable values and catch edge cases before running a macro on critical workbooks.',
    concept:
      'Debug.Print logs messages into the Execution panel below the editor, while MsgBox displays a popup notification when a task finishes.',
    code: `Dim total As Double
total = WorksheetFunction.Sum(Range("C2:C4"))
Debug.Print "Calculated total: " & total
MsgBox "Done! Total is $" & total`,
    keyTakeaways: [
      'Use & to concatenate text strings and numeric variables.',
      'Syntax errors highlight the exact line number before execution.',
      'Press Ctrl/Cmd + Enter anytime to test your macro immediately.',
    ],
    presetId: 'multi-sheet-rollup',
  },
];

export const QUICK_GUIDE_SECTIONS = [
  {
    title: 'Keyboard Shortcuts',
    items: [
      { key: 'Ctrl / ⌘ + Enter', desc: 'Run the active macro immediately' },
      { key: 'Ctrl + Space', desc: 'Open VBA command autocomplete hints' },
      { key: 'Tab / Enter', desc: 'Accept highlighted autocomplete command' },
      { key: 'Tab (in editor)', desc: 'Insert 4 spaces indentation' },
    ],
  },
  {
    title: 'Supported Excel Objects',
    items: [
      { key: 'Range("A1:C6")', desc: 'Read/write values, formulas, and styles' },
      { key: 'Cells(row, col)', desc: 'Dynamic numeric cell addressing' },
      { key: 'Sheets("Name")', desc: 'Target any worksheet in the workbook' },
      { key: '.Offset(r, c) / .End(xlUp)', desc: 'Relative navigation & last-row detection' },
      { key: 'WorksheetFunction', desc: 'Sum, Average, Count, Min, Max, Round' },
    ],
  },
  {
    title: 'Formatting & Control Flow',
    items: [
      { key: 'With ... End With', desc: 'Group multiple range properties cleanly' },
      { key: 'For ... Next / For Each', desc: 'Row and cell-range iteration' },
      { key: 'If ... ElseIf ... End If', desc: 'Conditional logic & Select Case' },
      { key: 'RGB(r, g, b)', desc: 'Custom colors for .Interior.Color & .Font.Color' },
      { key: 'Debug.Print / MsgBox', desc: 'Log diagnostics & alert messages' },
    ],
  },
];
