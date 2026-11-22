import { ExamplePreset, SheetData } from '../types/workbook';

function baseSheet(
  name: string,
  cells: SheetData['cells'],
  rows = 15,
  cols = 8
): SheetData {
  return {
    id: `sheet-${name.toLowerCase().replace(/\s+/g, '-')}`,
    name,
    rows,
    cols,
    cells,
    colWidths: {
      1: 118,
      2: 96,
      3: 108,
      4: 96,
      5: 96,
      6: 96,
      7: 96,
      8: 96,
    },
  };
}

export const EXAMPLE_PRESETS: ExamplePreset[] = [
  {
    id: 'format-sales',
    title: 'Format a sales report',
    category: 'Formatting',
    difficulty: 'Beginner',
    description:
      'Style a table header with RGB colors, apply currency formatting, and insert a dynamic SUM total row.',
    tipTitle: 'See the connection',
    tipBody:
      'Range("A1:C1") refers to cells A1 through C1. Try changing the header color, then run your macro again.',
    selectedRange: { start: { row: 6, col: 3 }, end: { row: 6, col: 3 } },
    code: `Sub FormatSales()
    ' Give the report a little polish
    With Range("A1:C1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    Range("C2:C6").NumberFormat = "$#,##0"
    Range("A6").Value = "Total"
    Range("C6").Formula = "=SUM(C2:C4)"
    Range("A6:C6").Font.Bold = True
End Sub`,
    sheets: [
      baseSheet('Sheet1', {
        A1: { value: 'Product' },
        B1: { value: 'Units' },
        C1: { value: 'Revenue' },
        A2: { value: 'Notebook' },
        B2: { value: 24 },
        C2: { value: 240 },
        A3: { value: 'Desk lamp' },
        B3: { value: 18 },
        C3: { value: 540 },
        A4: { value: 'Planner' },
        B4: { value: 32 },
        C4: { value: 480 },
      }),
    ],
  },
  {
    id: 'highlight-low-stock',
    title: 'Highlight low inventory',
    category: 'Loops & Logic',
    difficulty: 'Beginner',
    description:
      'Loop over stock levels with For...Next and automatically flag items below reorder threshold in amber or crimson.',
    tipTitle: 'Loops meet conditional logic',
    tipBody:
      'Cells(r, 2).Value reads row r, column B dynamically inside the For loop so every row is inspected automatically.',
    selectedRange: { start: { row: 2, col: 4 }, end: { row: 6, col: 4 } },
    code: `Sub FlagLowStock()
    Dim r As Integer
    Range("D1").Value = "Status"
    With Range("A1:D1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    For r = 2 To 6
        If Cells(r, 2).Value < 15 Then
            Cells(r, 4).Value = "Reorder"
            Cells(r, 4).Interior.Color = RGB(254, 226, 226)
            Cells(r, 4).Font.Color = RGB(185, 28, 28)
            Cells(r, 4).Font.Bold = True
        Else
            Cells(r, 4).Value = "Healthy"
            Cells(r, 4).Interior.Color = RGB(220, 252, 231)
            Cells(r, 4).Font.Color = RGB(21, 128, 61)
        End If
    Next r
End Sub`,
    sheets: [
      baseSheet('Sheet1', {
        A1: { value: 'SKU Item' },
        B1: { value: 'In Stock' },
        C1: { value: 'Min Target' },
        A2: { value: 'USB-C Hub' },
        B2: { value: 8 },
        C2: { value: 15 },
        A3: { value: 'Keycap Set' },
        B3: { value: 42 },
        C3: { value: 15 },
        A4: { value: 'Monitor Arm' },
        B4: { value: 11 },
        C4: { value: 15 },
        A5: { value: 'Desk Mat' },
        B5: { value: 29 },
        C5: { value: 15 },
        A6: { value: 'Webcam' },
        B6: { value: 5 },
        C6: { value: 15 },
      }),
    ],
  },
  {
    id: 'commission-tiers',
    title: 'Calculate sales commissions',
    category: 'Business Logic',
    difficulty: 'Intermediate',
    description:
      'Compute tiered commission rates with ElseIf branches, format percentages, and print summary metrics with Debug.Print.',
    tipTitle: 'Debug.Print in action',
    tipBody:
      'Use Debug.Print to inspect intermediate totals in the Execution console without modifying any spreadsheet cells.',
    selectedRange: { start: { row: 2, col: 4 }, end: { row: 6, col: 4 } },
    code: `Sub CalculateCommissions()
    Dim r As Integer
    Dim sales As Double
    Dim rate As Double
    Dim totalPayout As Double

    With Range("A1:D1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    totalPayout = 0
    For r = 2 To 5
        sales = Cells(r, 2).Value
        If sales >= 15000 Then
            rate = 0.12
        ElseIf sales >= 8000 Then
            rate = 0.08
        Else
            rate = 0.05
        End If

        Cells(r, 3).Value = rate
        Cells(r, 3).NumberFormat = "0.0%"
        Cells(r, 4).Value = sales * rate
        Cells(r, 4).NumberFormat = "$#,##0"
        totalPayout = totalPayout + (sales * rate)
    Next r

    Range("B2:B5").NumberFormat = "$#,##0"
    Range("A7").Value = "Total Payout"
    Range("D7").Formula = "=SUM(D2:D5)"
    Range("D7").NumberFormat = "$#,##0"
    Range("A7:D7").Font.Bold = True

    Debug.Print "Total Commission Payout: $" & totalPayout
End Sub`,
    sheets: [
      baseSheet('Sheet1', {
        A1: { value: 'Rep Name' },
        B1: { value: 'Monthly Sales' },
        C1: { value: 'Tier Rate' },
        D1: { value: 'Commission' },
        A2: { value: 'Claire Lin' },
        B2: { value: 18400 },
        A3: { value: 'Marcus Chen' },
        B3: { value: 9600 },
        A4: { value: 'Sofia Patel' },
        B4: { value: 6200 },
        A5: { value: 'Daniel Wu' },
        B5: { value: 15200 },
      }),
    ],
  },
  {
    id: 'clean-customer-text',
    title: 'Clean messy text entries',
    category: 'Data Cleaning',
    difficulty: 'Beginner',
    description:
      'Trim extra spaces and standardize email and region casing across rows using Trim, LCase, and UCase.',
    tipTitle: 'Built-in string helpers',
    tipBody:
      'For Each cell In Range(...) iterates through every cell in a range cleanly—ideal for cleaning messy imported columns.',
    selectedRange: { start: { row: 2, col: 1 }, end: { row: 5, col: 3 } },
    code: `Sub CleanCustomerList()
    Dim cell As Range

    For Each cell In Range("A2:A5")
        cell.Value = Trim(cell.Value)
    Next cell

    For Each cell In Range("B2:B5")
        cell.Value = LCase(Trim(cell.Value))
    Next cell

    For Each cell In Range("C2:C5")
        cell.Value = UCase(Trim(cell.Value))
        cell.HorizontalAlignment = xlCenter
    Next cell

    With Range("A1:C1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub`,
    sheets: [
      baseSheet('Sheet1', {
        A1: { value: 'Customer' },
        B1: { value: 'Email' },
        C1: { value: 'Region' },
        A2: { value: '  Alice Wang  ' },
        B2: { value: '  Alice.Wang@Example.COM ' },
        C2: { value: ' tw ' },
        A3: { value: '  Ben Miller' },
        B3: { value: 'BEN@ACME.IO  ' },
        C3: { value: 'us' },
        A4: { value: 'Carla Gomez  ' },
        B4: { value: ' Carla.G@Studio.Net ' },
        C4: { value: ' eu ' },
        A5: { value: 'David Park' },
        B5: { value: '  DPARK@MAIL.COM' },
        C5: { value: 'jp' },
      }),
    ],
  },
  {
    id: 'dynamic-last-row',
    title: 'Find last row with End(xlUp)',
    category: 'Automation',
    difficulty: 'Intermediate',
    description:
      'Locate the bottom of a variable-length table with Cells(Rows.Count, 1).End(xlUp).Row and append a dynamic summary.',
    tipTitle: 'Why End(xlUp) matters',
    tipBody:
      'Instead of hardcoding row 6, End(xlUp) finds the true last populated row even when new data is added later.',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 8, col: 3 } },
    code: `Sub SummarizeToLastRow()
    Dim lastRow As Long
    Dim totalRow As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    totalRow = lastRow + 1

    Cells(totalRow, 1).Value = "Grand Total"
    Cells(totalRow, 2).Formula = "=SUM(B2:B" & lastRow & ")"
    Cells(totalRow, 3).Formula = "=AVERAGE(C2:C" & lastRow & ")"

    Range("B2:B" & totalRow).NumberFormat = "$#,##0"
    Range("C2:C" & totalRow).NumberFormat = "$#,##0.00"

    With Range("A" & totalRow & ":C" & totalRow)
        .Font.Bold = True
        .Interior.Color = RGB(232, 244, 237)
        .Font.Color = RGB(20, 83, 45)
    End With

    Debug.Print "Detected last data row: " & lastRow
End Sub`,
    sheets: [
      baseSheet('Sheet1', {
        A1: { value: 'Department' },
        B1: { value: 'Q1 Spend' },
        C1: { value: 'Avg Order' },
        A2: { value: 'Design' },
        B2: { value: 4200 },
        C2: { value: 350 },
        A3: { value: 'Engineering' },
        B3: { value: 12800 },
        C3: { value: 640 },
        A4: { value: 'Marketing' },
        B4: { value: 7600 },
        C4: { value: 475 },
        A5: { value: 'Operations' },
        B5: { value: 5100 },
        C5: { value: 255 },
      }),
    ],
  },
  {
    id: 'multi-sheet-rollup',
    title: 'Multi-sheet KPI rollup',
    category: 'Multi-Sheet',
    difficulty: 'Intermediate',
    description:
      'Read numbers from a second worksheet using Sheets("Orders") and WorksheetFunction.Sum, then build an executive summary.',
    tipTitle: 'Cross-sheet references',
    tipBody:
      'Switch to the Orders tab at the bottom of the workbook to see the source rows that Sheets("Orders") summarizes.',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 5, col: 2 } },
    code: `Sub BuildSummary()
    Dim orderTotal As Double
    Dim orderAvg As Double
    Dim orderCount As Long

    orderTotal = WorksheetFunction.Sum(Sheets("Orders").Range("B2:B6"))
    orderAvg = WorksheetFunction.Average(Sheets("Orders").Range("B2:B6"))
    orderCount = WorksheetFunction.Count(Sheets("Orders").Range("B2:B6"))

    With Sheets("Summary")
        .Range("A1:B1").Interior.Color = RGB(33, 115, 70)
        .Range("A1:B1").Font.Color = RGB(255, 255, 255)
        .Range("A1:B1").Font.Bold = True
        .Range("A2").Value = "Total Orders"
        .Range("B2").Value = orderCount
        .Range("A3").Value = "Total Revenue"
        .Range("B3").Value = orderTotal
        .Range("B3").NumberFormat = "$#,##0"
        .Range("A4").Value = "Average Order"
        .Range("B4").Value = orderAvg
        .Range("B4").NumberFormat = "$#,##0.00"
    End With

    MsgBox "Summary updated from " & orderCount & " orders!"
End Sub`,
    sheets: [
      baseSheet('Summary', {
        A1: { value: 'KPI Metric' },
        B1: { value: 'Value' },
      }),
      baseSheet('Orders', {
        A1: { value: 'Order ID' },
        B1: { value: 'Amount' },
        A2: { value: 'ORD-101' },
        B2: { value: 450 },
        A3: { value: 'ORD-102' },
        B3: { value: 1280 },
        A4: { value: 'ORD-103' },
        B4: { value: 890 },
        A5: { value: 'ORD-104' },
        B5: { value: 620 },
        A6: { value: 'ORD-105' },
        B6: { value: 1540 },
      }),
    ],
  },
];
