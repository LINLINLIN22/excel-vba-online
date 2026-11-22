import { ExamplePreset, SheetData } from '../types/workbook';

export function createBlankSheet(
  name = 'Sheet1',
  rows = 15,
  cols = 7
): SheetData {
  return {
    id: `sheet-${name.toLowerCase().replace(/\s+/g, '-')}`,
    name,
    rows,
    cols,
    cells: {},
    colWidths: {
      1: 102,
      2: 76,
      3: 84,
      4: 74,
      5: 74,
      6: 74,
      7: 74,
    },
  };
}

function baseSheet(
  name: string,
  cells: SheetData['cells'],
  rows = 15,
  cols = 7
): SheetData {
  return {
    ...createBlankSheet(name, rows, cols),
    cells,
  };
}

export const EXAMPLE_PRESETS: ExamplePreset[] = [
  {
    id: 'blank-sales-report',
    title: '從空白工作表建立報表',
    category: '空白起手',
    difficulty: '初階',
    description:
      '從完全空白的 Excel 表格出發，用 VBA 填入商品銷售資料、套用標題樣式並加入自動加總公式。',
    tipTitle: '從空白表格開始認識 VBA',
    tipBody:
      '左邊目前是一張空白的 Excel 工作表。按下「Run macro」後，VBA 會依序寫入欄位內容、套用 RGB 色彩並建立 =SUM(C2:C4) 加總公式。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub BuildSalesReport()
    ' 從空白工作表填入標題與資料
    Range("A1").Value = "Product"
    Range("B1").Value = "Units"
    Range("C1").Value = "Revenue"

    Range("A2").Value = "Notebook"
    Range("B2").Value = 24
    Range("C2").Value = 240

    Range("A3").Value = "Desk lamp"
    Range("B3").Value = 18
    Range("C3").Value = 540

    Range("A4").Value = "Planner"
    Range("B4").Value = 32
    Range("C4").Value = 480

    ' 美化標題列樣式
    With Range("A1:C1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 套用貨幣格式與加總公式
    Range("C2:C6").NumberFormat = "$#,##0"
    Range("A6").Value = "Total"
    Range("C6").Formula = "=SUM(C2:C4)"
    Range("A6:C6").Font.Bold = True
End Sub`,
    sheets: [createBlankSheet('Sheet1')],
  },
  {
    id: 'blank-multiplication-grid',
    title: '在空白表格產生九九乘法表',
    category: '迴圈與邏輯',
    difficulty: '初階',
    description:
      '利用雙層 For...Next 迴圈，直接將九九乘法表自動填入空白工作表並標記標題列與欄。',
    tipTitle: '巢狀迴圈搭配 Cells(列, 欄)',
    tipBody:
      '在空白表格中使用 Cells(r, c) 可以用數字變數動態指定列與欄，非常適合批次產生表格或矩陣資料。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub GenerateTimesTable()
    Dim r As Integer
    Dim c As Integer

    Cells(1, 1).Value = "×"
    With Range("A1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    For r = 1 To 6
        Cells(r + 1, 1).Value = r
        Cells(r + 1, 1).Font.Bold = True
        Cells(r + 1, 1).HorizontalAlignment = xlCenter

        Cells(1, r + 1).Value = r
        Cells(1, r + 1).Font.Bold = True
        Cells(1, r + 1).HorizontalAlignment = xlCenter

        For c = 1 To 6
            Cells(r + 1, c + 1).Value = r * c
            Cells(r + 1, c + 1).HorizontalAlignment = xlCenter
        Next c
    Next r
End Sub`,
    sheets: [createBlankSheet('Sheet1')],
  },
  {
    id: 'highlight-low-stock',
    title: '低庫存自動標記警示',
    category: '迴圈與邏輯',
    difficulty: '初階',
    description:
      '使用 For...Next 逐列掃描庫存數量，低於安全庫存時自動填入「Reorder」並以紅底標示警示。',
    tipTitle: '迴圈結合條件判斷',
    tipBody:
      '透過 Cells(r, 2).Value 動態讀取第 r 列、第 B 欄數值，再用 If...Else 為每列填入對應狀態與顏色。',
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
    title: '業務階梯獎金試算',
    category: '商業計算',
    difficulty: '中階',
    description:
      '以 ElseIf 判斷各級業績對應抽成比例，套用百分比與貨幣格式，並透過 Debug.Print 輸出總獎金。',
    tipTitle: 'Debug.Print 即時除錯輸出',
    tipBody:
      '利用 Debug.Print 可將計算過程中的變數數值印出到右側 Execution 區域，無須改動儲存格即可檢查結果。',
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
    title: '批次清理文字與大小寫',
    category: '資料清理',
    difficulty: '初階',
    description:
      '使用 For Each 搭配 Trim、LCase、UCase 快速清除多餘空白並統一 Email 與地區代碼大小寫格式。',
    tipTitle: '字串清理實用函式',
    tipBody:
      'For Each cell In Range(...) 會逐一走訪指定範圍內的每個儲存格，特別適合清洗匯入後的雜亂文字資料。',
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
    id: 'multi-sheet-rollup',
    title: '跨工作表 KPI 彙總',
    category: '自動化',
    difficulty: '中階',
    description:
      '透過 Sheets("Orders") 與 WorksheetFunction 跨頁讀取訂單金額，在 Summary 分頁自動產出關鍵指標。',
    tipTitle: '跨工作表參照技巧',
    tipBody:
      '切換下方的 Orders 工作表分頁，即可檢視 Sheets("Orders") 實際讀取並彙總的原始訂單明細。',
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
