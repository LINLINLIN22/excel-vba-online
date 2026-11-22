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

export const EXAMPLE_PRESETS: ExamplePreset[] = [
  {
    id: 'blank-sales-report',
    title: '從空白工作表建立報表',
    category: '空白起手',
    difficulty: '初階',
    description:
      '載入時維持空白 Excel 表格，按下「Run macro」後由 VBA 自動匯入商品銷售資料、套用標題樣式並建立加總公式。',
    tipTitle: '從空白表格開始認識 VBA',
    tipBody:
      '左邊目前是一張空白的 Excel 工作表。按下「Run macro」後，VBA 會依序寫入欄位內容、套用 RGB 色彩並建立 =SUM(C2:C4) 加總公式。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub BuildSalesReport()
    ' 導入標題與商品銷售資料
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
      '利用雙層 For...Next 迴圈，按下執行後直接將九九乘法表填入空白工作表並標記標題列與欄。',
    tipTitle: '巢狀迴圈搭配 Cells(列, 欄)',
    tipBody:
      '先保持左側工作表空白，按下「Run macro」後程式會透過 Cells(r, c) 動態填入乘積結果。',
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
      '按下「Run macro」後先導入庫存清單，再使用 For...Next 逐列掃描數量，低於安全庫存時自動標示「Reorder」。',
    tipTitle: '迴圈結合條件判斷',
    tipBody:
      '按下「Run macro」導入資料後，程式會透過 Cells(r, 2).Value 讀取庫存數，並用 If...Else 自動填入狀態與紅綠標記。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub FlagLowStock()
    Dim r As Integer

    ' 1. 導入庫存初始資料至空白工作表
    Range("A1").Value = "SKU Item": Range("B1").Value = "In Stock": Range("C1").Value = "Min Target": Range("D1").Value = "Status"
    Range("A2").Value = "USB-C Hub": Range("B2").Value = 8: Range("C2").Value = 15
    Range("A3").Value = "Keycap Set": Range("B3").Value = 42: Range("C3").Value = 15
    Range("A4").Value = "Monitor Arm": Range("B4").Value = 11: Range("C4").Value = 15
    Range("A5").Value = "Desk Mat": Range("B5").Value = 29: Range("C5").Value = 15
    Range("A6").Value = "Webcam": Range("B6").Value = 5: Range("C6").Value = 15

    With Range("A1:D1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 2. 掃描庫存並自動標記警示
    For r = 2 To 6
        If Cells(r, 2).Value < Cells(r, 3).Value Then
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
    sheets: [createBlankSheet('Sheet1')],
  },
  {
    id: 'commission-tiers',
    title: '業務階梯獎金試算',
    category: '商業計算',
    difficulty: '中階',
    description:
      '按下「Run macro」後匯入業務員業績，以 ElseIf 判斷各級抽成比例並套用格式，同時透過 Debug.Print 輸出總獎金。',
    tipTitle: 'Debug.Print 即時除錯輸出',
    tipBody:
      '按下「Run macro」後除了填入業績與獎金報表，Debug.Print 還會把總獎金直接顯示在編輯器下方的訊息列。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub CalculateCommissions()
    Dim r As Integer
    Dim sales As Double
    Dim rate As Double
    Dim totalPayout As Double

    ' 1. 導入業務業績表
    Range("A1").Value = "Rep Name": Range("B1").Value = "Monthly Sales": Range("C1").Value = "Tier Rate": Range("D1").Value = "Commission"
    Range("A2").Value = "Claire Lin": Range("B2").Value = 18400
    Range("A3").Value = "Marcus Chen": Range("B3").Value = 9600
    Range("A4").Value = "Sofia Patel": Range("B4").Value = 6200
    Range("A5").Value = "Daniel Wu": Range("B5").Value = 15200

    With Range("A1:D1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 2. 計算階梯抽成與總獎金
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
    sheets: [createBlankSheet('Sheet1')],
  },
  {
    id: 'clean-customer-text',
    title: '批次清理文字與大小寫',
    category: '資料清理',
    difficulty: '初階',
    description:
      '按下「Run macro」後導入含多餘空白與大小寫不一的原始客戶名單，再以 For Each 搭配 Trim、LCase、UCase 自動清洗。',
    tipTitle: '字串清理實用函式',
    tipBody:
      '點擊「Run macro」後，VBA 會先寫入待整理的原始字串，再透過 For Each cell In Range(...) 逐格去除空白並統一大小寫。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub CleanCustomerList()
    Dim cell As Range

    ' 1. 導入待清理的原始客戶資料
    Range("A1").Value = "Customer": Range("B1").Value = "Email": Range("C1").Value = "Region"
    Range("A2").Value = "  Alice Wang  ": Range("B2").Value = "  Alice.Wang@Example.COM ": Range("C2").Value = " tw "
    Range("A3").Value = "  Ben Miller": Range("B3").Value = "BEN@ACME.IO  ": Range("C3").Value = "us"
    Range("A4").Value = "Carla Gomez  ": Range("B4").Value = " Carla.G@Studio.Net ": Range("C4").Value = " eu "
    Range("A5").Value = "David Park": Range("B5").Value = "  DPARK@MAIL.COM": Range("C5").Value = "jp"

    ' 2. 批次去除前後空白並統一格式
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
    sheets: [createBlankSheet('Sheet1')],
  },
  {
    id: 'multi-sheet-rollup',
    title: '跨工作表 KPI 彙總',
    category: '自動化',
    difficulty: '中階',
    description:
      '從單一空白工作表出發，按下「Run macro」後自動建立 Summary 與 Orders 雙分頁並跨表統計訂單關鍵指標。',
    tipTitle: '動態建立工作表與跨頁參照',
    tipBody:
      '按下「Run macro」執行後，可切換下方的 Orders 與 Summary 工作表分頁，檢視動態產生的明細與彙總結果。',
    selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    code: `Sub BuildSummary()
    Dim orderTotal As Double
    Dim orderAvg As Double
    Dim orderCount As Long

    ' 1. 將空白工作表命名為 Summary，並新增 Orders 分頁導入訂單資料
    ActiveSheet.Name = "Summary"
    Sheets.Add.Name = "Orders"

    With Sheets("Orders")
        .Range("A1").Value = "Order ID": .Range("B1").Value = "Amount"
        .Range("A2").Value = "ORD-101": .Range("B2").Value = 450
        .Range("A3").Value = "ORD-102": .Range("B3").Value = 1280
        .Range("A4").Value = "ORD-103": .Range("B4").Value = 890
        .Range("A5").Value = "ORD-104": .Range("B5").Value = 620
        .Range("A6").Value = "ORD-105": .Range("B6").Value = 1540
    End With

    ' 2. 跨工作表統計訂單指標並寫入 Summary
    orderTotal = WorksheetFunction.Sum(Sheets("Orders").Range("B2:B6"))
    orderAvg = WorksheetFunction.Average(Sheets("Orders").Range("B2:B6"))
    orderCount = WorksheetFunction.Count(Sheets("Orders").Range("B2:B6"))

    With Sheets("Summary")
        .Range("A1").Value = "KPI Metric"
        .Range("B1").Value = "Value"
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

    Sheets("Summary").Activate
    MsgBox "Summary updated from " & orderCount & " orders!"
End Sub`,
    sheets: [createBlankSheet('Sheet1')],
  },
];
