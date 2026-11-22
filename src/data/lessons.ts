import { LearnLesson } from '../types/workbook';

export const LEARN_LESSONS: LearnLesson[] = [
  {
    id: 'lesson-ranges',
    number: '01',
    title: '在空白表格中指定儲存格：Range 與 Cells',
    summary:
      '寫 VBA 的第一步，是告訴 Excel 要操作哪一格。面對空白工作表時，你可以用 Range("A1") 指定固定位址，或用 Cells(列, 欄) 以數字座標填入內容。',
    concept:
      'Range 接受標準的 A1 參照字串（如 "B2"、"A1:C1"），而 Cells(row, col) 則使用列號與欄號，非常適合搭配變數與迴圈自動填入資料。',
    code: `Sub WriteToBlankSheet()
    ' 在空白工作表寫入標題與第一筆資料
    Range("A1").Value = "Product"
    Range("B1").Value = "Units"
    Range("A1:B1").Font.Bold = True

    Cells(2, 1).Value = "Notebook"
    Cells(2, 2).Value = 24
End Sub`,
    keyTakeaways: [
      'Range("A1").Value 可讀取或寫入單一儲存格的內容。',
      'Range("A1:B1") 能一次將粗體或顏色套用到整個範圍。',
      'Cells(2, 1) 代表第 2 列、第 1 欄，也就是 A2 儲存格。',
    ],
    presetId: 'blank-sales-report',
  },
  {
    id: 'lesson-with',
    number: '02',
    title: '使用 With ... End With 整理空白表格樣式',
    summary:
      '當我們要替空白表格建立標題列並同時設定底色、文字顏色、粗體與置中對齊時，使用 With 區塊可以避免重複撰寫冗長的 Range("A1:C1")。',
    concept:
      '進入 With 區塊後，每一行只要以點號「.」開頭即可存取該物件的屬性，不僅程式碼更精簡，也更容易閱讀與維護。',
    code: `Sub StyleBlankHeader()
    Range("A1").Value = "Item"
    Range("B1").Value = "Qty"
    Range("C1").Value = "Amount"

    With Range("A1:C1")
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub`,
    keyTakeaways: [
      'RGB(紅, 綠, 藍) 各色版數值範圍為 0 到 255。',
      '每個 With 區塊結束時都必須搭配 End With 關閉。',
      '在編輯器中輸入「.」即可觸發常用屬性與指令提示。',
    ],
    presetId: 'blank-sales-report',
  },
  {
    id: 'lesson-loops',
    number: '03',
    title: '用 For ... Next 迴圈批次填滿空白工作表',
    summary:
      '只要幾行迴圈程式，就能在空白表格中自動產生連續編號、清單或計算結果，免去人工逐列複製貼上的時間。',
    concept:
      'For i = 起始值 To 結束值 會依序遞增變數 i，搭配 Cells(i, 欄) 就能從空白工作表快速長出整齊的資料列。',
    code: `Sub FillRowsFromBlank()
    Dim r As Integer
    Range("A1").Value = "No."
    Range("B1").Value = "Score"
    Range("C1").Value = "Result"
    Range("A1:C1").Font.Bold = True

    For r = 2 To 6
        Cells(r, 1).Value = r - 1
        Cells(r, 2).Value = (r - 1) * 20
        If Cells(r, 2).Value >= 60 Then
            Cells(r, 3).Value = "Pass"
        Else
            Cells(r, 3).Value = "Retry"
        End If
    Next r
End Sub`,
    keyTakeaways: [
      '使用 Dim r As Integer 明確宣告迴圈計數變數。',
      '在迴圈內搭配 If ... Then ... Else ... End If 可自動分類每列狀態。',
      '可先在空白工作表執行一次，觀察第 2 到第 6 列如何被自動填入。',
    ],
    presetId: 'blank-multiplication-grid',
  },
  {
    id: 'lesson-formulas',
    number: '04',
    title: '在空白表格寫入動態公式與數值格式',
    summary:
      '除了寫入固定數字外，透過 .Formula 屬性可以直接把 Excel 公式寫入空白儲存格，並用 .NumberFormat 設定貨幣或百分比顯示格式。',
    concept:
      '使用 .Formula = "=SUM(B2:B3)" 寫入的是活公式——之後只要透過 VBA 修改 B2 到 B3 的任一數值，總計欄位都會立即跟著重新計算。',
    code: `Sub BuildFormulaOnBlank()
    Range("A1").Value = "Item"
    Range("B1").Value = "Price"
    Range("A2").Value = "Keyboard"
    Range("B2").Value = 1200
    Range("A3").Value = "Mouse"
    Range("B3").Value = 650

    Range("A5").Value = "Total"
    Range("B5").Formula = "=SUM(B2:B3)"
    Range("B2:B5").NumberFormat = "$#,##0"
    Range("A5:B5").Font.Bold = True
End Sub`,
    keyTakeaways: [
      '.Value 寫入的是靜態數值；.Formula 寫入的則是可即時重算的公式。',
      '"$#,##0" 代表千分位貨幣格式，"0.0%" 代表一位小數百分比。',
      '寫入資料後再套用 NumberFormat，可確保版面整齊一致。',
    ],
    presetId: 'blank-sales-report',
  },
  {
    id: 'lesson-debug',
    number: '05',
    title: '使用 Debug.Print 與 MsgBox 進行除錯驗證',
    summary:
      '在撰寫複雜自動化巨集時，善用 Debug.Print 與 MsgBox 可以在不影響工作表版面的情況下，快速確認變數計算結果是否正確。',
    concept:
      'Debug.Print 會將訊息輸出至編輯器下方的訊息列；MsgBox 則會在執行結束時顯示提示訊息。',
    code: `Sub DebugOnBlankSheet()
    Dim total As Double
    Range("A1").Value = 150
    Range("A2").Value = 350

    total = WorksheetFunction.Sum(Range("A1:A2"))
    Debug.Print "計算加總結果：" & total
    MsgBox "空白表格計算完成，總計為 " & total
End Sub`,
    keyTakeaways: [
      '使用「&」運算子可將文字與數值變數串接成完整訊息。',
      '若語法有誤，編輯器行號列與下方錯誤提示列會標示確切錯誤行數。',
      '載入後按「Run macro」或 Ctrl / ⌘ + Enter 即可導入資料並檢視結果。',
    ],
    presetId: 'blank-sales-report',
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
