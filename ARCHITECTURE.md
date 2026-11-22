# 練練VBA — 技術架構文件 (Technical Architecture)

本文件說明 **練練VBA（Excel VBA Playground）** 的純前端靜態架構設計、核心模組職責、瀏覽器端 VBA 引擎執行流程、試算表公式運算引擎，以及 Light / Dark 雙主題與本地持久化策略。

---

## 1. 系統設計原則與架構邊界

- **純靜態前端單頁應用（Pure Static SPA）**：所有程式邏輯、語法檢查、VBA 直譯執行與試算表運算皆在瀏覽器端完成，**無任何後端伺服器、資料庫、外部 API 或 LLM / Gemini API 依賴**。
- **零環境相依與離線友善**：使用者無須安裝 Microsoft Excel 或本地編譯器，打包後（`dist/`）可直接部署至任何靜態託管平台（GitHub Pages、Cloudflare Pages、Netlify、Nginx、Vercel Static 等）。
- **資料不離機**：工作簿狀態、模組程式碼與介面主題僅存放於瀏覽器 `localStorage`，支援單機離線實驗與 CSV 匯出。

---

## 2. 技術棧（Tech Stack）

| 層級 | 選用技術 | 用途說明 |
| :--- | :--- | :--- |
| **建置工具** | Vite 8 + TypeScript 7 | ESM 開發伺服器與靜態資源打包輸出（`dist/`） |
| **UI 框架** | React 19 (Functional Components + Hooks) | 宣告式狀態管理與雙欄互動式 IDE 介面 |
| **樣式系統** | Tailwind CSS v4 (`@tailwindcss/vite`) | `@custom-variant dark` 驅動之 Light / Dark 雙主題樣式 |
| **圖示與字體** | `lucide-react`、`Plus Jakarta Sans`、`Noto Sans TC`、`JetBrains Mono` | 介面控制圖示、雙語排版與等寬程式字體 |
| **VBA 與公式核心** | 自研 TypeScript 詞法分析器、靜態 Linter、VBA 直譯器與公式求值器 | 純前端瀏覽器執行，無 `eval()` 或外部執行環境 |

---

## 3. 系統架構圖

```text
+-------------------------------------------------------------------------+
|                           Browser Runtime                               |
|                                                                         |
|  +------------------------+          +-------------------------------+  |
|  |   React UI Layer       |          |   Persistence & Theme         |  |
|  |  - TopBar & Modals     | <------> |  - localStorage (session)     |  |
|  |  - SpreadsheetPane     |          |  - localStorage (light/dark)  |  |
|  |  - VbaEditorPane       |          |  - html.dark class toggle     |  |
|  |  - Examples / Learn    |          +-------------------------------+  |
|  +-----------+------------+                                             |
|              |                                                          |
|              v                                                          |
|  +------------------------+          +-------------------------------+  |
|  |   VBA Language Core    |          |   Spreadsheet & Formula Core  |  |
|  |  1. Lexer (詞法掃描)   |          |  - Addressing (A1 / Range)    |  |
|  |  2. Linter (靜態除錯)  | -------> |  - NumberFormat 格式化        |  |
|  |  3. Suggestions (提示) |          |  - FormulaParser (SUM/IF/...) |  |
|  |  4. Interpreter (直譯) | =======> |  - 儲存格狀態更新與重算        |  |
|  +------------------------+          +-------------------------------+  |
+-------------------------------------------------------------------------+
```

---

## 4. 專案目錄與模組說明

```text
/
├── index.html                        # 網頁進入點、SEO Meta 與字體載入
├── metadata.json                     # 靜態應用描述資訊（無 API 權限需求）
├── package.json                      # 純前端依賴與 build/lint 腳本
├── vite.config.ts                    # Vite 建置與資源配置
├── ARCHITECTURE.md                   # 本技術架構文件
└── src/
    ├── main.tsx                      # React 掛載入口
    ├── index.css                     # Tailwind CSS v4 與 .dark 主題基底樣式
    ├── App.tsx                       # 全局狀態協調、快捷鍵、工作階段持久化
    ├── types/
    │   └── workbook.ts               # SheetData, CellData, VbaModule, ExecutionResult 型別
    ├── data/
    │   ├── examples.ts               # 內建範例巨集與對應初始工作表資料
    │   └── lessons.ts                # Learn VBA 教學單元與 Quick Guide 速查資料
    ├── spreadsheet/
    │   ├── addressing.ts             # A1/R1C1 座標轉換、範圍展開、NumberFormat、RGB 色彩轉換
    │   └── formulaEngine.ts          # 遞迴下降公式解析器（SUM, AVERAGE, IF, SUMIF 等）
    ├── vba/
    │   ├── lexer.ts                  # VBA 斷詞器（兼用於語法高亮與語法分析）
    │   ├── linter.ts                 # 即時靜態語法檢查（配對區塊、引號、括號、With 範圍）
    │   ├── suggestions.ts            # 語境感知指令自動補全與快速提示字典
    │   └── interpreter.ts            # 瀏覽器端 VBA 直譯執行引擎
    └── components/
        ├── TopBar.tsx                # 頂部導覽列、Quick guide 觸發與 Light/Dark 切換器
        ├── SpreadsheetPane.tsx       # 模擬 Excel Workbook、Ribbon 工具列、Formula Bar 與 Grid
        ├── VbaEditorPane.tsx         # 同步疊層高亮編輯器、指令提示彈窗與 Execution 面板
        ├── ExamplesView.tsx          # 範例瀏覽與一鍵載入執行頁面
        ├── LearnView.tsx             # 分步教學與程式碼練習頁面
        └── QuickGuideModal.tsx       # 快捷鍵與支援指令速查彈窗
```

---

## 5. 核心引擎運作機制

### 5.1 詞法分析與即時語法高亮（`src/vba/lexer.ts` + `VbaEditorPane.tsx`）
- **單一詞法來源**：編輯器視覺高亮與後端靜態檢查共用 `tokenizeVbaLine()`，確保看到的顏色分類與實際解析語意一致。
- **同步雙層編輯架構**：底層為標記化 `<pre>` 高亮圖層，上層為透明 `<textarea>`，透過同步 `scrollTop / scrollLeft` 維持原生的選取、複製貼上、中文輸入法（IME）及 Undo/Redo 體驗，避免傳統 `contenteditable` 造成的游標偏移問題。

### 5.2 即時語法除錯（`src/vba/linter.ts`）
在執行巨集前與編輯當下進行靜態掃描：
- 區塊配對驗證：`Sub / End Sub`、`Function / End Function`、`With / End With`、`If ... Then / End If`、`For / Next`、`Do / Loop`、`Select Case / End Select`。
- 行內檢查：未閉合字串引號 `"`、不平衡括號 `()`、在 `With` 區塊外使用 `.Property` 前綴。
- 輸出包含行號與明確錯誤訊息的 `VbaDiagnostic[]`，供行號欄紅點與執行區跳轉使用。

### 5.3 常見指令提示與自動補全（`src/vba/suggestions.ts`）
- **語境判斷**：偵測游標前是否剛輸入 `.`（觸發 `VBA_DOT_PROPERTIES`，如 `.Interior.Color`、`.Font.Bold`、`.NumberFormat`、`.Offset()`）或一般關鍵字前綴（觸發 `VBA_GLOBAL_SUGGESTIONS`，如 `Range(`、`Cells(`、`With ... End With`、`For ... Next`、`Debug.Print`）。
- **快速提示條（Hints）**：編輯器底部提供常用指令膠囊，支援點擊直接插入游標位置。

### 5.4 VBA 直譯執行引擎（`src/vba/interpreter.ts`）
1. **前處理**：先呼叫 `lintVbaCode()` 攔截語法錯誤，再建立 `ProcedureDef` 索引表（支援主程序呼叫輔助 `Sub` / `Function`）。
2. **安全執行防護**：內建步數上限保護（`MAX_STEPS = 50,000`），避免撰寫 `Do ... Loop` 或 `For` 迴圈時造成瀏覽器分頁卡死。
3. **Excel 物件模型模擬**：
   - `Range("A1:C6")`、`Cells(row, col)`、`Sheets("Name")`、`ActiveSheet`、`ActiveCell`、`Selection`、`Rows`、`Columns`。
   - 鏈式屬性與方法：`.Offset(r, c)`、`.Resize(r, c)`、`.End(xlUp | xlDown | xlToLeft | xlToRight)`、`.Clear`、`.ClearContents`、`.Select`。
   - 樣式寫入：`.Interior.Color = RGB(r, g, b)`、`.Font.Color`、`.Font.Bold / Italic / Underline / Size / Name`、`.NumberFormat`、`.HorizontalAlignment`、`.WrapText`、`.Borders`。
   - 內建函式與偵錯：`WorksheetFunction.Sum/Average/Count/Min/Max/Round`、`RGB()`、字串/數學轉換函式、`Debug.Print`、`MsgBox`。
4. **不可變狀態提交**：執行前複製工作表狀態，完成後觸發 `recalculateSheets()` 重算所有儲存格公式並回傳更動計數、耗時（ms）與輸出紀錄。

### 5.5 試算表公式與格式化引擎（`src/spreadsheet/`）
- `formulaEngine.ts` 以遞迴下降解析器處理 `=` 開頭的運算式，支援儲存格參照（`C2:C4`、跨工作表 `Orders!B2:B6`）、四則與比較運算，以及 `SUM`、`AVERAGE`、`COUNT`、`COUNTA`、`MIN`、`MAX`、`IF`、`SUMIF`、`COUNTIF`、`ROUND` 等常用函式，並以 `visited` 集合防止循環參照造成無窮遞迴。
- `addressing.ts` 解析 Excel `NumberFormat`（如 `$#,##0`、`0.0%`、`#,##0.00`）並將 VBA `RGB(r, g, b)` 24-bit 整數（`r + g*256 + b*65536`）轉為標準 CSS Hex 色碼。

---

## 6. Light / Dark 雙主題架構

- `src/index.css` 於 `@import "tailwindcss";` 下方宣告 `@custom-variant dark (&:where(.dark, .dark *));`，由 `App.tsx` 同步切換 `<html>` 的 `.dark` class 與 `localStorage['vba-playground-theme']`。
- **程式碼編輯區對比設計**：為保持語法高亮在兩種模式下皆具備高辨識度，編輯器畫布採深森林石板色調（Light/Dark 一致），其餘導覽列、Ribbon 工具列、工作表網格、執行狀態卡與提示橫幅皆成對配置高對比 Light / Dark 色票。
- **儲存格背景可讀性保護**：若巨集為儲存格指定淺色底色（如 `RGB(220, 252, 231)`）而未指定文字顏色，Dark Mode 下會依亮度自動補上深色前景文字，避免淺底白字失真。

---

## 7. 本地開發與靜態建置指令

```bash
# 安裝依賴
npm install

# 啟動本地開發伺服器 (Port 3000)
npm run dev

# 型別檢查
npm run lint

# 產生純靜態生產檔案 (輸出至 dist/)
npm run build
```

執行 `npm run build` 後，可直接將 `dist/` 目錄部署至任意靜態網站空間，無須配置環境變數（`.env`）或執行任何伺服器進程。
