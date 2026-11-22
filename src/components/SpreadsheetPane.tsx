import {
  ChevronLeft,
  ChevronRight,
  Redo2,
  Undo2,
} from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  colIndexToLetter,
  coordToAddress,
  formatCellDisplayValue,
  isCoordInRange,
  iterateRangeCoords,
  normalizeRange,
  rangeToAddress,
} from '../spreadsheet/addressing';
import {
  CellCoord,
  CellRange,
  HorizontalAlign,
  SheetData,
} from '../types/workbook';

interface SpreadsheetPaneProps {
  sheets: SheetData[];
  activeSheetId: string;
  selectedRange: CellRange;
  onSelectSheet: (sheetId: string) => void;
  onSelectRange: (range: CellRange) => void;
  canUndo: boolean;
  canRedo: boolean;
  historyIndex: number;
  historyLength: number;
  onUndo: () => void;
  onRedo: () => void;
}

export function SpreadsheetPane({
  sheets,
  activeSheetId,
  selectedRange,
  onSelectSheet,
  onSelectRange,
  canUndo,
  canRedo,
  historyIndex,
  historyLength,
  onUndo,
  onRedo,
}: SpreadsheetPaneProps) {
  const activeSheet = useMemo(
    () => sheets.find((s) => s.id === activeSheetId) ?? sheets[0],
    [sheets, activeSheetId]
  );

  const [isDragging, setIsDragging] = useState(false);
  const dragAnchorRef = useRef<CellCoord | null>(null);
  const tabsScrollRef = useRef<HTMLDivElement | null>(null);

  const activeCoord = selectedRange.start;
  const activeAddr = coordToAddress(activeCoord);
  const activeCell = activeSheet.cells[activeAddr];
  const rangeLabel = rangeToAddress(selectedRange);

  const formulaDisplay =
    activeCell?.formula ??
    (activeCell?.value !== null && activeCell?.value !== undefined
      ? String(activeCell.value)
      : '');

  useEffect(() => {
    const handleMouseUp = () => setIsDragging(false);
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  const displayRows = Math.max(13, activeSheet.rows);
  const displayCols = Math.max(7, activeSheet.cols);

  // Compute live Sum for status bar when inspecting cells
  const statusSummary = useMemo(() => {
    const coords = iterateRangeCoords(selectedRange);
    const nums: number[] = [];
    for (const c of coords) {
      const cell = activeSheet.cells[coordToAddress(c)];
      if (cell && cell.value !== null && cell.value !== '' && typeof cell.value !== 'boolean') {
        const n = Number(cell.value);
        if (Number.isFinite(n)) nums.push(n);
      }
    }
    if (nums.length === 0) return { count: coords.length, sum: null as number | null };
    const sum = nums.reduce((acc, n) => acc + n, 0);
    return { count: nums.length, sum };
  }, [activeSheet.cells, selectedRange]);

  const handleCellMouseDown = (coord: CellCoord, e: React.MouseEvent) => {
    if (e.shiftKey) {
      onSelectRange(normalizeRange({ start: selectedRange.start, end: coord }));
      return;
    }
    dragAnchorRef.current = coord;
    setIsDragging(true);
    onSelectRange({ start: coord, end: coord });
  };

  const handleCellMouseEnter = (coord: CellCoord) => {
    if (!isDragging || !dragAnchorRef.current) return;
    onSelectRange(normalizeRange({ start: dragAnchorRef.current, end: coord }));
  };

  const colLetters = useMemo(
    () => Array.from({ length: displayCols }, (_, idx) => colIndexToLetter(idx + 1)),
    [displayCols]
  );

  const totalSteps = Math.max(0, historyLength - 1);

  return (
    <div className="h-full flex flex-col rounded-xl border border-[#DCE5DF] dark:border-[#233A2C] bg-white dark:bg-[#14241B] shadow-xs overflow-hidden">
      {/* Window Title Bar */}
      <div className="flex items-center justify-between gap-3 px-4 py-1.5 border-b border-[#E7EEE9] dark:border-[#21362A]">
        <div className="flex items-center gap-2.5 min-w-0">
          <div
            aria-hidden="true"
            className="w-5 h-5 rounded bg-[#1D6F42] text-white font-bold text-[11px] flex items-center justify-center shrink-0"
          >
            X
          </div>
          <span className="text-xs font-semibold text-[#111C16] dark:text-[#E8F1EC] whitespace-nowrap">
            Workbook1
          </span>
          <span className="text-[11px] text-[#73847A] dark:text-[#7E9689] truncate">
            由 VBA 控制 · 唯讀預覽
          </span>
        </div>

        <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full bg-[#EDF5F0] dark:bg-[#1B3325] text-[#157347] dark:text-[#65CE91] text-[11px] font-medium whitespace-nowrap shrink-0">
          <span className="w-1.5 h-1.5 rounded-full bg-[#18804C] dark:bg-[#4ADE80]" />
          <span>Live preview</span>
        </div>
      </div>

      {/* VBA History Toolbar (Only Undo & Redo) */}
      <div className="flex items-center justify-between gap-3 px-3.5 py-1.5 border-b border-[#E7EEE9] dark:border-[#21362A] bg-[#FAFDFB] dark:bg-[#17291F]">
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={onUndo}
            disabled={!canUndo}
            title="回到上一個 VBA 執行狀態"
            className={`h-7 px-2.5 rounded-md border text-xs font-medium inline-flex items-center gap-1.5 transition-colors whitespace-nowrap ${
              canUndo
                ? 'border-[#D4DFD8] dark:border-[#294234] bg-white dark:bg-[#122018] text-[#111C16] dark:text-[#E8F1EC] hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] cursor-pointer'
                : 'border-[#E5ECE8] dark:border-[#1E3327] bg-[#F5F8F6] dark:bg-[#14231B] text-[#9EB0A6] dark:text-[#4F6659] cursor-not-allowed'
            }`}
          >
            <Undo2 className="w-3.5 h-3.5" />
            <span>上一步</span>
          </button>

          <button
            type="button"
            onClick={onRedo}
            disabled={!canRedo}
            title="恢復下一個 VBA 執行狀態"
            className={`h-7 px-2.5 rounded-md border text-xs font-medium inline-flex items-center gap-1.5 transition-colors whitespace-nowrap ${
              canRedo
                ? 'border-[#D4DFD8] dark:border-[#294234] bg-white dark:bg-[#122018] text-[#111C16] dark:text-[#E8F1EC] hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] cursor-pointer'
                : 'border-[#E5ECE8] dark:border-[#1E3327] bg-[#F5F8F6] dark:bg-[#14231B] text-[#9EB0A6] dark:text-[#4F6659] cursor-not-allowed'
            }`}
          >
            <Redo2 className="w-3.5 h-3.5" />
            <span>下一步</span>
          </button>
        </div>

        <span className="text-[11px] text-[#6B7E73] dark:text-[#849B8E] tabular-nums whitespace-nowrap">
          {totalSteps === 0
            ? '空白工作表（請按下 Run macro）'
            : `VBA 紀錄：第 ${historyIndex} / ${totalSteps} 步`}
        </span>
      </div>

      {/* Read-only Formula Inspector Bar */}
      <div className="flex items-center gap-2 px-3 py-1.5 border-b border-[#E7EEE9] dark:border-[#21362A] bg-white dark:bg-[#132219]">
        <div className="flex items-center justify-center min-w-[54px] px-2 py-1 rounded border border-[#E2EAE5] dark:border-[#243B2E] text-xs font-medium text-[#1D2D24] dark:text-[#D5E3DB] tabular-nums shrink-0">
          {rangeLabel}
        </div>

        <span
          aria-hidden="true"
          className="px-2 text-sm italic font-serif text-[#526359] dark:text-[#8FA699] select-none border-r border-[#E2EAE5] dark:border-[#243B2E]"
        >
          fx
        </span>

        <div
          aria-label="Cell content preview"
          className="flex-1 text-xs font-mono text-[#111C16] dark:text-[#E8F1EC] truncate py-1"
        >
          {formulaDisplay ? (
            <span>{formulaDisplay}</span>
          ) : (
            <span className="text-[#8DA094] dark:text-[#5B7466] font-sans">
              僅供檢視：請於右側編輯 VBA 並按下 Run macro 執行
            </span>
          )}
        </div>
      </div>

      {/* Read-only Spreadsheet Grid */}
      <div className="flex-1 min-h-0 overflow-auto vba-scroll bg-white dark:bg-[#122018] select-none">
        <table className="border-collapse table-fixed w-full text-[12px]">
          <thead>
            <tr>
              <th className="w-10 min-w-[40px] h-6 bg-[#F5F8F6] dark:bg-[#192B20] border-b border-r border-[#E2EAE5] dark:border-[#243A2D] sticky top-0 left-0 z-20" />
              {colLetters.map((letter, idx) => {
                const colNum = idx + 1;
                const isColSelected =
                  colNum >= selectedRange.start.col && colNum <= selectedRange.end.col;
                const widthPx = activeSheet.colWidths[colNum] ?? 96;
                return (
                  <th
                    key={letter}
                    style={{ width: `${widthPx}px` }}
                    onClick={() =>
                      onSelectRange({
                        start: { row: 1, col: colNum },
                        end: { row: displayRows, col: colNum },
                      })
                    }
                    className={`h-6 text-[11px] font-medium border-b border-r border-[#E2EAE5] dark:border-[#243A2D] sticky top-0 z-10 cursor-pointer transition-colors ${
                      isColSelected
                        ? 'bg-[#E5F2EA] dark:bg-[#223B2C] text-[#157347] dark:text-[#6EE7A0]'
                        : 'bg-[#F5F8F6] dark:bg-[#192B20] text-[#5F7267] dark:text-[#8CA396]'
                    }`}
                  >
                    {letter}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: displayRows }, (_, rIdx) => {
              const rowNum = rIdx + 1;
              const isRowSelected =
                rowNum >= selectedRange.start.row && rowNum <= selectedRange.end.row;

              return (
                <tr key={rowNum} className="h-[26px]">
                  {/* Row Number Header */}
                  <td
                    onClick={() =>
                      onSelectRange({
                        start: { row: rowNum, col: 1 },
                        end: { row: rowNum, col: displayCols },
                      })
                    }
                    className={`text-center text-[11px] tabular-nums border-b border-r border-[#E2EAE5] dark:border-[#243A2D] sticky left-0 z-10 cursor-pointer transition-colors ${
                      isRowSelected
                        ? 'bg-[#E5F2EA] dark:bg-[#223B2C] text-[#157347] dark:text-[#6EE7A0] font-semibold'
                        : 'bg-[#F5F8F6] dark:bg-[#192B20] text-[#5F7267] dark:text-[#8CA396]'
                    }`}
                  >
                    {rowNum}
                  </td>

                  {colLetters.map((_, cIdx) => {
                    const colNum = cIdx + 1;
                    const coord = { row: rowNum, col: colNum };
                    const addr = coordToAddress(coord);
                    const cell = activeSheet.cells[addr];
                    const style = cell?.style ?? {};
                    const isActive =
                      activeCoord.row === rowNum && activeCoord.col === colNum;
                    const inSelection = isCoordInRange(coord, selectedRange);

                    const rawDisplay = formatCellDisplayValue(
                      cell?.value,
                      style.numberFormat
                    );

                    const isNumeric =
                      typeof cell?.value === 'number' ||
                      (typeof cell?.value === 'string' &&
                        cell.value.trim() !== '' &&
                        Number.isFinite(Number(cell.value)));

                    const align: HorizontalAlign =
                      style.horizontalAlign ?? (isNumeric ? 'right' : 'left');

                    const hasCustomBg = Boolean(style.backgroundColor);
                    const customTextColor =
                      style.color ??
                      (hasCustomBg && isLightHex(style.backgroundColor!)
                        ? '#111C16'
                        : undefined);

                    return (
                      <td
                        key={addr}
                        onMouseDown={(e) => handleCellMouseDown(coord, e)}
                        onMouseEnter={() => handleCellMouseEnter(coord)}
                        style={{
                          backgroundColor: style.backgroundColor
                            ? style.backgroundColor
                            : inSelection && !isActive
                              ? 'rgba(33, 115, 70, 0.08)'
                              : undefined,
                          color: customTextColor,
                          fontWeight: style.bold ? 600 : undefined,
                          fontStyle: style.italic ? 'italic' : undefined,
                          textDecoration: style.underline ? 'underline' : undefined,
                          fontFamily: style.fontFamily,
                          fontSize: style.fontSize ? `${style.fontSize}px` : undefined,
                          textAlign: align,
                          whiteSpace: style.wrapText ? 'normal' : 'nowrap',
                          borderColor:
                            style.borderStyle === 'thin'
                              ? style.borderColor ?? '#94A3B8'
                              : undefined,
                        }}
                        className={`relative px-2 py-1 tabular-nums overflow-hidden text-ellipsis cursor-default transition-colors border-b border-r border-[#E6EDE9] dark:border-[#1E3227] ${
                          !customTextColor
                            ? 'text-[#18261E] dark:text-[#D5E3DB]'
                            : ''
                        }`}
                      >
                        <span>{rawDisplay}</span>

                        {/* Excel active selection box */}
                        {isActive && (
                          <div
                            aria-hidden="true"
                            className="pointer-events-none absolute inset-0 border-2 border-[#187E4B] dark:border-[#39A96B] z-10"
                          >
                            <span className="absolute -bottom-[3px] -right-[3px] w-[6px] h-[6px] bg-[#187E4B] dark:bg-[#39A96B] border border-white dark:border-[#122018]" />
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Sheet Tabs Bar (Switch sheets created by VBA; no manual add) */}
      <div className="flex items-center justify-between px-3 py-1.5 border-t border-[#E7EEE9] dark:border-[#21362A] bg-[#FAFDFB] dark:bg-[#15251C]">
        <div className="flex items-center gap-1 min-w-0">
          <button
            type="button"
            onClick={() => tabsScrollRef.current?.scrollBy({ left: -80, behavior: 'smooth' })}
            aria-label="Scroll sheets left"
            className="p-1 rounded text-[#73847A] hover:text-[#111C16] dark:hover:text-[#E8F1EC] hover:bg-[#EAF2ED] dark:hover:bg-[#1E3327]"
          >
            <ChevronLeft className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => tabsScrollRef.current?.scrollBy({ left: 80, behavior: 'smooth' })}
            aria-label="Scroll sheets right"
            className="p-1 rounded text-[#73847A] hover:text-[#111C16] dark:hover:text-[#E8F1EC] hover:bg-[#EAF2ED] dark:hover:bg-[#1E3327]"
          >
            <ChevronRight className="w-3.5 h-3.5" />
          </button>

          <div
            ref={tabsScrollRef}
            className="flex items-center gap-1 overflow-x-auto vba-scroll ml-1"
          >
            {sheets.map((sh) => {
              const isShActive = sh.id === activeSheet.id;
              return (
                <button
                  key={sh.id}
                  type="button"
                  onClick={() => onSelectSheet(sh.id)}
                  className={`relative px-3 py-1 text-xs font-medium rounded-t whitespace-nowrap shrink-0 transition-colors ${
                    isShActive
                      ? 'bg-white dark:bg-[#122018] text-[#157347] dark:text-[#6EE7A0]'
                      : 'text-[#526359] dark:text-[#8FA699] hover:text-[#111C16] dark:hover:text-[#E8F1EC]'
                  }`}
                >
                  {sh.name}
                  {isShActive && (
                    <span className="absolute bottom-0 left-1 right-1 h-[2px] bg-[#157347] dark:bg-[#39A96B] rounded" />
                  )}
                </button>
              );
            })}
          </div>
        </div>
      </div>

      {/* Workbook Bottom Status Bar */}
      <div className="flex items-center justify-between gap-4 px-4 py-1.5 border-t border-[#E7EEE9] dark:border-[#21362A] bg-[#F6F9F7] dark:bg-[#111F17] text-[11px] text-[#5F7267] dark:text-[#8CA396]">
        <span>Ready</span>

        <div className="flex items-center gap-4 tabular-nums">
          {statusSummary.sum !== null && (
            <span>
              Sum:{' '}
              <strong className="font-medium text-[#18261E] dark:text-[#D5E3DB]">
                ${statusSummary.sum.toLocaleString('en-US', { maximumFractionDigits: 2 })}
              </strong>
            </span>
          )}
        </div>
      </div>
    </div>
  );
}

function isLightHex(hex: string): boolean {
  const clean = hex.replace('#', '');
  if (clean.length !== 6) return false;
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.6;
}
