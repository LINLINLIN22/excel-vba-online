import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Plus,
} from 'lucide-react';
import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  colIndexToLetter,
  coordToAddress,
  formatCellDisplayValue,
  isCoordInRange,
  iterateRangeCoords,
  normalizeRange,
  parseCellAddress,
  rangeToAddress,
} from '../spreadsheet/addressing';
import {
  CellCoord,
  CellData,
  CellRange,
  CellStyle,
  HorizontalAlign,
  SheetData,
} from '../types/workbook';

type RibbonTab = 'File' | 'Home' | 'Insert' | 'Page layout' | 'Formulas' | 'Data' | 'View';

interface SpreadsheetPaneProps {
  sheets: SheetData[];
  activeSheetId: string;
  selectedRange: CellRange;
  onSelectSheet: (sheetId: string) => void;
  onAddSheet: () => void;
  onSelectRange: (range: CellRange) => void;
  onUpdateCell: (coord: CellCoord, rawValue: string) => void;
  onApplyStyle: (range: CellRange, stylePatch: Partial<CellStyle>) => void;
  onResetExample: () => void;
  onClearSheet: () => void;
  onExportCsv: () => void;
  onInsertSampleRow: () => void;
}

const FONT_COLORS = [
  '#111C16',
  '#FFFFFF',
  '#217346',
  '#DC2626',
  '#D97706',
  '#2563EB',
];

export function SpreadsheetPane({
  sheets,
  activeSheetId,
  selectedRange,
  onSelectSheet,
  onAddSheet,
  onSelectRange,
  onUpdateCell,
  onApplyStyle,
  onResetExample,
  onClearSheet,
  onExportCsv,
  onInsertSampleRow,
}: SpreadsheetPaneProps) {
  const activeSheet = useMemo(
    () => sheets.find((s) => s.id === activeSheetId) ?? sheets[0],
    [sheets, activeSheetId]
  );

  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('Home');
  const [showGridlines, setShowGridlines] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [zoom, setZoom] = useState(100);
  const [editingCoord, setEditingCoord] = useState<CellCoord | null>(null);
  const [editDraft, setEditDraft] = useState('');
  const [formulaDraft, setFormulaDraft] = useState('');
  const [clipboardCell, setClipboardCell] = useState<CellData | null>(null);
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const dragAnchorRef = useRef<CellCoord | null>(null);
  const tabsScrollRef = useRef<HTMLDivElement | null>(null);

  const activeCoord = selectedRange.start;
  const activeAddr = coordToAddress(activeCoord);
  const activeCell = activeSheet.cells[activeAddr];
  const rangeLabel = rangeToAddress(selectedRange);

  useEffect(() => {
    setFormulaDraft(
      activeCell?.formula ??
        (activeCell?.value !== null && activeCell?.value !== undefined
          ? String(activeCell.value)
          : '')
    );
  }, [activeAddr, activeCell?.formula, activeCell?.value]);

  useEffect(() => {
    const handleMouseUp = () => setIsDragging(false);
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  const displayRows = Math.max(13, activeSheet.rows);
  const displayCols = Math.max(7, activeSheet.cols);

  const activeStyle = activeCell?.style ?? {};

  // Compute live Sum for status bar
  const statusSummary = useMemo(() => {
    const coords = iterateRangeCoords(selectedRange);
    let nums: number[] = [];
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

  const startCellEdit = (coord: CellCoord, initialText?: string) => {
    const addr = coordToAddress(coord);
    const cell = activeSheet.cells[addr];
    setEditingCoord(coord);
    setEditDraft(
      initialText !== undefined
        ? initialText
        : cell?.formula ?? (cell?.value !== null && cell?.value !== undefined ? String(cell.value) : '')
    );
  };

  const commitCellEdit = () => {
    if (!editingCoord) return;
    onUpdateCell(editingCoord, editDraft);
    setEditingCoord(null);
  };

  const handleCellMouseDown = (coord: CellCoord, e: React.MouseEvent) => {
    if (editingCoord && (editingCoord.row !== coord.row || editingCoord.col !== coord.col)) {
      commitCellEdit();
    }
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

  const handleFormulaSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onUpdateCell(activeCoord, formulaDraft);
  };

  const handleCopyActive = () => {
    const cell = activeSheet.cells[activeAddr];
    setClipboardCell(cell ? { ...cell, style: cell.style ? { ...cell.style } : undefined } : { value: '' });
  };

  const handlePaste = () => {
    if (!clipboardCell) {
      handleCopyActive();
      return;
    }
    const coords = iterateRangeCoords(selectedRange);
    for (const c of coords) {
      const raw =
        clipboardCell.formula ??
        (clipboardCell.value !== null && clipboardCell.value !== undefined
          ? String(clipboardCell.value)
          : '');
      onUpdateCell(c, raw);
      if (clipboardCell.style) {
        onApplyStyle({ start: c, end: c }, clipboardCell.style);
      }
    }
  };

  const ribbonTabs: RibbonTab[] = [
    'File',
    'Home',
    'Insert',
    'Page layout',
    'Formulas',
    'Data',
    'View',
  ];

  const colLetters = useMemo(
    () => Array.from({ length: displayCols }, (_, idx) => colIndexToLetter(idx + 1)),
    [displayCols]
  );

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
            Saved for this session
          </span>
        </div>

        <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full bg-[#EDF5F0] dark:bg-[#1B3325] text-[#157347] dark:text-[#65CE91] text-[11px] font-medium whitespace-nowrap shrink-0">
          <span className="w-1.5 h-1.5 rounded-full bg-[#18804C] dark:bg-[#4ADE80]" />
          <span>Live preview</span>
        </div>
      </div>

      {/* Ribbon Menu Tabs */}
      <div className="flex items-center gap-5 px-4 border-b border-[#E7EEE9] dark:border-[#21362A] overflow-x-auto vba-scroll">
        {ribbonTabs.map((tab) => {
          const isActive = ribbonTab === tab;
          return (
            <button
              key={tab}
              type="button"
              onClick={() => setRibbonTab(tab)}
              className={`relative py-1.5 text-xs font-medium whitespace-nowrap shrink-0 transition-colors focus-visible:outline-none ${
                isActive
                  ? 'text-[#157347] dark:text-[#57C587]'
                  : 'text-[#526359] dark:text-[#9CB3A6] hover:text-[#111C16] dark:hover:text-[#E8F1EC]'
              }`}
            >
              {tab}
              {isActive && (
                <span className="absolute bottom-0 left-0 right-0 h-[2px] bg-[#157347] dark:bg-[#39A96B] rounded-t" />
              )}
            </button>
          );
        })}
      </div>

      {/* Ribbon Toolbar */}
      <div className="px-3.5 py-1.5 border-b border-[#E7EEE9] dark:border-[#21362A] bg-[#FAFDFB] dark:bg-[#17291F] min-h-[46px] flex items-center overflow-x-auto vba-scroll">
        {ribbonTab === 'Home' && (
          <div className="flex items-center gap-2.5 text-xs text-[#1D2D24] dark:text-[#D5E3DB]">
            {/* Clipboard */}
            <div className="flex flex-col items-start pr-2.5 border-r border-[#E2EAE5] dark:border-[#243B2E] shrink-0">
              <button
                type="button"
                onClick={handlePaste}
                title={clipboardCell ? 'Paste copied cell' : 'Copy active cell'}
                className="px-2 py-1 rounded hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] font-medium transition-colors whitespace-nowrap"
              >
                {clipboardCell ? 'Paste' : 'Copy'}
              </button>
              <span className="text-[10px] text-[#7A8C81] dark:text-[#6D8577] px-2">
                Clipboard
              </span>
            </div>

            {/* Font */}
            <div className="flex flex-col items-start pr-2.5 border-r border-[#E2EAE5] dark:border-[#243B2E] shrink-0">
              <div className="flex items-center gap-1.5 relative">
                <select
                  aria-label="Font family"
                  value={activeStyle.fontFamily ?? 'Calibri'}
                  onChange={(e) =>
                    onApplyStyle(selectedRange, { fontFamily: e.target.value })
                  }
                  className="h-7 px-2 rounded border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] text-xs focus:outline-none focus:ring-1 focus:ring-[#157347]"
                >
                  <option value="Calibri">Calibri</option>
                  <option value="Plus Jakarta Sans">Sans</option>
                  <option value="JetBrains Mono">Mono</option>
                </select>
                <select
                  aria-label="Font size"
                  value={activeStyle.fontSize ?? 11}
                  onChange={(e) =>
                    onApplyStyle(selectedRange, { fontSize: Number(e.target.value) })
                  }
                  className="h-7 px-1.5 rounded border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] text-xs tabular-nums focus:outline-none focus:ring-1 focus:ring-[#157347]"
                >
                  {[10, 11, 12, 14, 16].map((sz) => (
                    <option key={sz} value={sz}>
                      {sz}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { bold: !activeStyle.bold })
                  }
                  aria-pressed={Boolean(activeStyle.bold)}
                  title="Bold"
                  className={`w-7 h-7 rounded font-bold transition-colors ${
                    activeStyle.bold
                      ? 'bg-[#DFF0E6] dark:bg-[#244331] text-[#157347] dark:text-[#6EE7A0]'
                      : 'hover:bg-[#EAF2ED] dark:hover:bg-[#22392B]'
                  }`}
                >
                  B
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { italic: !activeStyle.italic })
                  }
                  aria-pressed={Boolean(activeStyle.italic)}
                  title="Italic"
                  className={`w-7 h-7 rounded italic font-serif transition-colors ${
                    activeStyle.italic
                      ? 'bg-[#DFF0E6] dark:bg-[#244331] text-[#157347] dark:text-[#6EE7A0]'
                      : 'hover:bg-[#EAF2ED] dark:hover:bg-[#22392B]'
                  }`}
                >
                  I
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { underline: !activeStyle.underline })
                  }
                  aria-pressed={Boolean(activeStyle.underline)}
                  title="Underline"
                  className={`w-7 h-7 rounded underline transition-colors ${
                    activeStyle.underline
                      ? 'bg-[#DFF0E6] dark:bg-[#244331] text-[#157347] dark:text-[#6EE7A0]'
                      : 'hover:bg-[#EAF2ED] dark:hover:bg-[#22392B]'
                  }`}
                >
                  U
                </button>
                <div className="relative">
                  <button
                    type="button"
                    onClick={() => setShowColorPicker((v) => !v)}
                    title="Font color"
                    className="w-7 h-7 rounded flex flex-col items-center justify-center hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] transition-colors"
                  >
                    <span className="leading-none font-semibold">A</span>
                    <span
                      className="w-3.5 h-0.5 mt-0.5 rounded-full"
                      style={{
                        backgroundColor: activeStyle.color ?? '#18804C',
                      }}
                    />
                  </button>
                  {showColorPicker && (
                    <div className="absolute top-8 left-0 z-20 p-1.5 rounded-lg border border-[#DCE5DF] dark:border-[#243B2E] bg-white dark:bg-[#14241B] shadow-lg flex items-center gap-1">
                      {FONT_COLORS.map((hex) => (
                        <button
                          key={hex}
                          type="button"
                          onClick={() => {
                            onApplyStyle(selectedRange, { color: hex });
                            setShowColorPicker(false);
                          }}
                          aria-label={`Set font color ${hex}`}
                          className="w-5 h-5 rounded border border-black/15"
                          style={{ backgroundColor: hex }}
                        />
                      ))}
                    </div>
                  )}
                </div>
              </div>
              <span className="text-[10px] text-[#7A8C81] dark:text-[#6D8577] mt-0.5">
                Font
              </span>
            </div>

            {/* Alignment */}
            <div className="flex flex-col items-start pr-2.5 border-r border-[#E2EAE5] dark:border-[#243B2E] shrink-0">
              <div className="flex items-center gap-1">
                {(
                  [
                    { label: 'Align left', align: 'left' },
                    { label: 'Center', align: 'center' },
                  ] as Array<{ label: string; align: HorizontalAlign }>
                ).map((item) => (
                  <button
                    key={item.align}
                    type="button"
                    onClick={() =>
                      onApplyStyle(selectedRange, { horizontalAlign: item.align })
                    }
                    className={`px-2 h-7 rounded whitespace-nowrap transition-colors ${
                      activeStyle.horizontalAlign === item.align
                        ? 'bg-[#DFF0E6] dark:bg-[#244331] text-[#157347] dark:text-[#6EE7A0]'
                        : 'hover:bg-[#EAF2ED] dark:hover:bg-[#22392B]'
                    }`}
                  >
                    {item.label}
                  </button>
                ))}
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { wrapText: !activeStyle.wrapText })
                  }
                  className={`px-2 h-7 rounded whitespace-nowrap transition-colors ${
                    activeStyle.wrapText
                      ? 'bg-[#DFF0E6] dark:bg-[#244331] text-[#157347] dark:text-[#6EE7A0]'
                      : 'hover:bg-[#EAF2ED] dark:hover:bg-[#22392B]'
                  }`}
                >
                  Wrap text
                </button>
              </div>
              <span className="text-[10px] text-[#7A8C81] dark:text-[#6D8577] mt-0.5 px-2">
                Alignment
              </span>
            </div>

            {/* Number */}
            <div className="flex flex-col items-start shrink-0">
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { numberFormat: '$#,##0' })
                  }
                  title="Currency format"
                  className="w-7 h-7 rounded hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] font-medium transition-colors"
                >
                  $
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { numberFormat: '0.0%' })
                  }
                  title="Percent format"
                  className="w-7 h-7 rounded hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] font-medium transition-colors"
                >
                  %
                </button>
                <button
                  type="button"
                  onClick={() =>
                    onApplyStyle(selectedRange, { numberFormat: '#,##0.00' })
                  }
                  title="Two decimals"
                  className="px-2 h-7 rounded hover:bg-[#EAF2ED] dark:hover:bg-[#22392B] font-mono text-[11px] transition-colors"
                >
                  .00
                </button>
              </div>
              <span className="text-[10px] text-[#7A8C81] dark:text-[#6D8577] mt-0.5 px-1.5">
                Number
              </span>
            </div>
          </div>
        )}

        {ribbonTab === 'File' && (
          <div className="flex items-center gap-2 text-xs">
            <button
              type="button"
              onClick={onResetExample}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-medium transition-colors whitespace-nowrap"
            >
              Reset sheet to example
            </button>
            <button
              type="button"
              onClick={onClearSheet}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-medium transition-colors whitespace-nowrap"
            >
              Clear sheet
            </button>
            <button
              type="button"
              onClick={onExportCsv}
              className="px-3 py-1.5 rounded-md bg-[#157347] hover:bg-[#12633B] text-white font-medium transition-colors whitespace-nowrap"
            >
              Export CSV
            </button>
          </div>
        )}

        {ribbonTab === 'Insert' && (
          <div className="flex items-center gap-2 text-xs">
            <button
              type="button"
              onClick={onInsertSampleRow}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-medium transition-colors whitespace-nowrap"
            >
              Insert sample row
            </button>
            <button
              type="button"
              onClick={onAddSheet}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-medium transition-colors whitespace-nowrap"
            >
              New worksheet
            </button>
          </div>
        )}

        {ribbonTab === 'Page layout' && (
          <div className="flex items-center gap-2 text-xs">
            <button
              type="button"
              onClick={() => setShowGridlines((v) => !v)}
              className={`px-3 py-1.5 rounded-md border font-medium transition-colors whitespace-nowrap ${
                showGridlines
                  ? 'border-[#157347] bg-[#EAF4EE] dark:bg-[#1F3A2B] text-[#157347] dark:text-[#6EE7A0]'
                  : 'border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018]'
              }`}
            >
              Gridlines: {showGridlines ? 'On' : 'Off'}
            </button>
          </div>
        )}

        {ribbonTab === 'Formulas' && (
          <div className="flex items-center gap-2 text-xs">
            <button
              type="button"
              onClick={() => onUpdateCell(activeCoord, '=SUM(C2:C4)')}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-mono transition-colors whitespace-nowrap"
            >
              =SUM(C2:C4)
            </button>
            <button
              type="button"
              onClick={() => onUpdateCell(activeCoord, '=AVERAGE(C2:C4)')}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-mono transition-colors whitespace-nowrap"
            >
              =AVERAGE(C2:C4)
            </button>
            <button
              type="button"
              onClick={() => setShowFormulas((v) => !v)}
              className={`px-3 py-1.5 rounded-md border font-medium transition-colors whitespace-nowrap ${
                showFormulas
                  ? 'border-[#157347] bg-[#EAF4EE] dark:bg-[#1F3A2B] text-[#157347] dark:text-[#6EE7A0]'
                  : 'border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018]'
              }`}
            >
              Show formulas: {showFormulas ? 'On' : 'Off'}
            </button>
          </div>
        )}

        {ribbonTab === 'Data' && (
          <div className="flex items-center gap-2 text-xs">
            <button
              type="button"
              onClick={onExportCsv}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-medium transition-colors whitespace-nowrap"
            >
              Download as CSV
            </button>
            <button
              type="button"
              onClick={onResetExample}
              className="px-3 py-1.5 rounded-md border border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018] hover:bg-[#EDF5F0] dark:hover:bg-[#1D3225] font-medium transition-colors whitespace-nowrap"
            >
              Restore initial dataset
            </button>
          </div>
        )}

        {ribbonTab === 'View' && (
          <div className="flex items-center gap-2 text-xs">
            {[85, 100, 115].map((pct) => (
              <button
                key={pct}
                type="button"
                onClick={() => setZoom(pct)}
                className={`px-2.5 py-1.5 rounded-md border tabular-nums font-medium transition-colors ${
                  zoom === pct
                    ? 'border-[#157347] bg-[#EAF4EE] dark:bg-[#1F3A2B] text-[#157347] dark:text-[#6EE7A0]'
                    : 'border-[#D7E2DC] dark:border-[#294234] bg-white dark:bg-[#122018]'
                }`}
              >
                {pct}%
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Formula Bar */}
      <form
        onSubmit={handleFormulaSubmit}
        className="flex items-center gap-2 px-3 py-1.5 border-b border-[#E7EEE9] dark:border-[#21362A] bg-white dark:bg-[#132219]"
      >
        <div className="flex items-center justify-between w-20 px-2 py-1 rounded border border-[#E2EAE5] dark:border-[#243B2E] text-xs font-medium text-[#1D2D24] dark:text-[#D5E3DB] tabular-nums shrink-0">
          <input
            aria-label="Active cell address"
            value={rangeLabel}
            onChange={(e) => {
              const coord = parseCellAddress(e.target.value);
              if (coord) onSelectRange({ start: coord, end: coord });
            }}
            className="w-12 bg-transparent focus:outline-none uppercase"
          />
          <ChevronDown className="w-3 h-3 opacity-60 shrink-0" />
        </div>

        <span
          aria-hidden="true"
          className="px-2 text-sm italic font-serif text-[#526359] dark:text-[#8FA699] select-none border-r border-[#E2EAE5] dark:border-[#243B2E]"
        >
          fx
        </span>

        <input
          aria-label="Formula bar"
          type="text"
          value={formulaDraft}
          onChange={(e) => setFormulaDraft(e.target.value)}
          onBlur={() => {
            const currentRaw =
              activeCell?.formula ??
              (activeCell?.value !== null && activeCell?.value !== undefined
                ? String(activeCell.value)
                : '');
            if (formulaDraft !== currentRaw) {
              onUpdateCell(activeCoord, formulaDraft);
            }
          }}
          placeholder="Enter value or formula (e.g. =SUM(C2:C4))"
          className="flex-1 bg-transparent text-xs font-mono text-[#111C16] dark:text-[#E8F1EC] placeholder:text-[#8DA094] focus:outline-none py-1"
        />
      </form>

      {/* Spreadsheet Grid */}
      <div className="flex-1 min-h-0 overflow-auto vba-scroll bg-white dark:bg-[#122018] select-none">
        <table
          className="border-collapse table-fixed w-full"
          style={{
            fontSize: `${Math.round((12 * zoom) / 100)}px`,
          }}
        >
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
                    style={{ width: `${Math.round((widthPx * zoom) / 100)}px` }}
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
                    const isEditing =
                      editingCoord?.row === rowNum && editingCoord?.col === colNum;

                    const rawDisplay = showFormulas && cell?.formula
                      ? cell.formula
                      : formatCellDisplayValue(cell?.value, style.numberFormat);

                    const isNumeric =
                      typeof cell?.value === 'number' ||
                      (typeof cell?.value === 'string' &&
                        cell.value.trim() !== '' &&
                        Number.isFinite(Number(cell.value)));

                    const align: HorizontalAlign =
                      style.horizontalAlign ?? (isNumeric ? 'right' : 'left');

                    // Ensure explicit VBA light background stays legible in dark mode
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
                        onDoubleClick={() => startCellEdit(coord)}
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
                        className={`relative px-2 py-1 tabular-nums overflow-hidden text-ellipsis cursor-cell transition-colors ${
                          showGridlines
                            ? 'border-b border-r border-[#E6EDE9] dark:border-[#1E3227]'
                            : 'border-b border-r border-transparent'
                        } ${
                          !customTextColor
                            ? 'text-[#18261E] dark:text-[#D5E3DB]'
                            : ''
                        }`}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            value={editDraft}
                            onChange={(e) => setEditDraft(e.target.value)}
                            onBlur={commitCellEdit}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                commitCellEdit();
                                onSelectRange({
                                  start: { row: Math.min(displayRows, rowNum + 1), col: colNum },
                                  end: { row: Math.min(displayRows, rowNum + 1), col: colNum },
                                });
                              } else if (e.key === 'Escape') {
                                setEditingCoord(null);
                              }
                            }}
                            className="absolute inset-0 w-full h-full px-2 bg-white dark:bg-[#0D1A12] text-[#111C16] dark:text-[#E8F1EC] text-xs font-mono focus:outline-none z-20"
                          />
                        ) : (
                          <span>{rawDisplay}</span>
                        )}

                        {/* Excel active selection box + fill handle */}
                        {isActive && !isEditing && (
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

      {/* Sheet Tabs Bar */}
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

          <button
            type="button"
            onClick={onAddSheet}
            title="Add worksheet"
            aria-label="Add worksheet"
            className="ml-1 p-1 rounded text-[#526359] dark:text-[#9CB3A6] hover:bg-[#EAF2ED] dark:hover:bg-[#1E3327] transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
          </button>
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

        <div className="flex items-center gap-2 tabular-nums shrink-0">
          <button
            type="button"
            onClick={() => setZoom((z) => Math.max(75, z - 10))}
            aria-label="Zoom out"
            className="px-1 hover:text-[#111C16] dark:hover:text-[#E8F1EC]"
          >
            -
          </button>
          <input
            aria-label="Zoom slider"
            type="range"
            min={75}
            max={130}
            step={5}
            value={zoom}
            onChange={(e) => setZoom(Number(e.target.value))}
            className="w-16 accent-[#157347] dark:accent-[#39A96B] cursor-pointer"
          />
          <button
            type="button"
            onClick={() => setZoom((z) => Math.min(130, z + 10))}
            aria-label="Zoom in"
            className="px-1 hover:text-[#111C16] dark:hover:text-[#E8F1EC]"
          >
            +
          </button>
          <span className="w-9 text-right">{zoom}%</span>
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
