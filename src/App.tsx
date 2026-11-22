import { ArrowUpRight, Info, Play } from 'lucide-react';
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ExamplesView } from './components/ExamplesView';
import { LearnView } from './components/LearnView';
import { QuickGuideModal } from './components/QuickGuideModal';
import { SpreadsheetPane } from './components/SpreadsheetPane';
import { ActiveTab, TopBar } from './components/TopBar';
import { VbaEditorPane } from './components/VbaEditorPane';
import { EXAMPLE_PRESETS } from './data/examples';
import { coordToAddress, iterateRangeCoords } from './spreadsheet/addressing';
import { recalculateSheets } from './spreadsheet/formulaEngine';
import {
  CellCoord,
  CellRange,
  CellStyle,
  ExamplePreset,
  ExecutionResult,
  SheetData,
  VbaModule,
} from './types/workbook';
import { executeVbaMacro } from './vba/interpreter';
import { lintVbaCode } from './vba/linter';

const THEME_STORAGE_KEY = 'vba-playground-theme';
const SESSION_STORAGE_KEY = 'vba-playground-session-v1';
const SPLIT_STORAGE_KEY = 'vba-playground-split-percent';
const DEFAULT_SPLIT_PERCENT = 58;

interface SavedSession {
  activePresetId: string;
  sheets: SheetData[];
  activeSheetId: string;
  selectedRange: CellRange;
  modules: VbaModule[];
  activeModuleId: string;
  runCount: number;
}

function cloneSheets(sheets: SheetData[]): SheetData[] {
  return sheets.map((s) => ({
    ...s,
    colWidths: { ...s.colWidths },
    cells: Object.fromEntries(
      Object.entries(s.cells).map(([k, v]) => [
        k,
        { ...v, style: v.style ? { ...v.style } : undefined },
      ])
    ),
  }));
}

function buildInitialState() {
  const defaultPreset = EXAMPLE_PRESETS[0];

  try {
    const raw = localStorage.getItem(SESSION_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as SavedSession;
      if (parsed.sheets?.length > 0 && parsed.modules?.length > 0) {
        return {
          ...parsed,
          executionResult: {
            status: 'idle' as const,
            headline: 'Session restored',
            detail: 'Click Run macro to execute your saved module.',
            durationMs: 0,
            logs: [],
            changedCellsCount: 0,
          },
        };
      }
    }
  } catch {
    // Ignore invalid storage
  }

  const initialSheets = cloneSheets(defaultPreset.sheets);
  const initialSelected = defaultPreset.selectedRange ?? {
    start: { row: 6, col: 3 },
    end: { row: 6, col: 3 },
  };

  // Run the default example once on fresh visit so the live spreadsheet & execution state match
  const initialRun = executeVbaMacro({
    code: defaultPreset.code,
    sheets: initialSheets,
    activeSheetId: initialSheets[0].id,
    selectedRange: initialSelected,
  });

  return {
    activePresetId: defaultPreset.id,
    sheets: initialRun.sheets,
    activeSheetId: initialRun.activeSheetId,
    selectedRange: initialSelected,
    modules: [
      {
        id: 'mod-1',
        name: 'Module1.bas',
        code: defaultPreset.code,
      },
    ],
    activeModuleId: 'mod-1',
    runCount: 1,
    executionResult: initialRun.result,
  };
}

export default function App() {
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    try {
      const stored = localStorage.getItem(THEME_STORAGE_KEY);
      if (stored === 'dark' || stored === 'light') return stored;
    } catch {
      // Ignore storage access issues
    }
    return 'light';
  });

  useEffect(() => {
    document.documentElement.classList.toggle('dark', theme === 'dark');
    try {
      localStorage.setItem(THEME_STORAGE_KEY, theme);
    } catch {
      // Ignore
    }
  }, [theme]);

  const [initial] = useState(() => buildInitialState());

  const [activeTab, setActiveTab] = useState<ActiveTab>('playground');
  const [quickGuideOpen, setQuickGuideOpen] = useState(false);

  const [splitPercent, setSplitPercent] = useState<number>(() => {
    try {
      const saved = Number(localStorage.getItem(SPLIT_STORAGE_KEY));
      if (Number.isFinite(saved) && saved >= 32 && saved <= 72) return saved;
    } catch {
      // Ignore
    }
    return DEFAULT_SPLIT_PERCENT;
  });
  const [isResizing, setIsResizing] = useState(false);
  const splitContainerRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem(SPLIT_STORAGE_KEY, String(splitPercent));
    } catch {
      // Ignore
    }
  }, [splitPercent]);

  useEffect(() => {
    if (!isResizing) return;

    const handlePointerMove = (e: PointerEvent) => {
      const container = splitContainerRef.current;
      if (!container) return;
      const rect = container.getBoundingClientRect();
      if (rect.width <= 0) return;
      const rawPct = ((e.clientX - rect.left) / rect.width) * 100;
      const clamped = Math.max(32, Math.min(72, Math.round(rawPct * 10) / 10));
      setSplitPercent(clamped);
    };

    const handlePointerUp = () => {
      setIsResizing(false);
    };

    document.body.style.cursor = 'col-resize';
    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);

    return () => {
      document.body.style.cursor = '';
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [isResizing]);

  const [activePresetId, setActivePresetId] = useState<string>(initial.activePresetId);
  const [sheets, setSheets] = useState<SheetData[]>(initial.sheets);
  const [activeSheetId, setActiveSheetId] = useState<string>(initial.activeSheetId);
  const [selectedRange, setSelectedRange] = useState<CellRange>(initial.selectedRange);
  const [modules, setModules] = useState<VbaModule[]>(initial.modules);
  const [activeModuleId, setActiveModuleId] = useState<string>(initial.activeModuleId);
  const [runCount, setRunCount] = useState<number>(initial.runCount);
  const [executionResult, setExecutionResult] = useState<ExecutionResult>(
    initial.executionResult
  );

  // Persist session locally in browser
  useEffect(() => {
    try {
      const payload: SavedSession = {
        activePresetId,
        sheets,
        activeSheetId,
        selectedRange,
        modules,
        activeModuleId,
        runCount,
      };
      localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // Ignore quota errors
    }
  }, [activePresetId, sheets, activeSheetId, selectedRange, modules, activeModuleId, runCount]);

  const activePreset = useMemo(
    () => EXAMPLE_PRESETS.find((p) => p.id === activePresetId) ?? EXAMPLE_PRESETS[0],
    [activePresetId]
  );

  const activeModule = useMemo(
    () => modules.find((m) => m.id === activeModuleId) ?? modules[0],
    [modules, activeModuleId]
  );

  // Real-time syntax check (debounced slightly or memoized directly since lexer/linter is <1ms)
  const diagnostics = useMemo(() => lintVbaCode(activeModule.code), [activeModule.code]);

  const handleRunMacro = useCallback(
    (overrideCode?: string, overrideSheets?: SheetData[], overrideSheetId?: string, overrideRange?: CellRange) => {
      const targetCode = overrideCode ?? activeModule.code;
      const targetSheets = overrideSheets ?? sheets;
      const targetSheetId = overrideSheetId ?? activeSheetId;
      const targetRange = overrideRange ?? selectedRange;

      const output = executeVbaMacro({
        code: targetCode,
        sheets: targetSheets,
        activeSheetId: targetSheetId,
        selectedRange: targetRange,
      });

      setSheets(output.sheets);
      setActiveSheetId(output.activeSheetId);
      setSelectedRange(output.selectedRange);
      setExecutionResult(output.result);
      if (output.result.status === 'success') {
        setRunCount((c) => c + 1);
      }
    },
    [activeModule.code, sheets, activeSheetId, selectedRange]
  );

  // Global shortcut: Ctrl/Cmd + Enter when in Playground
  useEffect(() => {
    const onGlobalKey = (e: KeyboardEvent) => {
      if (activeTab === 'playground' && (e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        const tag = (e.target as HTMLElement)?.tagName;
        if (tag !== 'TEXTAREA') {
          e.preventDefault();
          handleRunMacro();
        }
      }
    };
    window.addEventListener('keydown', onGlobalKey);
    return () => window.removeEventListener('keydown', onGlobalKey);
  }, [activeTab, handleRunMacro]);

  const handleLoadPreset = (
    preset: ExamplePreset,
    autoRun = true,
    customCode?: string
  ) => {
    const freshSheets = cloneSheets(preset.sheets);
    const freshRange = preset.selectedRange ?? {
      start: { row: 1, col: 1 },
      end: { row: 1, col: 1 },
    };
    const nextCode = customCode ?? preset.code;

    setActivePresetId(preset.id);
    setModules((prev) => {
      const updated = [...prev];
      const idx = updated.findIndex((m) => m.id === activeModuleId);
      if (idx !== -1) {
        updated[idx] = { ...updated[idx], code: nextCode };
      } else {
        updated[0] = { ...updated[0], code: nextCode };
      }
      return updated;
    });

    if (autoRun) {
      const output = executeVbaMacro({
        code: nextCode,
        sheets: freshSheets,
        activeSheetId: freshSheets[0].id,
        selectedRange: freshRange,
      });
      setSheets(output.sheets);
      setActiveSheetId(output.activeSheetId);
      setSelectedRange(output.selectedRange);
      setExecutionResult(output.result);
      if (output.result.status === 'success') {
        setRunCount((c) => c + 1);
      }
    } else {
      setSheets(freshSheets);
      setActiveSheetId(freshSheets[0].id);
      setSelectedRange(freshRange);
      setExecutionResult({
        status: 'idle',
        headline: 'Ready',
        detail: 'Click Run macro to execute.',
        durationMs: 0,
        logs: [],
        changedCellsCount: 0,
      });
    }

    setActiveTab('playground');
  };

  const handleUpdateCell = (coord: CellCoord, rawValue: string) => {
    setSheets((prev) => {
      const updated = prev.map((sh) => {
        if (sh.id !== activeSheetId) return sh;
        const addr = coordToAddress(coord);
        const existing = sh.cells[addr] ?? { value: null };
        let nextCell;
        if (rawValue.trim().startsWith('=')) {
          nextCell = { ...existing, formula: rawValue.trim(), value: rawValue.trim() };
        } else if (rawValue.trim() === '') {
          nextCell = { ...existing, formula: undefined, value: null };
        } else {
          const num = Number(rawValue);
          const parsedVal =
            rawValue.trim() !== '' && Number.isFinite(num) ? num : rawValue;
          nextCell = { ...existing, formula: undefined, value: parsedVal };
        }
        return {
          ...sh,
          cells: {
            ...sh.cells,
            [addr]: nextCell,
          },
        };
      });
      return recalculateSheets(updated);
    });
  };

  const handleApplyStyle = (range: CellRange, stylePatch: Partial<CellStyle>) => {
    const coords = iterateRangeCoords(range);
    setSheets((prev) =>
      prev.map((sh) => {
        if (sh.id !== activeSheetId) return sh;
        const nextCells = { ...sh.cells };
        for (const c of coords) {
          const addr = coordToAddress(c);
          const existing = nextCells[addr] ?? { value: null };
          nextCells[addr] = {
            ...existing,
            style: { ...(existing.style ?? {}), ...stylePatch },
          };
        }
        return { ...sh, cells: nextCells };
      })
    );
  };

  const handleAddSheet = () => {
    const nextNum = sheets.length + 1;
    const newId = `sheet-${Date.now()}`;
    const newSheet: SheetData = {
      id: newId,
      name: `Sheet${nextNum}`,
      rows: 15,
      cols: 7,
      cells: {},
      colWidths: { 1: 102, 2: 76, 3: 84, 4: 74, 5: 74, 6: 74, 7: 74 },
    };
    setSheets((prev) => [...prev, newSheet]);
    setActiveSheetId(newId);
    setSelectedRange({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
  };

  const handleResetExample = () => {
    const freshSheets = cloneSheets(activePreset.sheets);
    setSheets(freshSheets);
    setActiveSheetId(freshSheets[0].id);
    setSelectedRange({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
    setExecutionResult({
      status: 'idle',
      headline: 'Workbook reset',
      detail: 'Restored initial data. Click Run macro to test your code.',
      durationMs: 0,
      logs: [],
      changedCellsCount: 0,
    });
  };

  const handleClearSheet = () => {
    setSheets((prev) =>
      prev.map((sh) => (sh.id === activeSheetId ? { ...sh, cells: {} } : sh))
    );
  };

  const handleInsertSampleRow = () => {
    const sh = sheets.find((s) => s.id === activeSheetId) ?? sheets[0];
    let lastUsedRow = 1;
    for (const addr of Object.keys(sh.cells)) {
      const m = /(\d+)$/.exec(addr);
      if (m) lastUsedRow = Math.max(lastUsedRow, parseInt(m[1], 10));
    }
    const targetRow = Math.min(sh.rows, lastUsedRow + 1);
    handleUpdateCell({ row: targetRow, col: 1 }, 'Stylus Pen');
    handleUpdateCell({ row: targetRow, col: 2 }, '12');
    handleUpdateCell({ row: targetRow, col: 3 }, '360');
  };

  const handleExportCsv = () => {
    const sh = sheets.find((s) => s.id === activeSheetId) ?? sheets[0];
    const lines: string[] = [];
    for (let r = 1; r <= sh.rows; r++) {
      const rowVals: string[] = [];
      let hasData = false;
      for (let c = 1; c <= sh.cols; c++) {
        const cell = sh.cells[coordToAddress({ row: r, col: c })];
        const val = cell?.value ?? '';
        if (val !== '' && val !== null) hasData = true;
        const str = String(val).replace(/"/g, '""');
        rowVals.push(str.includes(',') ? `"${str}"` : str);
      }
      if (hasData) lines.push(rowVals.join(','));
    }
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sh.name}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleAddModule = () => {
    const nextIdx = modules.length + 1;
    const newMod: VbaModule = {
      id: `mod-${Date.now()}`,
      name: `Module${nextIdx}.bas`,
      code: `Sub Macro${nextIdx}()\n    ' Write your VBA macro here\n    Range("A1").Value = "Hello VBA"\nEnd Sub`,
    };
    setModules((prev) => [...prev, newMod]);
    setActiveModuleId(newMod.id);
  };

  const handleCloseModule = (id: string) => {
    if (modules.length <= 1) return;
    const remaining = modules.filter((m) => m.id !== id);
    setModules(remaining);
    if (activeModuleId === id) {
      setActiveModuleId(remaining[0].id);
    }
  };

  const handleChangeModuleCode = (nextCode: string) => {
    setModules((prev) =>
      prev.map((m) => (m.id === activeModuleId ? { ...m, code: nextCode } : m))
    );
  };

  return (
    <div className="min-h-screen flex flex-col bg-[#F6F8F7] dark:bg-[#0D1912] text-[#111C16] dark:text-[#E8F1EC] transition-colors">
      <TopBar
        activeTab={activeTab}
        onChangeTab={setActiveTab}
        theme={theme}
        onToggleTheme={setTheme}
        onOpenQuickGuide={() => setQuickGuideOpen(true)}
      />

      <main className="flex-1 w-full max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8 space-y-6">
        {activeTab === 'playground' && (
          <>
            {/* Hero & Primary Action Row */}
            <section className="flex flex-col md:flex-row md:items-end justify-between gap-4">
              <div className="space-y-1.5 max-w-2xl">
                <h1 className="text-2xl sm:text-[30px] font-bold tracking-tight text-[#111C16] dark:text-[#E8F1EC] leading-tight text-balance">
                  A little VBA. A lot of possibility.
                </h1>
                <p className="text-sm text-[#526359] dark:text-[#9CB3A6]">
                  Write a macro, run it, and watch your spreadsheet come to life. No Excel required.
                </p>
              </div>

              {/* Example Selector & Run Macro CTA */}
              <div className="flex flex-wrap sm:flex-nowrap items-start gap-3 shrink-0">
                <div className="flex flex-col gap-1">
                  <label
                    htmlFor="example-select"
                    className="text-[10px] font-semibold uppercase tracking-wider text-[#6B7E73] dark:text-[#849B8E]"
                  >
                    Try an example
                  </label>
                  <select
                    id="example-select"
                    value={activePresetId}
                    onChange={(e) => {
                      const chosen = EXAMPLE_PRESETS.find((p) => p.id === e.target.value);
                      if (chosen) handleLoadPreset(chosen, true);
                    }}
                    className="h-10 min-w-[200px] px-3.5 rounded-lg border border-[#D4DFD8] dark:border-[#243B2E] bg-white dark:bg-[#14241B] text-xs font-medium text-[#111C16] dark:text-[#E8F1EC] focus:outline-none focus:ring-2 focus:ring-[#157347] cursor-pointer"
                  >
                    {EXAMPLE_PRESETS.map((preset) => (
                      <option key={preset.id} value={preset.id}>
                        {preset.title}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="flex flex-col items-center gap-1 pt-[18px]">
                  <button
                    type="button"
                    onClick={() => handleRunMacro()}
                    className="h-10 px-5 rounded-lg bg-[#187E4B] hover:bg-[#146B40] active:scale-[0.99] text-white text-xs font-semibold inline-flex items-center gap-2 shadow-xs transition-all whitespace-nowrap cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-2 focus-visible:ring-[#187E4B]"
                  >
                    <Play className="w-3.5 h-3.5 fill-current" />
                    <span>Run macro</span>
                  </button>
                  <span className="text-[10px] text-[#73847A] dark:text-[#789082] whitespace-nowrap">
                    ⌘ Enter to run
                  </span>
                </div>
              </div>
            </section>

            {/* Dual-Pane Resizable Workspace: Left Spreadsheet, Right VBA Editor on laptop+ screens */}
            <section
              ref={splitContainerRef}
              style={{ '--split-left': `${splitPercent}%` } as React.CSSProperties}
              className={`flex flex-col min-[900px]:flex-row items-stretch gap-5 min-[900px]:gap-0 ${
                isResizing ? 'select-none' : ''
              }`}
            >
              <div className="w-full min-[900px]:w-[calc(var(--split-left)-10px)] min-[900px]:shrink-0 min-w-0">
                <SpreadsheetPane
                  sheets={sheets}
                  activeSheetId={activeSheetId}
                  selectedRange={selectedRange}
                  onSelectSheet={setActiveSheetId}
                  onAddSheet={handleAddSheet}
                  onSelectRange={setSelectedRange}
                  onUpdateCell={handleUpdateCell}
                  onApplyStyle={handleApplyStyle}
                  onResetExample={handleResetExample}
                  onClearSheet={handleClearSheet}
                  onExportCsv={handleExportCsv}
                  onInsertSampleRow={handleInsertSampleRow}
                />
              </div>

              {/* Interactive Resizer Handle (visible on side-by-side laptop/desktop viewports) */}
              <div
                role="separator"
                tabIndex={0}
                aria-orientation="vertical"
                aria-label="Resize spreadsheet and VBA editor panels"
                aria-valuemin={32}
                aria-valuemax={72}
                aria-valuenow={Math.round(splitPercent)}
                title="Drag to resize · Double-click to reset"
                onPointerDown={(e) => {
                  e.preventDefault();
                  setIsResizing(true);
                }}
                onDoubleClick={() => setSplitPercent(DEFAULT_SPLIT_PERCENT)}
                onKeyDown={(e) => {
                  if (e.key === 'ArrowLeft') {
                    e.preventDefault();
                    setSplitPercent((p) => Math.max(32, p - 2));
                  } else if (e.key === 'ArrowRight') {
                    e.preventDefault();
                    setSplitPercent((p) => Math.min(72, p + 2));
                  } else if (e.key === 'Home') {
                    e.preventDefault();
                    setSplitPercent(DEFAULT_SPLIT_PERCENT);
                  }
                }}
                className="hidden min-[900px]:flex w-5 shrink-0 items-center justify-center cursor-col-resize group focus-visible:outline-none"
              >
                <div
                  className={`w-1 h-12 rounded-full transition-colors ${
                    isResizing
                      ? 'bg-[#157347] dark:bg-[#39A96B]'
                      : 'bg-[#CFDCD4] dark:bg-[#274031] group-hover:bg-[#157347] dark:group-hover:bg-[#39A96B] group-focus-visible:bg-[#157347]'
                  }`}
                />
              </div>

              <div className="w-full min-[900px]:flex-1 min-w-0">
                <VbaEditorPane
                  modules={modules}
                  activeModuleId={activeModuleId}
                  onSelectModule={setActiveModuleId}
                  onAddModule={handleAddModule}
                  onCloseModule={handleCloseModule}
                  onChangeCode={handleChangeModuleCode}
                  onRunMacro={() => handleRunMacro()}
                  diagnostics={diagnostics}
                  executionResult={executionResult}
                  runCount={runCount}
                  onClearExecution={() =>
                    setExecutionResult({
                      status: 'idle',
                      headline: 'Cleared',
                      detail: 'Ready to run.',
                      durationMs: 0,
                      logs: [],
                      changedCellsCount: 0,
                    })
                  }
                />
              </div>
            </section>

            {/* Contextual Callout Strip ("See the connection") */}
            <section className="rounded-xl border border-[#D8E7DD] dark:border-[#21382B] bg-[#EDF5F0] dark:bg-[#14251C] px-4 sm:px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-start sm:items-center gap-3">
                <div className="w-7 h-7 rounded-md bg-[#DCECE2] dark:bg-[#1E3829] text-[#157347] dark:text-[#57C587] flex items-center justify-center shrink-0">
                  <Info className="w-4 h-4" />
                </div>
                <div className="text-xs leading-relaxed">
                  <p className="font-semibold text-[#111C16] dark:text-[#E8F1EC]">
                    {activePreset.tipTitle}
                  </p>
                  <p className="text-[#4B6155] dark:text-[#9CB3A6]">
                    {activePreset.tipBody}
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setActiveTab('learn')}
                className="inline-flex items-center gap-1 text-xs font-semibold text-[#157347] dark:text-[#57C587] hover:underline whitespace-nowrap shrink-0 self-end sm:self-center"
              >
                <span>Explore the basics</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            </section>
          </>
        )}

        {activeTab === 'examples' && (
          <ExamplesView
            activePresetId={activePresetId}
            onLoadPreset={(preset, autoRun) => handleLoadPreset(preset, autoRun)}
          />
        )}

        {activeTab === 'learn' && (
          <LearnView
            onTryInPlayground={(preset, customCode) =>
              handleLoadPreset(preset, true, customCode)
            }
          />
        )}
      </main>

      {/* Footer */}
      <footer className="w-full border-t border-[#E2EAE5] dark:border-[#1E3327] mt-auto">
        <div className="max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-[#6B7E73] dark:text-[#7E9689]">
          <span>A safe place to experiment. Changes stay in this session.</span>
          <span>VBA simulator · Common Excel commands supported</span>
        </div>
      </footer>

      <QuickGuideModal
        open={quickGuideOpen}
        onClose={() => setQuickGuideOpen(false)}
        onJumpToLearn={() => setActiveTab('learn')}
      />
    </div>
  );
}
