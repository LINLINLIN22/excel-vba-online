import { ArrowUpRight, Info, Play } from 'lucide-react';
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ExamplesView } from './components/ExamplesView';
import { LearnView } from './components/LearnView';
import { QuickGuideModal } from './components/QuickGuideModal';
import { SpreadsheetPane } from './components/SpreadsheetPane';
import { ActiveTab, TopBar } from './components/TopBar';
import { VbaEditorPane } from './components/VbaEditorPane';
import { createBlankSheet, EXAMPLE_PRESETS } from './data/examples';
import {
  CellRange,
  ExamplePreset,
  ExecutionResult,
  SheetData,
  VbaModule,
} from './types/workbook';
import { executeVbaMacro } from './vba/interpreter';
import { lintVbaCode } from './vba/linter';

const THEME_STORAGE_KEY = 'vba-playground-theme';
const SESSION_STORAGE_KEY = 'vba-playground-session-v3';
const SPLIT_STORAGE_KEY = 'vba-playground-split-percent';
const DEFAULT_SPLIT_PERCENT = 58;

interface SavedSession {
  activePresetId: string;
  modules: VbaModule[];
  activeModuleId: string;
}

interface WorkbookSnapshot {
  sheets: SheetData[];
  activeSheetId: string;
  selectedRange: CellRange;
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

function cloneSnapshot(snap: WorkbookSnapshot): WorkbookSnapshot {
  return {
    sheets: cloneSheets(snap.sheets),
    activeSheetId: snap.activeSheetId,
    selectedRange: {
      start: { ...snap.selectedRange.start },
      end: { ...snap.selectedRange.end },
    },
  };
}

function buildInitialState() {
  const defaultPreset = EXAMPLE_PRESETS[0];
  const blankSheets = [createBlankSheet('Sheet1')];
  const blankSelected: CellRange = {
    start: { row: 1, col: 1 },
    end: { row: 1, col: 1 },
  };

  try {
    const raw = localStorage.getItem(SESSION_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as SavedSession;
      if (parsed.modules?.length > 0) {
        return {
          activePresetId: parsed.activePresetId || defaultPreset.id,
          sheets: blankSheets,
          activeSheetId: blankSheets[0].id,
          selectedRange: blankSelected,
          modules: parsed.modules,
          activeModuleId: parsed.activeModuleId || parsed.modules[0].id,
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

  return {
    activePresetId: defaultPreset.id,
    sheets: blankSheets,
    activeSheetId: blankSheets[0].id,
    selectedRange: blankSelected,
    modules: [
      {
        id: 'mod-1',
        name: 'Module1.bas',
        code: defaultPreset.code,
      },
    ],
    activeModuleId: 'mod-1',
    executionResult: {
      status: 'idle' as const,
      headline: 'Ready',
      detail: 'Click Run macro to populate and format the blank sheet.',
      durationMs: 0,
      logs: [],
      changedCellsCount: 0,
    },
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
  const [executionResult, setExecutionResult] = useState<ExecutionResult>(
    initial.executionResult
  );

  // VBA-only Undo / Redo history stack
  const [historyState, setHistoryState] = useState<{
    items: WorkbookSnapshot[];
    index: number;
  }>(() => ({
    items: [
      cloneSnapshot({
        sheets: initial.sheets,
        activeSheetId: initial.activeSheetId,
        selectedRange: initial.selectedRange,
      }),
    ],
    index: 0,
  }));

  // Persist active preset and VBA modules locally in browser (spreadsheet always loads blank)
  useEffect(() => {
    try {
      const payload: SavedSession = {
        activePresetId,
        modules,
        activeModuleId,
      };
      localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // Ignore quota errors
    }
  }, [activePresetId, modules, activeModuleId]);

  const activePreset = useMemo(
    () => EXAMPLE_PRESETS.find((p) => p.id === activePresetId) ?? EXAMPLE_PRESETS[0],
    [activePresetId]
  );

  const activeModule = useMemo(
    () => modules.find((m) => m.id === activeModuleId) ?? modules[0],
    [modules, activeModuleId]
  );

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

      setExecutionResult(output.result);

      if (output.result.status === 'success') {
        const nextSnap = cloneSnapshot({
          sheets: output.sheets,
          activeSheetId: output.activeSheetId,
          selectedRange: output.selectedRange,
        });
        setSheets(cloneSheets(nextSnap.sheets));
        setActiveSheetId(nextSnap.activeSheetId);
        setSelectedRange(nextSnap.selectedRange);

        setHistoryState((prev) => {
          const base = prev.items.slice(0, prev.index + 1);
          const nextItems = [...base, nextSnap];
          return {
            items: nextItems,
            index: nextItems.length - 1,
          };
        });
      }
    },
    [activeModule.code, sheets, activeSheetId, selectedRange]
  );

  const canUndo = historyState.index > 0;
  const canRedo = historyState.index < historyState.items.length - 1;

  const handleUndo = useCallback(() => {
    if (historyState.index <= 0) return;
    const prevIdx = historyState.index - 1;
    const snap = cloneSnapshot(historyState.items[prevIdx]);
    setHistoryState((prev) => ({ ...prev, index: prevIdx }));
    setSheets(snap.sheets);
    setActiveSheetId(snap.activeSheetId);
    setSelectedRange(snap.selectedRange);
  }, [historyState]);

  const handleRedo = useCallback(() => {
    if (historyState.index >= historyState.items.length - 1) return;
    const nextIdx = historyState.index + 1;
    const snap = cloneSnapshot(historyState.items[nextIdx]);
    setHistoryState((prev) => ({ ...prev, index: nextIdx }));
    setSheets(snap.sheets);
    setActiveSheetId(snap.activeSheetId);
    setSelectedRange(snap.selectedRange);
  }, [historyState]);

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
    customCode?: string
  ) => {
    const freshSheets = [createBlankSheet('Sheet1')];
    const freshRange: CellRange = {
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

    const blankSnap = cloneSnapshot({
      sheets: freshSheets,
      activeSheetId: freshSheets[0].id,
      selectedRange: freshRange,
    });

    setSheets(blankSnap.sheets);
    setActiveSheetId(blankSnap.activeSheetId);
    setSelectedRange(blankSnap.selectedRange);
    setHistoryState({ items: [blankSnap], index: 0 });
    setExecutionResult({
      status: 'idle',
      headline: 'Ready',
      detail: 'Click Run macro to execute.',
      durationMs: 0,
      logs: [],
      changedCellsCount: 0,
    });

    setActiveTab('playground');
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

      <main className="flex-1 w-full max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-8 py-3 space-y-3">
        {activeTab === 'playground' && (
          <>
            {/* Compact Workspace Action Bar */}
            <section className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <label
                  htmlFor="example-select"
                  className="text-[11px] font-semibold uppercase tracking-wider text-[#6B7E73] dark:text-[#849B8E] whitespace-nowrap"
                >
                  Try an example
                </label>
                <select
                  id="example-select"
                  value={activePresetId}
                  onChange={(e) => {
                    const chosen = EXAMPLE_PRESETS.find((p) => p.id === e.target.value);
                    if (chosen) handleLoadPreset(chosen);
                  }}
                  className="h-9 min-w-[220px] px-3 rounded-lg border border-[#D4DFD8] dark:border-[#243B2E] bg-white dark:bg-[#14241B] text-xs font-medium text-[#111C16] dark:text-[#E8F1EC] focus:outline-none focus:ring-2 focus:ring-[#157347] cursor-pointer"
                >
                  {EXAMPLE_PRESETS.map((preset) => (
                    <option key={preset.id} value={preset.id}>
                      {preset.title}
                    </option>
                  ))}
                </select>
              </div>

              <div className="flex items-center gap-2.5 shrink-0">
                <span className="hidden sm:inline text-[11px] text-[#73847A] dark:text-[#789082] whitespace-nowrap">
                  ⌘ Enter to run
                </span>
                <button
                  type="button"
                  onClick={() => handleRunMacro()}
                  className="h-9 px-4 rounded-lg bg-[#187E4B] hover:bg-[#146B40] active:scale-[0.99] text-white text-xs font-semibold inline-flex items-center gap-2 shadow-xs transition-all whitespace-nowrap cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-2 focus-visible:ring-[#187E4B]"
                >
                  <Play className="w-3.5 h-3.5 fill-current" />
                  <span>Run macro</span>
                </button>
              </div>
            </section>

            {/* Dual-Pane Resizable Workspace: Fits laptop viewport without outer page scrolling */}
            <section
              ref={splitContainerRef}
              style={{ '--split-left': `${splitPercent}%` } as React.CSSProperties}
              className={`flex flex-col min-[900px]:flex-row items-stretch gap-4 min-[900px]:gap-0 min-[900px]:h-[clamp(480px,calc(100dvh-202px),740px)] ${
                isResizing ? 'select-none' : ''
              }`}
            >
              <div className="w-full min-[900px]:w-[calc(var(--split-left)-10px)] min-[900px]:h-full min-[900px]:shrink-0 min-w-0">
                <SpreadsheetPane
                  sheets={sheets}
                  activeSheetId={activeSheetId}
                  selectedRange={selectedRange}
                  onSelectSheet={setActiveSheetId}
                  onSelectRange={setSelectedRange}
                  canUndo={canUndo}
                  canRedo={canRedo}
                  historyIndex={historyState.index}
                  historyLength={historyState.items.length}
                  onUndo={handleUndo}
                  onRedo={handleRedo}
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

              <div className="w-full min-[900px]:flex-1 min-[900px]:h-full min-w-0">
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
            <section className="rounded-xl border border-[#D8E7DD] dark:border-[#21382B] bg-[#EDF5F0] dark:bg-[#14251C] px-4 py-2.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
              <div className="flex items-start sm:items-center gap-3">
                <div className="w-6 h-6 rounded-md bg-[#DCECE2] dark:bg-[#1E3829] text-[#157347] dark:text-[#57C587] flex items-center justify-center shrink-0">
                  <Info className="w-3.5 h-3.5" />
                </div>
                <div className="text-xs leading-snug">
                  <span className="font-semibold text-[#111C16] dark:text-[#E8F1EC] mr-2">
                    {activePreset.tipTitle}:
                  </span>
                  <span className="text-[#4B6155] dark:text-[#9CB3A6]">
                    {activePreset.tipBody}
                  </span>
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
            onLoadPreset={(preset) => handleLoadPreset(preset)}
          />
        )}

        {activeTab === 'learn' && (
          <LearnView
            onTryInPlayground={(preset, customCode) =>
              handleLoadPreset(preset, customCode)
            }
          />
        )}
      </main>

      {/* Footer */}
      <footer className="w-full border-t border-[#E2EAE5] dark:border-[#1E3327] mt-auto">
        <div className="max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-8 py-2.5 flex flex-col sm:flex-row items-center justify-between gap-2 text-[11px] text-[#6B7E73] dark:text-[#7E9689]">
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
