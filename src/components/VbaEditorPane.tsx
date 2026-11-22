import {
  AlertCircle,
  CheckCircle2,
  Plus,
  Sparkles,
  Terminal,
  X,
} from 'lucide-react';
import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  ExecutionResult,
  VbaDiagnostic,
  VbaModule,
} from '../types/workbook';
import { tokenizeVbaLine, VbaTokenType } from '../vba/lexer';
import {
  getAutocompleteSuggestions,
  QUICK_HINT_CHIPS,
  VbaSuggestion,
} from '../vba/suggestions';

interface VbaEditorPaneProps {
  modules: VbaModule[];
  activeModuleId: string;
  onSelectModule: (id: string) => void;
  onAddModule: () => void;
  onCloseModule: (id: string) => void;
  onChangeCode: (code: string) => void;
  onRunMacro: () => void;
  diagnostics: VbaDiagnostic[];
  executionResult: ExecutionResult;
  runCount: number;
  onClearExecution: () => void;
}

const TOKEN_COLORS: Record<VbaTokenType, string> = {
  keyword: '#5CE09A',
  builtin: '#7DD3FC',
  constant: '#FBBF24',
  string: '#FCD34D',
  number: '#FDA4AF',
  comment: '#6E917E',
  operator: '#9DC4B0',
  punctuation: '#9DC4B0',
  property: '#D1FAE5',
  ident: '#E8F3ED',
  whitespace: 'inherit',
};

export function VbaEditorPane({
  modules,
  activeModuleId,
  onSelectModule,
  onAddModule,
  onCloseModule,
  onChangeCode,
  onRunMacro,
  diagnostics,
  executionResult,
  runCount,
  onClearExecution,
}: VbaEditorPaneProps) {
  const activeModule = useMemo(
    () => modules.find((m) => m.id === activeModuleId) ?? modules[0],
    [modules, activeModuleId]
  );

  const code = activeModule.code;
  const lines = useMemo(() => code.split('\n'), [code]);

  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const preRef = useRef<HTMLPreElement | null>(null);
  const gutterRef = useRef<HTMLDivElement | null>(null);

  const [cursorPos, setCursorPos] = useState(0);
  const [acItems, setAcItems] = useState<VbaSuggestion[]>([]);
  const [acIndex, setAcIndex] = useState(0);
  const [acReplaceStart, setAcReplaceStart] = useState(0);
  const [acOpen, setAcOpen] = useState(false);

  // Calculate Ln, Col from cursorPos
  const { cursorLine, cursorCol } = useMemo(() => {
    const clamped = Math.max(0, Math.min(cursorPos, code.length));
    const slice = code.slice(0, clamped);
    const split = slice.split('\n');
    return {
      cursorLine: split.length,
      cursorCol: split[split.length - 1].length + 1,
    };
  }, [code, cursorPos]);

  // Map diagnostics by line
  const diagByLine = useMemo(() => {
    const map = new Map<number, VbaDiagnostic[]>();
    for (const d of diagnostics) {
      const list = map.get(d.line) ?? [];
      list.push(d);
      map.set(d.line, list);
    }
    return map;
  }, [diagnostics]);

  const errorCount = diagnostics.filter((d) => d.severity === 'error').length;

  const syncScroll = () => {
    const ta = textareaRef.current;
    if (!ta) return;
    if (preRef.current) {
      preRef.current.scrollTop = ta.scrollTop;
      preRef.current.scrollLeft = ta.scrollLeft;
    }
    if (gutterRef.current) {
      gutterRef.current.scrollTop = ta.scrollTop;
    }
  };

  const updateAutocomplete = (nextCode: string, nextCursor: number, force = false) => {
    const res = getAutocompleteSuggestions(nextCode, nextCursor);
    if (res.items.length > 0 && (force || res.prefix.length >= 2 || nextCode[nextCursor - 1] === '.')) {
      setAcItems(res.items);
      setAcReplaceStart(res.replaceStart);
      setAcIndex(0);
      setAcOpen(true);
    } else {
      setAcOpen(false);
    }
  };

  const applySuggestion = (item: VbaSuggestion) => {
    const ta = textareaRef.current;
    const before = code.slice(0, acReplaceStart);
    const after = code.slice(cursorPos);
    const nextCode = `${before}${item.insertText}${after}`;
    const nextPos = before.length + item.insertText.length;

    onChangeCode(nextCode);
    setAcOpen(false);
    requestAnimationFrame(() => {
      if (ta) {
        ta.focus();
        ta.setSelectionRange(nextPos, nextPos);
        setCursorPos(nextPos);
      }
    });
  };

  const insertSnippetAtCursor = (snippet: string) => {
    const ta = textareaRef.current;
    const start = ta?.selectionStart ?? code.length;
    const end = ta?.selectionEnd ?? code.length;
    const before = code.slice(0, start);
    const after = code.slice(end);
    const nextCode = `${before}${snippet}${after}`;
    const nextPos = start + snippet.length;
    onChangeCode(nextCode);
    requestAnimationFrame(() => {
      if (ta) {
        ta.focus();
        ta.setSelectionRange(nextPos, nextPos);
        setCursorPos(nextPos);
      }
    });
  };

  const jumpToLine = (lineNum: number) => {
    const ta = textareaRef.current;
    if (!ta) return;
    const allLines = code.split('\n');
    let offset = 0;
    for (let i = 0; i < Math.min(lineNum - 1, allLines.length); i++) {
      offset += allLines[i].length + 1;
    }
    ta.focus();
    ta.setSelectionRange(offset, offset);
    setCursorPos(offset);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const ta = e.currentTarget;

    // Run macro shortcut: Ctrl+Enter or Cmd+Enter
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      setAcOpen(false);
      onRunMacro();
      return;
    }

    // Trigger autocomplete: Ctrl+Space
    if (e.ctrlKey && e.code === 'Space') {
      e.preventDefault();
      updateAutocomplete(code, ta.selectionStart, true);
      return;
    }

    // Autocomplete keyboard navigation
    if (acOpen && acItems.length > 0) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setAcIndex((prev) => (prev + 1) % acItems.length);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setAcIndex((prev) => (prev - 1 + acItems.length) % acItems.length);
        return;
      }
      if (e.key === 'Tab' || e.key === 'Enter') {
        e.preventDefault();
        applySuggestion(acItems[acIndex]);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setAcOpen(false);
        return;
      }
    }

    // Tab indentation (4 spaces)
    if (e.key === 'Tab') {
      e.preventDefault();
      const start = ta.selectionStart;
      const end = ta.selectionEnd;
      const next = `${code.slice(0, start)}    ${code.slice(end)}`;
      onChangeCode(next);
      requestAnimationFrame(() => {
        ta.setSelectionRange(start + 4, start + 4);
        setCursorPos(start + 4);
      });
      return;
    }

    // Preserve indentation on Enter
    if (e.key === 'Enter') {
      const start = ta.selectionStart;
      const lineStart = code.lastIndexOf('\n', start - 1) + 1;
      const currentLine = code.slice(lineStart, start);
      const indentMatch = /^(\s*)/.exec(currentLine);
      let indent = indentMatch ? indentMatch[1] : '';
      const trimmedLower = currentLine.trim().toLowerCase();
      if (
        trimmedLower.startsWith('sub ') ||
        trimmedLower.startsWith('function ') ||
        trimmedLower.startsWith('with ') ||
        trimmedLower.startsWith('for ') ||
        trimmedLower.endsWith(' then') ||
        trimmedLower.startsWith('do')
      ) {
        indent += '    ';
      }
      if (indent.length > 0) {
        e.preventDefault();
        const end = ta.selectionEnd;
        const next = `${code.slice(0, start)}\n${indent}${code.slice(end)}`;
        const nextPos = start + 1 + indent.length;
        onChangeCode(next);
        requestAnimationFrame(() => {
          ta.setSelectionRange(nextPos, nextPos);
          setCursorPos(nextPos);
        });
      }
    }
  };

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (textareaRef.current && !textareaRef.current.contains(e.target as Node)) {
        // Let popover clicks fire before closing
        setTimeout(() => setAcOpen(false), 120);
      }
    };
    window.addEventListener('mousedown', handleClickOutside);
    return () => window.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Compute approximate popup vertical offset inside editor
  const acTopPx = Math.min(220, Math.max(28, cursorLine * 22 + 8));

  return (
    <div className="h-full flex flex-col rounded-xl border border-[#DCE5DF] dark:border-[#233A2C] bg-white dark:bg-[#14241B] shadow-xs overflow-hidden">
      {/* Editor Card Header */}
      <div className="flex items-center justify-between gap-3 px-4 py-3 border-b border-[#E7EEE9] dark:border-[#21362A]">
        <div>
          <h2 className="text-sm font-semibold text-[#111C16] dark:text-[#E8F1EC]">
            VBA editor
          </h2>
          <p className="text-xs text-[#6B7E73] dark:text-[#89A193]">
            A little code, a visible result.
          </p>
        </div>

        <span className="px-2.5 py-1 rounded-md border border-[#DCE5DF] dark:border-[#294234] bg-[#F7FAF8] dark:bg-[#192D22] text-[11px] font-medium text-[#526359] dark:text-[#9CB3A6]">
          VBA
        </span>
      </div>

      {/* Module Tab Strip (Dark Forest Bar) */}
      <div className="flex items-center justify-between px-3 bg-[#1A2E23] dark:bg-[#16281E] border-b border-[#263F31] text-xs">
        <div className="flex items-center gap-1 overflow-x-auto vba-scroll">
          {modules.map((mod) => {
            const isActive = mod.id === activeModule.id;
            return (
              <div
                key={mod.id}
                className={`group flex items-center gap-2 px-3 py-2 border-r border-[#243B2E] cursor-pointer transition-colors whitespace-nowrap shrink-0 ${
                  isActive
                    ? 'bg-[#102118] dark:bg-[#0D1A12] text-[#E6F2EC]'
                    : 'text-[#89A595] hover:text-[#D5E6DC] hover:bg-[#1E3528]'
                }`}
                onClick={() => onSelectModule(mod.id)}
              >
                <span className="font-mono text-[11px]">{mod.name}</span>
                {modules.length > 1 && (
                  <button
                    type="button"
                    aria-label={`Close ${mod.name}`}
                    onClick={(e) => {
                      e.stopPropagation();
                      onCloseModule(mod.id);
                    }}
                    className="opacity-60 hover:opacity-100 text-[#9CB8A8]"
                  >
                    <X className="w-3 h-3" />
                  </button>
                )}
              </div>
            );
          })}
          <button
            type="button"
            onClick={onAddModule}
            title="New VBA module"
            aria-label="New VBA module"
            className="p-2 text-[#89A595] hover:text-[#E6F2EC] transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
          </button>
        </div>

        <span className="text-[11px] text-[#7D9B8B] whitespace-nowrap pr-1">
          Auto-saved
        </span>
      </div>

      {/* Code Editor Surface (Always Dark Forest Slate for crisp syntax contrast) */}
      <div className="relative flex flex-1 bg-[#102118] dark:bg-[#0B1710] text-[#E8F3ED] min-h-[300px] overflow-hidden">
        {/* Line Numbers Gutter */}
        <div
          ref={gutterRef}
          aria-hidden="true"
          className="w-11 py-3 px-2 text-right font-mono text-xs leading-[22px] text-[#668574] bg-[#0D1C14] dark:bg-[#09140E] border-r border-[#1C3226] select-none overflow-hidden tabular-nums shrink-0"
        >
          {lines.map((_, idx) => {
            const lineNum = idx + 1;
            const lineDiags = diagByLine.get(lineNum);
            const hasErr =
              lineDiags?.some((d) => d.severity === 'error') ||
              executionResult.errorLine === lineNum;
            const isCurrent = cursorLine === lineNum;
            return (
              <div
                key={lineNum}
                title={lineDiags?.map((d) => d.message).join('\n')}
                className={`h-[22px] flex items-center justify-end gap-1 ${
                  hasErr
                    ? 'text-[#FCA5A5] font-semibold'
                    : isCurrent
                      ? 'text-[#B4D4C2]'
                      : ''
                }`}
              >
                {hasErr && (
                  <span className="w-1.5 h-1.5 rounded-full bg-[#EF4444] shrink-0" />
                )}
                <span>{lineNum}</span>
              </div>
            );
          })}
        </div>

        {/* Syntax Overlay + Transparent Editable Textarea */}
        <div className="relative flex-1 overflow-hidden">
          <pre
            ref={preRef}
            aria-hidden="true"
            className="absolute inset-0 m-0 py-3 px-3.5 font-mono text-[12.5px] leading-[22px] whitespace-pre overflow-auto vba-scroll pointer-events-none select-none"
          >
            {lines.map((lineText, idx) => {
              const lineNum = idx + 1;
              const tokens = tokenizeVbaLine(lineText, lineNum);
              const isCurrent = cursorLine === lineNum;
              const hasErr =
                diagByLine.has(lineNum) &&
                diagByLine.get(lineNum)!.some((d) => d.severity === 'error');
              const isRuntimeErr = executionResult.errorLine === lineNum;

              return (
                <div
                  key={lineNum}
                  className={`h-[22px] -mx-3.5 px-3.5 ${
                    hasErr || isRuntimeErr
                      ? 'bg-[#7F1D1D]/30 underline decoration-wavy decoration-[#F87171]'
                      : isCurrent
                        ? 'bg-[#1A3325]/60'
                        : ''
                  }`}
                >
                  {tokens.length === 0 ? (
                    ' '
                  ) : (
                    tokens.map((tok, tIdx) => (
                      <span
                        key={tIdx}
                        style={{
                          color: TOKEN_COLORS[tok.type],
                          fontStyle: tok.type === 'comment' ? 'italic' : undefined,
                        }}
                      >
                        {tok.value}
                      </span>
                    ))
                  )}
                </div>
              );
            })}
          </pre>

          <textarea
            ref={textareaRef}
            aria-label="VBA code editor"
            spellCheck={false}
            autoCapitalize="off"
            autoComplete="off"
            autoCorrect="off"
            value={code}
            onScroll={syncScroll}
            onChange={(e) => {
              const next = e.target.value;
              const pos = e.target.selectionStart;
              onChangeCode(next);
              setCursorPos(pos);
              updateAutocomplete(next, pos);
            }}
            onClick={(e) => setCursorPos(e.currentTarget.selectionStart)}
            onKeyUp={(e) => {
              if (!['ArrowUp', 'ArrowDown', 'Enter', 'Tab', 'Escape'].includes(e.key)) {
                setCursorPos(e.currentTarget.selectionStart);
              }
            }}
            onKeyDown={handleKeyDown}
            className="vba-code-input vba-scroll relative z-10 w-full h-full m-0 py-3 px-3.5 bg-transparent text-transparent caret-[#5CE09A] font-mono text-[12.5px] leading-[22px] whitespace-pre resize-none focus:outline-none overflow-auto selection:bg-[#2A5E42]/60"
          />

          {/* Command Autocomplete Popover */}
          {acOpen && acItems.length > 0 && (
            <div
              role="listbox"
              aria-label="VBA command hints"
              style={{ top: `${acTopPx}px` }}
              className="absolute left-4 z-30 w-80 max-w-[90%] rounded-lg border border-[#2C4B39] bg-[#13261C] shadow-xl overflow-hidden"
            >
              <div className="px-2.5 py-1 border-b border-[#223C2D] flex items-center justify-between text-[10px] text-[#84A592]">
                <span>Command hints</span>
                <span>Tab / Enter to insert</span>
              </div>
              <div className="max-h-48 overflow-y-auto vba-scroll divide-y divide-[#1D3326]">
                {acItems.map((item, idx) => {
                  const selected = idx === acIndex;
                  return (
                    <button
                      key={item.label}
                      type="button"
                      onMouseDown={(e) => {
                        e.preventDefault();
                        applySuggestion(item);
                      }}
                      className={`w-full text-left px-2.5 py-1.5 transition-colors ${
                        selected ? 'bg-[#1F3F2D]' : 'hover:bg-[#193224]'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-mono text-xs font-medium text-[#6EE7A0]">
                          {item.label}
                        </span>
                        <span className="text-[10px] text-[#7D9B8B]">
                          {item.kind}
                        </span>
                      </div>
                      <p className="text-[11px] text-[#A8C2B3] truncate mt-0.5">
                        {item.description}
                      </p>
                    </button>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Quick Command Hints Strip (常見指令快速提示列) */}
      <div className="flex items-center gap-1.5 px-3 py-1.5 bg-[#14271D] dark:bg-[#0F1E16] border-t border-[#1E3528] overflow-x-auto vba-scroll">
        <span className="inline-flex items-center gap-1 text-[10px] font-medium text-[#7DA08D] whitespace-nowrap shrink-0 pr-1">
          <Sparkles className="w-3 h-3 text-[#5CE09A]" />
          <span>Hints:</span>
        </span>
        {QUICK_HINT_CHIPS.map((chip) => (
          <button
            key={chip.label}
            type="button"
            onClick={() => insertSnippetAtCursor(chip.insert)}
            className="px-2 py-0.5 rounded bg-[#1B3326] hover:bg-[#244432] border border-[#284736] text-[11px] font-mono text-[#B7DEC8] hover:text-white whitespace-nowrap shrink-0 transition-colors"
          >
            {chip.label}
          </button>
        ))}
      </div>

      {/* Editor Status Bar */}
      <div className="flex items-center justify-between gap-3 px-4 py-1.5 bg-[#182B21] dark:bg-[#132219] border-t border-[#243B2E] text-[11px] text-[#8DA99A] tabular-nums">
        <div className="flex items-center gap-3">
          <span>
            Ln {cursorLine}, Col {cursorCol}
          </span>
          {errorCount > 0 ? (
            <button
              type="button"
              onClick={() => {
                const first = diagnostics.find((d) => d.severity === 'error');
                if (first) jumpToLine(first.line);
              }}
              className="text-[#FCA5A5] hover:underline font-medium"
            >
              {errorCount} syntax {errorCount === 1 ? 'issue' : 'issues'}
            </button>
          ) : (
            <span className="text-[#6EE7A0]">Syntax OK</span>
          )}
        </div>

        <div className="flex items-center gap-4">
          <span>Spaces: 4</span>
          <span>VBA</span>
        </div>
      </div>

      {/* Execution Panel */}
      <div className="p-4 bg-white dark:bg-[#14241B] border-t border-[#E7EEE9] dark:border-[#21362A] space-y-3">
        <div className="flex items-center justify-between">
          <h3 className="text-xs font-semibold text-[#111C16] dark:text-[#E8F1EC]">
            Execution
          </h3>
          <button
            type="button"
            onClick={onClearExecution}
            className="text-xs text-[#6B7E73] dark:text-[#89A193] hover:text-[#111C16] dark:hover:text-[#E8F1EC] transition-colors"
          >
            Clear
          </button>
        </div>

        {/* Status Card */}
        {executionResult.status === 'success' && (
          <div className="flex items-start gap-3 p-3 rounded-lg border border-[#D4E7DA] dark:border-[#244331] bg-[#F2F8F4] dark:bg-[#182E22]">
            <div className="w-6 h-6 rounded-full bg-[#DFF0E5] dark:bg-[#234833] text-[#157347] dark:text-[#4ADE80] flex items-center justify-center shrink-0 mt-0.5">
              <CheckCircle2 className="w-3.5 h-3.5" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-xs font-semibold text-[#135334] dark:text-[#6EE7A0]">
                {executionResult.headline}
              </p>
              <p className="text-xs text-[#3C6E54] dark:text-[#9AC2AB] mt-0.5">
                {executionResult.detail}
              </p>
            </div>
          </div>
        )}

        {executionResult.status === 'error' && (
          <button
            type="button"
            onClick={() => {
              if (executionResult.errorLine) jumpToLine(executionResult.errorLine);
            }}
            className="w-full text-left flex items-start gap-3 p-3 rounded-lg border border-[#FECACA] dark:border-[#5C2323] bg-[#FEF2F2] dark:bg-[#2A1515] hover:opacity-95 transition-opacity"
          >
            <div className="w-6 h-6 rounded-full bg-[#FEE2E2] dark:bg-[#4B1D1D] text-[#DC2626] dark:text-[#F87171] flex items-center justify-center shrink-0 mt-0.5">
              <AlertCircle className="w-3.5 h-3.5" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-xs font-semibold text-[#991B1B] dark:text-[#FCA5A5]">
                {executionResult.headline}
              </p>
              <p className="text-xs text-[#B91C1C] dark:text-[#F8B4B4] mt-0.5">
                {executionResult.detail}
              </p>
            </div>
          </button>
        )}

        {executionResult.status === 'idle' && (
          <div className="flex items-center gap-2.5 p-3 rounded-lg border border-[#E7EEE9] dark:border-[#21362A] bg-[#FAFDFB] dark:bg-[#111F17] text-xs text-[#5F7267] dark:text-[#8CA396]">
            <Terminal className="w-4 h-4 shrink-0 opacity-70" />
            <span>Ready. Click &ldquo;Run macro&rdquo; or press Ctrl/⌘ + Enter to execute.</span>
          </div>
        )}

        {/* Debug.Print & MsgBox Console Logs */}
        {executionResult.logs.length > 0 && executionResult.status === 'success' && (
          <div className="rounded-lg border border-[#E2EAE5] dark:border-[#22392B] bg-[#F8FBF9] dark:bg-[#0F1D15] px-3 py-2 space-y-1 max-h-24 overflow-y-auto vba-scroll">
            {executionResult.logs.map((entry) => (
              <div
                key={entry.id}
                className="flex items-start gap-2 text-xs font-mono text-[#1D2D24] dark:text-[#CBE0D3]"
              >
                <span className="text-[10px] uppercase tracking-wider px-1.5 py-0.5 rounded bg-[#E6F1EA] dark:bg-[#1E3628] text-[#157347] dark:text-[#6EE7A0] shrink-0">
                  {entry.type === 'msgbox' ? 'MsgBox' : 'Print'}
                </span>
                <span className="break-all">{entry.message}</span>
              </div>
            ))}
          </div>
        )}

        {/* Execution Footer */}
        <div className="flex items-center justify-between text-[11px] text-[#6B7E73] dark:text-[#849B8E] tabular-nums pt-0.5">
          <span>
            {executionResult.status === 'idle'
              ? 'Not yet run'
              : `Completed in ${(executionResult.durationMs / 1000).toFixed(2)}s`}
          </span>
          <span>
            {runCount} {runCount === 1 ? 'macro' : 'macros'} executed
          </span>
        </div>
      </div>
    </div>
  );
}
