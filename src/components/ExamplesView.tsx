import { ArrowRight, Code2 } from 'lucide-react';
import { useMemo, useState } from 'react';
import { EXAMPLE_PRESETS } from '../data/examples';
import { ExamplePreset } from '../types/workbook';

interface ExamplesViewProps {
  activePresetId: string;
  onLoadPreset: (preset: ExamplePreset, autoRun: boolean) => void;
}

export function ExamplesView({
  activePresetId,
  onLoadPreset,
}: ExamplesViewProps) {
  const [selectedCategory, setSelectedCategory] = useState<string>('全部');

  const categories = useMemo(
    () => ['全部', ...Array.from(new Set(EXAMPLE_PRESETS.map((p) => p.category)))],
    []
  );

  const filtered = useMemo(
    () =>
      selectedCategory === '全部'
        ? EXAMPLE_PRESETS
        : EXAMPLE_PRESETS.filter((p) => p.category === selectedCategory),
    [selectedCategory]
  );

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4">
        <div className="space-y-1.5">
          <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-[#111C16] dark:text-[#E8F1EC] text-balance">
            即開即用的 VBA 範例庫
          </h1>
          <p className="text-sm text-[#526359] dark:text-[#9CB3A6] max-w-2xl">
            挑選常見的 Excel 自動化應用情境，將巨集程式碼載入 Playground 的空白工作表後，按下「Run macro」即可導入資料並執行。
          </p>
        </div>

        {/* Interactive Category Filter */}
        <div
          role="group"
          aria-label="依分類篩選範例"
          className="flex flex-wrap items-center gap-1.5 p-1 rounded-lg border border-[#DCE5DF] dark:border-[#233A2C] bg-white dark:bg-[#14241B]"
        >
          {categories.map((cat) => {
            const active = selectedCategory === cat;
            return (
              <button
                key={cat}
                type="button"
                onClick={() => setSelectedCategory(cat)}
                className={`px-3 py-1.5 text-xs font-medium rounded-md whitespace-nowrap transition-colors ${
                  active
                    ? 'bg-[#157347] text-white'
                    : 'text-[#526359] dark:text-[#9CB3A6] hover:text-[#111C16] dark:hover:text-[#E8F1EC]'
                }`}
              >
                {cat}
              </button>
            );
          })}
        </div>
      </div>

      {/* Examples Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
        {filtered.map((preset) => {
          const isCurrent = preset.id === activePresetId;
          const previewLines = preset.code.split('\n').slice(0, 6).join('\n');

          return (
            <article
              key={preset.id}
              className={`flex flex-col justify-between rounded-xl border bg-white dark:bg-[#14241B] p-5 transition-colors ${
                isCurrent
                  ? 'border-[#157347] dark:border-[#39A96B]'
                  : 'border-[#DCE5DF] dark:border-[#233A2C] hover:border-[#9EC4AE] dark:hover:border-[#345B44]'
              }`}
            >
              <div className="space-y-3">
                {/* Unboxed metadata with typographic separator */}
                <div className="flex items-center gap-2 text-xs text-[#526359] dark:text-[#8FA699]">
                  <span className="font-medium text-[#157347] dark:text-[#57C587]">
                    {preset.category}
                  </span>
                  <span aria-hidden="true">·</span>
                  <span>{preset.difficulty}</span>
                  {isCurrent && (
                    <>
                      <span aria-hidden="true">·</span>
                      <span className="text-[#157347] dark:text-[#6EE7A0] font-medium">
                        目前使用中
                      </span>
                    </>
                  )}
                </div>

                <div>
                  <h2 className="text-base font-semibold text-[#111C16] dark:text-[#E8F1EC]">
                    {preset.title}
                  </h2>
                  <p className="text-xs text-[#526359] dark:text-[#9CB3A6] mt-1 leading-relaxed">
                    {preset.description}
                  </p>
                </div>

                {/* Compact Code Preview */}
                <pre className="p-3 rounded-lg bg-[#102118] dark:bg-[#0B1710] border border-[#1E3528] text-[11px] font-mono leading-relaxed text-[#CBE5D7] overflow-x-auto vba-scroll">
                  <code>{previewLines}</code>
                </pre>
              </div>

              <div className="flex items-center justify-end gap-2 pt-4 mt-4 border-t border-[#EBF1ED] dark:border-[#1E3327]">
                <button
                  type="button"
                  onClick={() => onLoadPreset(preset, false)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-[#157347] hover:bg-[#12633B] text-white text-xs font-medium transition-colors whitespace-nowrap cursor-pointer"
                >
                  <Code2 className="w-3.5 h-3.5" />
                  <span>載入到 Playground</span>
                  <ArrowRight className="w-3 h-3" />
                </button>
              </div>
            </article>
          );
        })}
      </div>
    </div>
  );
}
