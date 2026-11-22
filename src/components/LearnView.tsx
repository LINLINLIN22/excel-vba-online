import { ArrowUpRight, Check, Play } from 'lucide-react';
import { createBlankSheet, EXAMPLE_PRESETS } from '../data/examples';
import { LEARN_LESSONS } from '../data/lessons';
import { ExamplePreset } from '../types/workbook';

interface LearnViewProps {
  onTryInPlayground: (preset: ExamplePreset, customCode?: string) => void;
}

export function LearnView({ onTryInPlayground }: LearnViewProps) {
  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="space-y-1.5">
        <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-[#111C16] dark:text-[#E8F1EC] text-balance">
          循序漸進學 Excel VBA
        </h1>
        <p className="text-sm text-[#526359] dark:text-[#9CB3A6] max-w-2xl">
          從空白的 Excel 工作表出發，透過實際的表格操作掌握 VBA 核心語法。點擊任一單元的按鈕即可直接在即時編輯器中動手練習。
        </p>
      </div>

      {/* Lessons List */}
      <div className="space-y-4">
        {LEARN_LESSONS.map((lesson) => {
          const basePreset =
            EXAMPLE_PRESETS.find((p) => p.id === lesson.presetId) ?? EXAMPLE_PRESETS[0];

          // 確保點擊「在 Playground 試寫」時一律以空白工作表為對象執行教學片段
          const blankTargetPreset: ExamplePreset = {
            ...basePreset,
            sheets: [createBlankSheet('Sheet1')],
            selectedRange: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
          };

          return (
            <section
              key={lesson.id}
              className="rounded-xl border border-[#DCE5DF] dark:border-[#233A2C] bg-white dark:bg-[#14241B] p-5 sm:p-6"
            >
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                {/* Left Explanation */}
                <div className="lg:col-span-6 space-y-3">
                  <div className="text-xs font-medium text-[#157347] dark:text-[#57C587] tabular-nums">
                    單元 {lesson.number}
                  </div>
                  <h2 className="text-lg font-semibold text-[#111C16] dark:text-[#E8F1EC]">
                    {lesson.title}
                  </h2>
                  <p className="text-sm text-[#3F5147] dark:text-[#B3C8BC] leading-relaxed">
                    {lesson.summary}
                  </p>
                  <p className="text-xs text-[#5F7267] dark:text-[#8CA396] leading-relaxed">
                    {lesson.concept}
                  </p>

                  <ul className="space-y-1.5 pt-1">
                    {lesson.keyTakeaways.map((tip) => (
                      <li
                        key={tip}
                        className="flex items-start gap-2 text-xs text-[#2B3D33] dark:text-[#CBE0D3]"
                      >
                        <Check className="w-3.5 h-3.5 text-[#157347] dark:text-[#4ADE80] shrink-0 mt-0.5" />
                        <span>{tip}</span>
                      </li>
                    ))}
                  </ul>
                </div>

                {/* Right Code Snippet + Action */}
                <div className="lg:col-span-6 space-y-3">
                  <pre className="p-4 rounded-lg bg-[#102118] dark:bg-[#0B1710] border border-[#1E3528] text-xs font-mono leading-relaxed text-[#D8EDE1] overflow-x-auto vba-scroll">
                    <code>{lesson.code}</code>
                  </pre>

                  <div className="flex flex-wrap items-center justify-end gap-2.5">
                    <button
                      type="button"
                      onClick={() => onTryInPlayground(basePreset)}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg border border-[#DCE5DF] dark:border-[#243B2E] bg-[#F6F9F7] dark:bg-[#182B20] hover:bg-[#EAF2ED] dark:hover:bg-[#20382A] text-xs font-medium text-[#111C16] dark:text-[#E8F1EC] transition-colors"
                    >
                      <span>開啟完整範例</span>
                      <ArrowUpRight className="w-3.5 h-3.5" />
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        onTryInPlayground(
                          blankTargetPreset,
                          lesson.code.includes('Sub ')
                            ? lesson.code
                            : `Sub PracticeLesson()\n${lesson.code
                                .split('\n')
                                .map((l) => `    ${l}`)
                                .join('\n')}\nEnd Sub`
                        )
                      }
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[#157347] hover:bg-[#12633B] text-white text-xs font-medium transition-colors"
                    >
                      <Play className="w-3 h-3 fill-current" />
                      <span>在空白表格試寫</span>
                    </button>
                  </div>
                </div>
              </div>
            </section>
          );
        })}
      </div>
    </div>
  );
}
