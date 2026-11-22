import { BookOpen, Keyboard, X } from 'lucide-react';
import { useEffect } from 'react';
import { QUICK_GUIDE_SECTIONS } from '../data/lessons';

interface QuickGuideModalProps {
  open: boolean;
  onClose: () => void;
  onJumpToLearn: () => void;
}

export function QuickGuideModal({
  open,
  onClose,
  onJumpToLearn,
}: QuickGuideModalProps) {
  useEffect(() => {
    if (!open) return;
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, [open, onClose]);

  if (!open) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/45 backdrop-blur-[2px]"
      role="dialog"
      aria-modal="true"
      aria-labelledby="quick-guide-title"
      onClick={onClose}
    >
      <div
        className="w-full max-w-2xl rounded-xl border border-[#DCE5DF] dark:border-[#243B2E] bg-white dark:bg-[#13231A] text-[#111C16] dark:text-[#E8F1EC] shadow-2xl overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-[#E5ECE8] dark:border-[#21362A]">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-[#EAF4EE] dark:bg-[#1B3325] text-[#157347] dark:text-[#4BBF80] flex items-center justify-center">
              <Keyboard className="w-4 h-4" />
            </div>
            <div>
              <h2 id="quick-guide-title" className="text-base font-semibold">
                VBA Quick Reference &amp; Shortcuts
              </h2>
              <p className="text-xs text-[#526359] dark:text-[#9CB3A6]">
                Client-side Excel VBA simulator reference
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close quick guide"
            className="p-2 rounded-lg text-[#526359] dark:text-[#9CB3A6] hover:bg-[#F1F5F2] dark:hover:bg-[#1B3124] transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Content */}
        <div className="p-6 max-h-[70vh] overflow-y-auto vba-scroll space-y-6">
          {QUICK_GUIDE_SECTIONS.map((sec) => (
            <section key={sec.title} className="space-y-2.5">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-[#157347] dark:text-[#57C587]">
                {sec.title}
              </h3>
              <div className="rounded-lg border border-[#E5ECE8] dark:border-[#21362A] divide-y divide-[#E5ECE8] dark:divide-[#21362A]">
                {sec.items.map((item) => (
                  <div
                    key={item.key}
                    className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 sm:gap-4 px-3.5 py-2.5 text-sm"
                  >
                    <code className="font-mono text-xs font-medium text-[#157347] dark:text-[#6EE7A0] bg-[#F1F7F3] dark:bg-[#192F22] px-2 py-0.5 rounded w-fit">
                      {item.key}
                    </code>
                    <span className="text-xs text-[#526359] dark:text-[#A3B9AC]">
                      {item.desc}
                    </span>
                  </div>
                ))}
              </div>
            </section>
          ))}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-6 py-3.5 border-t border-[#E5ECE8] dark:border-[#21362A] bg-[#F7FAF8] dark:bg-[#0F1D15]">
          <span className="text-xs text-[#526359] dark:text-[#8FA699]">
            All macros run locally in your browser without uploading data.
          </span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => {
                onClose();
                onJumpToLearn();
              }}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium rounded-lg bg-[#157347] hover:bg-[#12633B] text-white transition-colors"
            >
              <BookOpen className="w-3.5 h-3.5" />
              <span>Open Learn VBA</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
