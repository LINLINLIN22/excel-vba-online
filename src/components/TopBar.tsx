import { ArrowUpRight, Moon, Sun } from 'lucide-react';

export type ActiveTab = 'playground' | 'examples' | 'learn';

interface TopBarProps {
  activeTab: ActiveTab;
  onChangeTab: (tab: ActiveTab) => void;
  theme: 'light' | 'dark';
  onToggleTheme: (theme: 'light' | 'dark') => void;
  onOpenQuickGuide: () => void;
}

export function TopBar({
  activeTab,
  onChangeTab,
  theme,
  onToggleTheme,
  onOpenQuickGuide,
}: TopBarProps) {
  const navItems: Array<{ id: ActiveTab; label: string }> = [
    { id: 'playground', label: 'Playground' },
    { id: 'examples', label: 'Examples' },
    { id: 'learn', label: 'Learn VBA' },
  ];

  return (
    <header className="w-full border-b border-[#DCE5DF] dark:border-[#1E3327] bg-white/95 dark:bg-[#0F1E16]/95 backdrop-blur-sm transition-colors">
      <div className="max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4 sm:gap-8">
        {/* Brand */}
        <div className="flex items-center gap-2.5 whitespace-nowrap shrink-0">
          <button
            type="button"
            onClick={() => onChangeTab('playground')}
            className="text-lg font-bold tracking-tight text-[#157347] dark:text-[#39A96B] hover:opacity-90 transition-opacity focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#157347] rounded"
          >
            練練VBA
          </button>
          <span className="text-[#B8C7BE] dark:text-[#334F3E] select-none">/</span>
          <span className="text-sm font-medium text-[#526359] dark:text-[#9CB3A6]">
            VBA playground
          </span>
        </div>

        {/* Navigation Links */}
        <nav
          aria-label="Primary navigation"
          className="flex items-center gap-1 sm:gap-6 h-full"
        >
          {navItems.map((item) => {
            const isActive = activeTab === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => onChangeTab(item.id)}
                className={`relative h-full px-2 sm:px-1 flex items-center text-sm font-medium whitespace-nowrap shrink-0 transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#157347] ${
                  isActive
                    ? 'text-[#157347] dark:text-[#57C587]'
                    : 'text-[#526359] dark:text-[#8FA699] hover:text-[#111C16] dark:hover:text-[#E8F1EC]'
                }`}
              >
                {item.label}
                {isActive && (
                  <span
                    aria-hidden="true"
                    className="absolute bottom-0 left-0 right-0 h-[2px] bg-[#157347] dark:bg-[#39A96B] rounded-t"
                  />
                )}
              </button>
            );
          })}
        </nav>

        {/* Right Utility Controls */}
        <div className="flex items-center gap-3 shrink-0">
          <div className="hidden lg:flex items-center gap-2 text-xs text-[#526359] dark:text-[#8FA699] whitespace-nowrap">
            <span
              aria-hidden="true"
              className="w-2 h-2 rounded-full bg-[#1B8750] dark:bg-[#39A96B]"
            />
            <span>Runs in your browser</span>
          </div>

          <button
            type="button"
            onClick={onOpenQuickGuide}
            className="hidden sm:inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium rounded-lg border border-[#DCE5DF] dark:border-[#243B2E] bg-[#F4F7F5] dark:bg-[#172A20] text-[#1D2D24] dark:text-[#DCE9E1] hover:bg-[#E9F0EC] dark:hover:bg-[#1E3529] transition-colors whitespace-nowrap shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#157347]"
          >
            <span>Quick guide</span>
            <ArrowUpRight className="w-3.5 h-3.5 opacity-75" />
          </button>

          {/* Light / Dark Segmented Toggle */}
          <div
            role="group"
            aria-label="Color theme"
            className="inline-flex items-center p-0.5 rounded-lg border border-[#DCE5DF] dark:border-[#243B2E] bg-[#EDF2EF] dark:bg-[#15261D]"
          >
            <button
              type="button"
              onClick={() => onToggleTheme('light')}
              aria-pressed={theme === 'light'}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#157347] ${
                theme === 'light'
                  ? 'bg-white text-[#111C16] shadow-xs'
                  : 'text-[#8FA699] hover:text-[#E8F1EC]'
              }`}
            >
              <Sun className="w-3.5 h-3.5 sm:hidden" />
              <span>Light</span>
            </button>
            <button
              type="button"
              onClick={() => onToggleTheme('dark')}
              aria-pressed={theme === 'dark'}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#157347] ${
                theme === 'dark'
                  ? 'bg-[#213B2C] text-[#E8F1EC] shadow-xs'
                  : 'text-[#526359] hover:text-[#111C16]'
              }`}
            >
              <Moon className="w-3.5 h-3.5 sm:hidden" />
              <span>Dark</span>
            </button>
          </div>
        </div>
      </div>
    </header>
  );
}
