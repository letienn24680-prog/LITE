import type { LucideIcon } from 'lucide-react';
import {
  BookOpen,
  FileSpreadsheet,
  FileText,
  Folder,
  Home,
  Presentation,
  Share2,
} from 'lucide-react';
import type { SidebarNavKey } from '../../types/document';

export interface MobileBottomNavProps {
  activeNav: SidebarNavKey;
  onSelectNav: (nav: SidebarNavKey) => void;
}

interface MobileNavItemConfig {
  key: SidebarNavKey;
  label: string;
  icon: LucideIcon;
}

const MOBILE_NAV_ITEMS: readonly MobileNavItemConfig[] = [
  { key: 'home', label: 'Trang chủ', icon: Home },
  { key: 'slide', label: 'Slide', icon: Presentation },
  { key: 'word', label: 'Word', icon: FileText },
  { key: 'excel', label: 'Excel', icon: FileSpreadsheet },
  { key: 'pdf', label: 'PDF', icon: BookOpen },
  { key: 'folders', label: 'Vị trí', icon: Folder },
  { key: 'shared', label: 'Chia sẻ', icon: Share2 },
];

export function MobileBottomNav({ activeNav, onSelectNav }: MobileBottomNavProps) {
  return (
    <nav
      aria-label="Điều hướng di động"
      className="fixed inset-x-0 bottom-0 z-30 select-none border-t border-[#EAE6DF] bg-white/95 backdrop-blur-md pb-safe md:hidden"
    >
      <div className="flex h-16 items-center overflow-x-auto no-scrollbar px-1">
        {MOBILE_NAV_ITEMS.map((item) => {
          const Icon = item.icon;
          const isActive = activeNav === item.key;

          return (
            <button
              key={item.key}
              type="button"
              onClick={() => onSelectNav(item.key)}
              aria-current={isActive ? 'page' : undefined}
              className={`relative flex h-full min-h-[44px] min-w-[64px] flex-1 shrink-0 flex-col items-center justify-center gap-1 rounded-lg px-1.5 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-indigo-500/40 ${
                isActive
                  ? 'text-[#4F46E5]'
                  : 'text-slate-500 hover:text-slate-800 active:text-slate-900'
              }`}
            >
              {isActive && (
                <span
                  aria-hidden="true"
                  className="absolute left-1/2 top-0 h-[2.5px] w-8 -translate-x-1/2 rounded-b-full bg-[#4F46E5]"
                />
              )}

              <Icon
                className="h-[19px] w-[19px] shrink-0"
                strokeWidth={isActive ? 2.2 : 1.75}
                aria-hidden="true"
              />
              <span className="whitespace-nowrap text-[11px] font-medium leading-none tracking-[-0.01em]">
                {item.label}
              </span>
            </button>
          );
        })}
      </div>
    </nav>
  );
}
