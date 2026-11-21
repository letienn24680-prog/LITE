import type { LucideIcon } from 'lucide-react';
import {
  BookOpen,
  FileSpreadsheet,
  FileText,
  Folder,
  HelpCircle,
  Home,
  LogIn,
  LogOut,
  Presentation,
  Settings,
  Share2,
  User,
} from 'lucide-react';
import type { CloudUserSummary, SidebarNavKey, SyncStatus } from '../../types/document';

export interface SidebarRailProps {
  activeNav: SidebarNavKey;
  onSelectNav: (nav: SidebarNavKey) => void;
  currentUser?: CloudUserSummary | null;
  syncStatus?: SyncStatus;
  onOpenSettings?: () => void;
  onOpenHelp?: () => void;
  onAuthAction?: () => void;
}

interface NavItemConfig {
  key: SidebarNavKey;
  label: string;
  icon: LucideIcon;
}

const NAV_ITEMS: readonly NavItemConfig[] = [
  { key: 'home', label: 'Trang chủ', icon: Home },
  { key: 'slide', label: 'Slide', icon: Presentation },
  { key: 'word', label: 'Word', icon: FileText },
  { key: 'excel', label: 'Excel', icon: FileSpreadsheet },
  { key: 'pdf', label: 'PDF', icon: BookOpen },
  { key: 'folders', label: 'Thư mục', icon: Folder },
  { key: 'shared', label: 'Chia sẻ', icon: Share2 },
];

function getUserInitial(user: CloudUserSummary): string {
  const source = (user.displayName || user.email || 'U').trim();
  return source.charAt(0).toUpperCase() || 'U';
}

export function SidebarRail({
  activeNav,
  onSelectNav,
  currentUser = null,
  syncStatus = 'local',
  onOpenSettings,
  onOpenHelp,
  onAuthAction,
}: SidebarRailProps) {
  return (
    <aside
      aria-label="Điều hướng chính"
      className="hidden h-full w-[72px] shrink-0 select-none flex-col items-center overflow-y-auto no-scrollbar border-r border-[#EAE6DF] bg-[#FFFFFF] pt-4 md:flex"
    >
      {/* Logo monogram LITE */}
      <button
        type="button"
        onClick={() => onSelectNav('home')}
        title="LITE — Trang chủ"
        aria-label="LITE — Trang chủ"
        className="group relative flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-indigo-600 text-white shadow-sm shadow-indigo-600/25 transition-all duration-150 hover:bg-indigo-700 active:scale-[0.97] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
      >
        <svg
          width="20"
          height="20"
          viewBox="0 0 24 24"
          fill="none"
          aria-hidden="true"
          className="transition-transform duration-150 group-hover:scale-[1.04]"
        >
          <path
            d="M6.5 4.25H13.75L17.75 8.25V19.25C17.75 19.8023 17.3023 20.25 16.75 20.25H6.5C5.94772 20.25 5.5 19.8023 5.5 19.25V5.25C5.5 4.69772 5.94772 4.25 6.5 4.25Z"
            stroke="currentColor"
            strokeWidth="1.7"
            strokeLinejoin="round"
            className="opacity-45"
          />
          <path
            d="M13.5 4.5V8.5H17.5"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
            strokeLinejoin="round"
            className="opacity-55"
          />
          <path
            d="M9.25 9.25V16.25H14.5"
            stroke="currentColor"
            strokeWidth="2.35"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>

      {/* Primary navigation */}
      <nav
        aria-label="Danh mục tài liệu"
        className="mt-5 flex w-full flex-1 flex-col items-center gap-1 pb-3"
      >
        {NAV_ITEMS.map((item) => {
          const Icon = item.icon;
          const isActive = activeNav === item.key;

          return (
            <div key={item.key} className="relative flex w-full shrink-0 justify-center">
              {isActive && (
                <span
                  aria-hidden="true"
                  className="absolute left-0 top-1/2 h-7 w-[3px] -translate-y-1/2 rounded-r-full bg-[#4F46E5]"
                />
              )}

              <button
                type="button"
                onClick={() => onSelectNav(item.key)}
                title={item.label}
                aria-current={isActive ? 'page' : undefined}
                className={`group flex w-[58px] flex-col items-center justify-center gap-1 rounded-xl px-1 py-2 transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40 ${
                  isActive
                    ? 'bg-indigo-50 text-[#4F46E5]'
                    : 'text-slate-500 hover:bg-[#F6F3EE] hover:text-slate-800'
                }`}
              >
                <Icon
                  className="h-[19px] w-[19px] shrink-0 transition-transform duration-150 group-hover:scale-[1.04]"
                  strokeWidth={isActive ? 2.15 : 1.75}
                />
                <span className="text-[11px] font-medium leading-none tracking-[-0.01em]">
                  {item.label}
                </span>
              </button>
            </div>
          );
        })}
      </nav>

      {/* Account avatar/status + 3 pinned utility actions */}
      <div className="mt-auto flex w-full shrink-0 flex-col items-center gap-2.5 border-t border-[#F0ECE4] pb-4 pt-3">
        {currentUser ? (
          <button
            type="button"
            onClick={onOpenSettings}
            title={`Đã đăng nhập: ${currentUser.email}`}
            aria-label={`Đã đăng nhập: ${currentUser.email}`}
            className="group relative flex items-center justify-center rounded-full p-0.5 transition-transform duration-150 hover:scale-[1.04] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
          >
            <span className="relative flex h-8 w-8 items-center justify-center rounded-full bg-indigo-600 text-[11px] font-bold text-white shadow-sm shadow-indigo-600/25">
              {getUserInitial(currentUser)}
              <span
                aria-hidden="true"
                className="absolute -bottom-0.5 -right-0.5 h-2.5 w-2.5 rounded-full bg-emerald-500 ring-2 ring-white"
              />
            </span>
          </button>
        ) : (
          <button
            type="button"
            onClick={onAuthAction}
            title="Chế độ Khách (Lưu cục bộ)"
            aria-label="Chế độ Khách (Lưu cục bộ)"
            className="group flex flex-col items-center gap-1 rounded-xl px-1.5 py-0.5 transition-colors hover:bg-[#F6F3EE] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
          >
            <span className="flex h-8 w-8 items-center justify-center rounded-full bg-[#F4F0E8] text-slate-500 transition-colors group-hover:bg-indigo-50 group-hover:text-indigo-600">
              <User className="h-4 w-4" aria-hidden="true" />
            </span>
            <span className="text-[9.5px] font-medium leading-none text-slate-400">
              Khách
            </span>
          </button>
        )}

        <div className="flex flex-col items-center gap-1">
          <button
            type="button"
            onClick={onOpenSettings}
            title={
              syncStatus === 'syncing'
                ? 'Đang đồng bộ — Cài đặt & Đồng bộ'
                : 'Cài đặt & Đồng bộ'
            }
            aria-label="Cài đặt & Đồng bộ"
            className="relative flex h-9 w-9 items-center justify-center rounded-xl text-slate-500 transition-colors duration-150 hover:bg-[#F6F3EE] hover:text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
          >
            <Settings
              className={`h-[18px] w-[18px] ${
                syncStatus === 'syncing' ? 'animate-spin [animation-duration:3s] text-indigo-600' : ''
              }`}
              strokeWidth={1.8}
              aria-hidden="true"
            />
            {syncStatus === 'syncing' && (
              <span
                aria-hidden="true"
                className="absolute right-1.5 top-1.5 h-2 w-2 rounded-full bg-indigo-500 ring-2 ring-white animate-pulse"
              />
            )}
            {syncStatus === 'error' && (
              <span
                aria-hidden="true"
                className="absolute right-1.5 top-1.5 h-2 w-2 rounded-full bg-rose-500 ring-2 ring-white"
              />
            )}
          </button>

          <button
            type="button"
            onClick={onOpenHelp}
            title="Trợ giúp & Hướng dẫn"
            aria-label="Trợ giúp & Hướng dẫn"
            className="flex h-9 w-9 items-center justify-center rounded-xl text-slate-500 transition-colors duration-150 hover:bg-[#F6F3EE] hover:text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
          >
            <HelpCircle className="h-[18px] w-[18px]" strokeWidth={1.8} aria-hidden="true" />
          </button>

          {currentUser ? (
            <button
              type="button"
              onClick={onAuthAction}
              title="Đăng xuất tài khoản"
              aria-label="Đăng xuất tài khoản"
              className="flex h-9 w-9 items-center justify-center rounded-xl text-slate-500 transition-colors duration-150 hover:bg-rose-50 hover:text-rose-600 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-rose-500/35"
            >
              <LogOut className="h-[18px] w-[18px]" strokeWidth={1.8} aria-hidden="true" />
            </button>
          ) : (
            <button
              type="button"
              onClick={onAuthAction}
              title="Đăng nhập / Đăng ký để đồng bộ"
              aria-label="Đăng nhập / Đăng ký để đồng bộ"
              className="flex h-9 w-9 items-center justify-center rounded-xl text-indigo-600 transition-colors duration-150 hover:bg-indigo-50 hover:text-indigo-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
            >
              <LogIn className="h-[18px] w-[18px]" strokeWidth={1.85} aria-hidden="true" />
            </button>
          )}
        </div>
      </div>
    </aside>
  );
}
