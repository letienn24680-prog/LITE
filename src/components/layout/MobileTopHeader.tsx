import { HelpCircle, LogIn, LogOut, Settings } from 'lucide-react';
import type { CloudUserSummary, SyncStatus } from '../../types/document';

export interface MobileTopHeaderProps {
  currentUser: CloudUserSummary | null;
  syncStatus: SyncStatus;
  onLogoClick: () => void;
  onOpenHelp: () => void;
  onOpenSettings: () => void;
  onAuthAction: () => void;
}

function getUserInitial(user: CloudUserSummary): string {
  const source = (user.displayName || user.email || 'U').trim();
  return source.charAt(0).toUpperCase() || 'U';
}

function getShortDisplayName(user: CloudUserSummary): string {
  const trimmed = user.displayName.trim();
  if (trimmed.length > 0) {
    return trimmed;
  }
  const localPart = user.email.split('@')[0]?.trim();
  return localPart || user.email;
}

export function MobileTopHeader({
  currentUser,
  syncStatus,
  onLogoClick,
  onOpenHelp,
  onOpenSettings,
  onAuthAction,
}: MobileTopHeaderProps) {
  return (
    <header className="z-20 flex h-14 shrink-0 select-none items-center justify-between gap-2 border-b border-[#EAE6DF] bg-white/95 px-3 backdrop-blur-md md:hidden">
      {/* Left: Logo monogram + LITE + compact status badge */}
      <div className="flex min-w-0 items-center gap-1.5">
        <button
          type="button"
          onClick={onLogoClick}
          aria-label="LITE — Trang chủ"
          className="flex shrink-0 items-center gap-1.5 rounded-xl py-1 pr-0.5 transition-opacity active:opacity-80 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
        >
          <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-indigo-600 text-white shadow-sm shadow-indigo-600/20">
            <svg
              width="17"
              height="17"
              viewBox="0 0 24 24"
              fill="none"
              aria-hidden="true"
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
          </span>
          <span className="text-[13px] font-bold tracking-[0.06em] text-slate-900">
            LITE
          </span>
        </button>

        {currentUser ? (
          <button
            type="button"
            onClick={onOpenSettings}
            title={`Đã đăng nhập: ${currentUser.email}`}
            className="inline-flex min-w-0 max-w-[116px] items-center gap-1.5 rounded-full border border-emerald-200/80 bg-emerald-50/85 px-2 py-0.5 text-[11px] font-medium text-emerald-700 transition-colors hover:bg-emerald-100/80"
          >
            <span
              aria-hidden="true"
              className={`h-1.5 w-1.5 shrink-0 rounded-full ${
                syncStatus === 'syncing'
                  ? 'bg-indigo-500 animate-pulse'
                  : syncStatus === 'error'
                    ? 'bg-rose-500'
                    : 'bg-emerald-500'
              }`}
            />
            <span className="truncate">{getShortDisplayName(currentUser)}</span>
          </button>
        ) : (
          <span className="inline-flex shrink-0 items-center whitespace-nowrap rounded-full border border-[#E6E1D8] bg-[#F6F3EE] px-2 py-0.5 text-[10px] font-medium text-slate-500">
            <span className="hidden min-[390px]:inline">Khách ·&nbsp;</span>Local
          </span>
        )}
      </div>

      {/* Right: Help, Settings, Account / Auth action */}
      <div className="flex shrink-0 items-center gap-1">
        <button
          type="button"
          onClick={onOpenHelp}
          aria-label="Trợ giúp"
          title="Trợ giúp & Hướng dẫn"
          className="flex min-h-[38px] min-w-[38px] items-center justify-center rounded-xl text-slate-500 transition-colors hover:bg-[#F6F3EE] hover:text-slate-800 active:scale-[0.97] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
        >
          <HelpCircle className="h-[18px] w-[18px]" strokeWidth={1.8} aria-hidden="true" />
        </button>

        <button
          type="button"
          onClick={onOpenSettings}
          aria-label="Cài đặt & Đồng bộ"
          title="Cài đặt & Đồng bộ"
          className="relative flex min-h-[38px] min-w-[38px] items-center justify-center rounded-xl text-slate-500 transition-colors hover:bg-[#F6F3EE] hover:text-slate-800 active:scale-[0.97] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
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

        {currentUser ? (
          <button
            type="button"
            onClick={onAuthAction}
            aria-label="Đăng xuất tài khoản"
            title={`Đăng xuất (${currentUser.email})`}
            className="inline-flex min-h-[38px] min-w-[38px] items-center gap-1.5 rounded-xl border border-[#EAE6DF] bg-white pl-1.5 pr-2 text-slate-600 transition-colors hover:border-rose-200 hover:bg-rose-50 hover:text-rose-600 active:scale-[0.97] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-rose-500/35"
          >
            <span className="relative flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-indigo-600 text-[10px] font-bold text-white">
              {getUserInitial(currentUser)}
              <span
                aria-hidden="true"
                className="absolute -bottom-0.5 -right-0.5 h-2 w-2 rounded-full bg-emerald-500 ring-1 ring-white"
              />
            </span>
            <LogOut className="h-3.5 w-3.5 shrink-0" aria-hidden="true" />
          </button>
        ) : (
          <button
            type="button"
            onClick={onAuthAction}
            aria-label="Đăng nhập"
            title="Đăng nhập / Đăng ký để đồng bộ"
            className="inline-flex min-h-[38px] min-w-[38px] items-center gap-1 whitespace-nowrap rounded-xl bg-indigo-50 px-2.5 text-xs font-semibold text-indigo-700 transition-colors hover:bg-indigo-100 active:scale-[0.97] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
          >
            <LogIn className="h-3.5 w-3.5 shrink-0" aria-hidden="true" />
            <span>Đăng nhập</span>
          </button>
        )}
      </div>
    </header>
  );
}
