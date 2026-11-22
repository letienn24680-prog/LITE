import { useEffect } from 'react';
import {
  CheckCircle2,
  ChevronRight,
  Clock,
  Cloud,
  FolderOpen,
  Laptop,
  Plus,
  Smartphone,
  Upload,
  Users,
  X,
} from 'lucide-react';
import type { CloudUserSummary } from '../../types/document';
import { startGoogleOAuthFlow } from '../../lib/supabase';

export type LocationSourceKey =
  | 'all'
  | 'recent'
  | 'shared'
  | 'cloud'
  | 'device';

export interface LocationPanelProps {
  isOpen: boolean;
  activeLocation: LocationSourceKey;
  currentUser: CloudUserSummary | null;
  hasGoogleToken: boolean;
  counts: {
    recent: number;
    shared: number;
    cloud: number;
    device: number;
  };
  onSelectLocation: (loc: LocationSourceKey) => void;
  onPickDeviceFiles: () => void;
  onBrowseGoogleDrive: () => void;
  onAddGoogleAccount?: () => void;
  onCloseMobile: () => void;
}

function GoogleDriveTricolorIcon({ className }: { className?: string }) {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      aria-hidden="true"
      className={className}
    >
      <path
        d="M15.5 3.5H8.5L15 14.8H22L15.5 3.5Z"
        fill="#FBBC05"
      />
      <path
        d="M8.5 3.5L2 14.8L5.5 20.8L12 9.5L8.5 3.5Z"
        fill="#34A853"
      />
      <path
        d="M5.5 20.8H18.5L22 14.8H9L5.5 20.8Z"
        fill="#4285F4"
      />
    </svg>
  );
}

export function LocationPanel({
  isOpen,
  activeLocation,
  currentUser,
  hasGoogleToken,
  counts,
  onSelectLocation,
  onPickDeviceFiles,
  onBrowseGoogleDrive,
  onAddGoogleAccount,
  onCloseMobile,
}: LocationPanelProps): JSX.Element | null {
  useEffect(() => {
    if (!isOpen) {
      return;
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onCloseMobile();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, onCloseMobile]);

  if (!isOpen) {
    return null;
  }

  const handleChooseLocation = (loc: LocationSourceKey, isMobile: boolean) => {
    onSelectLocation(loc);
    if (isMobile) {
      onCloseMobile();
    }
  };

  const renderPanelContent = (isMobile: boolean) => (
    <div className="flex h-full flex-col overflow-hidden">
      {/* Panel Header */}
      <div className="shrink-0 flex items-center justify-between gap-2 border-b border-[#EAE6DF] px-4 py-3.5">
        <div className="flex items-center gap-2.5 min-w-0">
          <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-indigo-50 text-indigo-600 border border-indigo-100">
            <FolderOpen className="h-4 w-4" aria-hidden="true" />
          </span>
          <div className="min-w-0">
            <h2 className="text-sm font-semibold text-slate-900 truncate">
              Vị trí lưu trữ
            </h2>
            <p className="text-[11px] text-slate-500 truncate">
              Nguồn tài liệu &amp; đám mây
            </p>
          </div>
        </div>

        <div className="flex items-center gap-1.5 shrink-0">
          {activeLocation !== 'all' && (
            <button
              type="button"
              onClick={() => handleChooseLocation('all', isMobile)}
              className="rounded-lg border border-[#E5DFD5] bg-[#FAF8F5] px-2 py-1 text-[11px] font-medium text-slate-600 hover:bg-[#F2EDE5] hover:text-slate-900 transition-colors"
            >
              Tất cả
            </button>
          )}
          {isMobile && (
            <button
              type="button"
              onClick={onCloseMobile}
              aria-label="Đóng bảng vị trí"
              className="flex h-9 w-9 items-center justify-center rounded-xl bg-slate-100 text-slate-600 hover:bg-slate-200 hover:text-slate-900 transition-colors"
            >
              <X className="h-4 w-4" />
            </button>
          )}
        </div>
      </div>

      {/* Scrollable Locations List */}
      <div className="flex-1 min-h-0 overflow-y-auto px-3 py-3 space-y-1.5">
        <p className="px-2 pb-1 text-[10.5px] font-semibold uppercase tracking-wider text-slate-400">
          Nguồn tài liệu nhanh
        </p>

        {/* 1. Gần đây */}
        <button
          type="button"
          onClick={() => handleChooseLocation('recent', isMobile)}
          aria-current={activeLocation === 'recent' ? 'true' : undefined}
          className={`w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-left transition-all min-h-[46px] ${
            activeLocation === 'recent'
              ? 'bg-indigo-50 text-indigo-700 ring-1 ring-indigo-500/25 shadow-sm'
              : 'text-slate-700 hover:bg-[#FAF8F5] hover:text-slate-950'
          }`}
        >
          <span
            className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border ${
              activeLocation === 'recent'
                ? 'bg-indigo-600 border-indigo-600 text-white'
                : 'bg-[#FAF8F5] border-[#EAE5DC] text-slate-600'
            }`}
          >
            <Clock className="h-4 w-4" aria-hidden="true" />
          </span>
          <div className="min-w-0 flex-1">
            <p className="text-[13px] font-semibold leading-tight truncate">
              Gần đây
            </p>
            <p className="text-[11px] text-slate-500 truncate mt-0.5">
              Vừa mở hoặc chỉnh sửa gần nhất
            </p>
          </div>
          <span
            className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-semibold tabular-nums ${
              activeLocation === 'recent'
                ? 'bg-indigo-600/15 text-indigo-700'
                : 'bg-slate-100 text-slate-600'
            }`}
          >
            {counts.recent}
          </span>
        </button>

        {/* 2. Được chia sẻ với tôi */}
        <button
          type="button"
          onClick={() => handleChooseLocation('shared', isMobile)}
          aria-current={activeLocation === 'shared' ? 'true' : undefined}
          className={`w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-left transition-all min-h-[46px] ${
            activeLocation === 'shared'
              ? 'bg-indigo-50 text-indigo-700 ring-1 ring-indigo-500/25 shadow-sm'
              : 'text-slate-700 hover:bg-[#FAF8F5] hover:text-slate-950'
          }`}
        >
          <span
            className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border ${
              activeLocation === 'shared'
                ? 'bg-indigo-600 border-indigo-600 text-white'
                : 'bg-[#FAF8F5] border-[#EAE5DC] text-slate-600'
            }`}
          >
            <Users className="h-4 w-4" aria-hidden="true" />
          </span>
          <div className="min-w-0 flex-1">
            <p className="text-[13px] font-semibold leading-tight truncate">
              Được chia sẻ với tôi
            </p>
            <p className="text-[11px] text-slate-500 truncate mt-0.5">
              Tài liệu đã bật liên kết chia sẻ
            </p>
          </div>
          <span
            className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-semibold tabular-nums ${
              activeLocation === 'shared'
                ? 'bg-indigo-600/15 text-indigo-700'
                : 'bg-slate-100 text-slate-600'
            }`}
          >
            {counts.shared}
          </span>
        </button>

        <div className="my-2 border-t border-[#F0ECE4]" />

        <p className="px-2 pb-1 text-[10.5px] font-semibold uppercase tracking-wider text-slate-400">
          Kho lưu trữ
        </p>

        {/* 3. Google & LITE Cloud - Cá nhân */}
        <button
          type="button"
          onClick={() => handleChooseLocation('cloud', isMobile)}
          aria-current={activeLocation === 'cloud' ? 'true' : undefined}
          className={`w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-left transition-all min-h-[50px] ${
            activeLocation === 'cloud'
              ? 'bg-sky-50 text-sky-900 ring-1 ring-sky-500/30 shadow-sm'
              : 'text-slate-700 hover:bg-[#FAF8F5] hover:text-slate-950'
          }`}
        >
          <span
            className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border ${
              activeLocation === 'cloud'
                ? 'bg-sky-600 border-sky-600 text-white'
                : 'bg-sky-50 border-sky-200/80 text-sky-600'
            }`}
          >
            <Cloud className="h-4 w-4" aria-hidden="true" />
          </span>
          <div className="min-w-0 flex-1">
            <p className="text-[13px] font-semibold leading-tight truncate">
              Google &amp; LITE Cloud - Cá nhân
            </p>
            <p className="text-[11px] text-slate-500 truncate mt-0.5">
              {currentUser?.email ? currentUser.email : 'Chưa kết nối tài khoản'}
            </p>
          </div>
          <span
            className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-semibold tabular-nums ${
              activeLocation === 'cloud'
                ? 'bg-sky-600/15 text-sky-700'
                : 'bg-slate-100 text-slate-600'
            }`}
          >
            {counts.cloud}
          </span>
        </button>

        {/* 4. Thiết bị này + nút Chọn tệp từ máy */}
        <div
          className={`rounded-xl border transition-all ${
            activeLocation === 'device'
              ? 'border-indigo-500/35 bg-indigo-50/70 shadow-sm'
              : 'border-transparent hover:bg-[#FAF8F5]'
          }`}
        >
          <button
            type="button"
            onClick={() => handleChooseLocation('device', isMobile)}
            aria-current={activeLocation === 'device' ? 'true' : undefined}
            className="w-full flex items-center gap-3 px-3 py-2.5 text-left min-h-[48px]"
          >
            <span
              className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border ${
                activeLocation === 'device'
                  ? 'bg-indigo-600 border-indigo-600 text-white'
                  : 'bg-[#FAF8F5] border-[#EAE5DC] text-slate-600'
              }`}
            >
              {isMobile ? (
                <Smartphone className="h-4 w-4" aria-hidden="true" />
              ) : (
                <Laptop className="h-4 w-4" aria-hidden="true" />
              )}
            </span>
            <div className="min-w-0 flex-1">
              <p
                className={`text-[13px] font-semibold leading-tight truncate ${
                  activeLocation === 'device'
                    ? 'text-indigo-800'
                    : 'text-slate-800'
                }`}
              >
                Thiết bị này
              </p>
              <p className="text-[11px] text-slate-500 truncate mt-0.5">
                Tệp lưu cục bộ trên máy
              </p>
            </div>
            <span
              className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-semibold tabular-nums ${
                activeLocation === 'device'
                  ? 'bg-indigo-600/15 text-indigo-700'
                  : 'bg-slate-100 text-slate-600'
              }`}
            >
              {counts.device}
            </span>
          </button>

          <div className="px-3 pb-2.5 pl-[52px]">
            <button
              type="button"
              onClick={() => {
                onPickDeviceFiles();
                if (isMobile) {
                  onCloseMobile();
                }
              }}
              className="inline-flex items-center gap-1.5 rounded-lg border border-[#E2DCD2] bg-white px-2.5 py-1.5 text-[11.5px] font-medium text-slate-700 hover:border-indigo-300 hover:bg-indigo-50/50 hover:text-indigo-700 transition-colors"
            >
              <Upload className="h-3.5 w-3.5 text-indigo-600 shrink-0" />
              <span>Chọn tệp từ máy</span>
            </button>
          </div>
        </div>

        <div className="my-2 border-t border-[#F0ECE4]" />

        <p className="px-2 pb-1 text-[10.5px] font-semibold uppercase tracking-wider text-slate-400">
          Dịch vụ đám mây liên kết
        </p>

        {/* 5. Duyệt Google Drive */}
        <button
          type="button"
          onClick={() => {
            onBrowseGoogleDrive();
            if (isMobile) {
              onCloseMobile();
            }
          }}
          className="group w-full flex items-start gap-3 rounded-xl border border-[#EAE6DF] bg-[#FAF8F5]/70 hover:bg-white hover:border-indigo-300 px-3 py-2.5 text-left transition-all min-h-[56px]"
        >
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-white border border-[#EAE6DF] shadow-sm">
            <GoogleDriveTricolorIcon className="h-5 w-5" />
          </span>
          <div className="min-w-0 flex-1">
            <div className="flex items-center justify-between gap-1.5">
              <p className="text-[13px] font-semibold text-slate-900 truncate group-hover:text-indigo-700 transition-colors">
                Duyệt Google Drive
              </p>
              {hasGoogleToken ? (
                <span className="shrink-0 inline-flex items-center gap-1 rounded-full bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 text-[10px] font-semibold text-emerald-700">
                  <CheckCircle2 className="h-2.5 w-2.5" />
                  Đã liên kết
                </span>
              ) : (
                <ChevronRight className="h-4 w-4 text-slate-400 shrink-0 group-hover:text-indigo-600 transition-colors" />
              )}
            </div>
            <p className="text-[11px] text-slate-500 leading-snug mt-0.5 line-clamp-2">
              Google Drive, tài liệu trực tuyến và hơn thế nữa...
            </p>
          </div>
        </button>

        {/* 6. Thêm vị trí */}
        <button
          type="button"
          onClick={() => {
            if (onAddGoogleAccount) {
              onAddGoogleAccount();
            } else {
              void startGoogleOAuthFlow();
            }
            if (isMobile) {
              onCloseMobile();
            }
          }}
          className="group w-full flex items-start gap-3 rounded-xl border border-dashed border-[#D8D1C6] bg-white hover:bg-[#FAF8F5] hover:border-indigo-400 px-3 py-2.5 text-left transition-all min-h-[52px]"
        >
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-indigo-50 text-indigo-600 border border-indigo-100 group-hover:bg-indigo-600 group-hover:text-white transition-colors">
            <Plus className="h-4 w-4" aria-hidden="true" />
          </span>
          <div className="min-w-0 flex-1">
            <p className="text-[13px] font-semibold text-slate-900 truncate">
              Thêm vị trí
            </p>
            <p className="text-[11px] text-slate-500 leading-snug mt-0.5">
              Kết nối tài khoản Google Drive để đồng bộ tệp
            </p>
          </div>
        </button>
      </div>
    </div>
  );

  return (
    <>
      {/* Desktop Side Panel */}
      <aside
        aria-label="Bảng vị trí lưu trữ"
        className="hidden md:flex md:w-[272px] shrink-0 h-full flex-col border-r border-[#EAE6DF] bg-[#FFFFFF] select-none"
      >
        {renderPanelContent(false)}
      </aside>

      {/* Mobile Bottom Sheet / Slide-over Panel */}
      <div
        role="dialog"
        aria-modal="true"
        aria-label="Chọn vị trí lưu trữ"
        className="fixed inset-0 z-50 md:hidden flex flex-col justify-end"
      >
        <div
          aria-hidden="true"
          onClick={onCloseMobile}
          className="absolute inset-0 bg-slate-950/50 backdrop-blur-[2px]"
        />
        <div className="relative z-10 max-h-[82dvh] w-full rounded-t-3xl bg-[#FFFFFF] border-t border-[#EAE6DF] shadow-2xl overflow-hidden pb-safe flex flex-col">
          <div className="mx-auto mt-2.5 mb-1 h-1 w-10 rounded-full bg-slate-300 shrink-0" />
          <div className="flex-1 min-h-0 overflow-hidden">
            {renderPanelContent(true)}
          </div>
        </div>
      </div>
    </>
  );
}
