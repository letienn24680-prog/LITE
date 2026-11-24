import { useEffect, useState } from 'react';
import {
  ArrowDownToLine,
  CheckCircle2,
  Loader2,
  Rocket,
  Sparkles,
  X,
} from 'lucide-react';
import {
  launchApkUpdateDownload,
  type AppVersionManifest,
} from '../../services/appUpdater';

export interface ApkUpdateModalProps {
  isOpen: boolean;
  currentVersion: string;
  currentBuildNumber: number;
  latestManifest: AppVersionManifest | null;
  onClose: () => void;
}

export function ApkUpdateModal({
  isOpen,
  currentVersion,
  currentBuildNumber,
  latestManifest,
  onClose,
}: ApkUpdateModalProps): JSX.Element | null {
  const [isDownloading, setIsDownloading] = useState(false);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !isDownloading) {
        onClose();
      }
    };
    const handleHardwareBack = (event: Event) => {
      if (!isDownloading) {
        event.preventDefault();
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    window.addEventListener('lite:hardware-back', handleHardwareBack);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('lite:hardware-back', handleHardwareBack);
    };
  }, [isOpen, isDownloading, onClose]);

  if (!isOpen || !latestManifest) {
    return null;
  }

  const handleUpdateClick = async () => {
    if (isDownloading) return;
    setIsDownloading(true);
    try {
      await launchApkUpdateDownload(latestManifest.downloadUrl);
    } finally {
      window.setTimeout(() => {
        setIsDownloading(false);
      }, 1200);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="apk-update-title"
      onClick={() => {
        if (!isDownloading) onClose();
      }}
      className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/60 p-4 backdrop-blur-sm pb-safe"
    >
      <div
        onClick={(event) => event.stopPropagation()}
        className="w-full max-w-md overflow-hidden rounded-3xl border border-[#EAE5DC] bg-white shadow-2xl"
      >
        <div className="bg-gradient-to-br from-[#4F46E5] to-[#185ABD] px-6 py-5 text-white relative">
          <button
            type="button"
            onClick={onClose}
            disabled={isDownloading}
            aria-label="Đóng thông báo cập nhật"
            className="absolute right-4 top-4 flex h-8 w-8 items-center justify-center rounded-full bg-white/15 text-white hover:bg-white/25 disabled:opacity-50"
          >
            <X className="h-4 w-4" />
          </button>

          <div className="flex items-center gap-3">
            <span className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-white/20 text-white shadow-inner">
              <Rocket className="h-5 w-5" aria-hidden="true" />
            </span>
            <div className="min-w-0">
              <span className="inline-flex items-center gap-1 rounded-full bg-white/20 px-2 py-0.5 text-[10.5px] font-semibold uppercase tracking-wide">
                <Sparkles className="h-3 w-3" />
                <span>Phiên bản mới khả dụng</span>
              </span>
              <h2
                id="apk-update-title"
                className="mt-1 text-base font-bold tracking-tight"
              >
                LITE Workspace v{latestManifest.version}
              </h2>
            </div>
          </div>
        </div>

        <div className="p-5 sm:p-6 space-y-4">
          <div className="flex items-center justify-between rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] px-3.5 py-2.5 text-xs">
            <span className="text-slate-600">
              Hiện tại:{' '}
              <strong className="text-slate-900">
                v{currentVersion} (#{currentBuildNumber})
              </strong>
            </span>
            <span className="inline-flex items-center gap-1 font-bold text-emerald-700">
              <CheckCircle2 className="h-3.5 w-3.5" />
              Mới: v{latestManifest.version} (#{latestManifest.buildNumber})
            </span>
          </div>

          <div className="space-y-1.5">
            <p className="text-xs font-bold uppercase tracking-wider text-slate-500">
              Nội dung cập nhật
            </p>
            <p className="rounded-2xl border border-slate-200 bg-slate-50 p-3.5 text-xs leading-relaxed text-slate-700">
              {latestManifest.changelog}
            </p>
          </div>

          <p className="text-[11px] leading-relaxed text-slate-500">
            Bấm <strong>Cập nhật ngay</strong> để tải tệp APK phát hành mới nhất và
            mở trình cài đặt gói của Android. Dữ liệu tài liệu trên máy vẫn được
            giữ nguyên vẹn.
          </p>

          <div className="flex flex-col-reverse gap-2 pt-1 sm:flex-row sm:items-center sm:justify-end">
            <button
              type="button"
              onClick={onClose}
              disabled={isDownloading}
              className="inline-flex min-h-[42px] items-center justify-center rounded-xl border border-[#E2DDD5] bg-white px-4 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50 disabled:opacity-50"
            >
              Để sau
            </button>

            <button
              type="button"
              onClick={() => {
                void handleUpdateClick();
              }}
              disabled={isDownloading}
              className="inline-flex min-h-[42px] items-center justify-center gap-2 rounded-xl bg-[#4F46E5] px-5 py-2 text-xs font-semibold text-white shadow-sm shadow-indigo-600/25 hover:bg-indigo-700 disabled:opacity-70"
            >
              {isDownloading ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" />
                  <span>Đang mở trình tải APK...</span>
                </>
              ) : (
                <>
                  <ArrowDownToLine className="h-4 w-4" />
                  <span>Cập nhật ngay</span>
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
