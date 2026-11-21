import { useEffect } from 'react';
import {
  CloudUpload,
  FileText,
  Loader2,
  LogOut,
  ShieldCheck,
  X,
} from 'lucide-react';

export interface CloudActionModalProps {
  mode: 'logoutConfirm' | 'migrateLocal' | null;
  userEmail?: string;
  localCandidateCount: number;
  isProcessing?: boolean;
  onClose: () => void;
  onConfirmLogout: () => void;
  onConfirmMigrateLocal: () => void;
  onSkipMigrateLocal: () => void;
}

export function CloudActionModal({
  mode,
  userEmail,
  localCandidateCount,
  isProcessing = false,
  onClose,
  onConfirmLogout,
  onConfirmMigrateLocal,
  onSkipMigrateLocal,
}: CloudActionModalProps) {
  useEffect(() => {
    if (!mode) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !isProcessing) {
        onClose();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [mode, isProcessing, onClose]);

  if (!mode) {
    return null;
  }

  const isLogout = mode === 'logoutConfirm';

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="cloud-action-modal-title"
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm pb-safe"
    >
      <div
        onClick={(event) => {
          if (event.target === event.currentTarget && !isProcessing) {
            onClose();
          }
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        <div className="w-full max-w-md max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
          {/* Header */}
          <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <span
                className={`h-9 w-9 shrink-0 rounded-xl border flex items-center justify-center ${
                  isLogout
                    ? 'bg-rose-50 border-rose-100 text-rose-600'
                    : 'bg-indigo-50 border-indigo-100 text-[#4F46E5]'
                }`}
              >
                {isLogout ? (
                  <LogOut className="h-4 w-4" />
                ) : (
                  <CloudUpload className="h-4 w-4" />
                )}
              </span>
              <div className="min-w-0">
                <h2
                  id="cloud-action-modal-title"
                  className="text-sm font-semibold text-slate-900 truncate"
                >
                  {isLogout
                    ? 'Xác nhận đăng xuất'
                    : 'Đồng bộ tệp cục bộ lên đám mây'}
                </h2>
                {userEmail && (
                  <p className="text-xs text-slate-500 truncate mt-0.5">
                    {userEmail}
                  </p>
                )}
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              disabled={isProcessing}
              aria-label="Đóng hộp thoại"
              className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-white hover:bg-slate-100 border border-[#E6E0D6] text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Body */}
          <div className="flex-1 min-h-0 overflow-y-auto p-4 sm:p-6 space-y-4">
            {isLogout ? (
              <>
                <p className="text-sm text-slate-700 leading-relaxed">
                  Bạn có chắc muốn đăng xuất khỏi tài khoản này không?
                </p>

                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-3.5 flex items-start gap-3">
                  <ShieldCheck className="h-5 w-5 text-emerald-600 shrink-0 mt-0.5" />
                  <p className="text-xs text-slate-600 leading-relaxed">
                    Dữ liệu đã đồng bộ của bạn vẫn được lưu giữ an toàn trên
                    tài khoản đám mây. Sau khi đăng xuất, ứng dụng sẽ chuyển về{' '}
                    <strong className="text-slate-800">
                      Chế độ cục bộ (Local Mode)
                    </strong>{' '}
                    trên thiết bị này.
                  </p>
                </div>

                <div className="flex flex-col-reverse sm:flex-row items-stretch sm:items-center justify-end gap-2.5 pt-2">
                  <button
                    type="button"
                    onClick={onClose}
                    disabled={isProcessing}
                    className="min-h-[42px] px-4 rounded-xl border border-[#DDD6CA] bg-white hover:bg-[#FAF8F5] text-xs font-semibold text-slate-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    Hủy
                  </button>
                  <button
                    type="button"
                    onClick={onConfirmLogout}
                    disabled={isProcessing}
                    className="min-h-[42px] px-4 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold shadow-sm shadow-rose-600/20 flex items-center justify-center gap-2 transition-colors disabled:opacity-60 disabled:cursor-not-allowed"
                  >
                    {isProcessing ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin" />
                        <span>Đang đăng xuất...</span>
                      </>
                    ) : (
                      <>
                        <LogOut className="h-3.5 w-3.5" />
                        <span>Đăng xuất</span>
                      </>
                    )}
                  </button>
                </div>
              </>
            ) : (
              <>
                <p className="text-sm font-medium text-slate-900 leading-relaxed">
                  Bạn có muốn tải các tệp hiện có trên máy này lên tài khoản để
                  đồng bộ không?
                </p>

                <div className="rounded-2xl border border-indigo-100 bg-indigo-50/60 p-3.5 flex items-center gap-3">
                  <span className="h-10 w-10 rounded-xl bg-white border border-indigo-100 text-[#4F46E5] flex items-center justify-center shrink-0">
                    <FileText className="h-5 w-5" />
                  </span>
                  <div className="min-w-0">
                    <p className="text-xs font-semibold text-slate-900">
                      Phát hiện{' '}
                      <span className="text-[#4F46E5] tabular-nums">
                        {localCandidateCount}
                      </span>{' '}
                      tệp cục bộ trên máy này
                    </p>
                    <p className="text-[11.5px] text-slate-600 leading-relaxed mt-0.5">
                      Tải lên để truy cập và chỉnh sửa đồng nhất trên cả máy
                      tính lẫn điện thoại.
                    </p>
                  </div>
                </div>

                <div className="flex flex-col-reverse sm:flex-row items-stretch sm:items-center justify-end gap-2.5 pt-2">
                  <button
                    type="button"
                    onClick={onSkipMigrateLocal}
                    disabled={isProcessing}
                    className="min-h-[42px] px-3.5 rounded-xl border border-[#DDD6CA] bg-white hover:bg-[#FAF8F5] text-xs font-semibold text-slate-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    Bỏ qua, chỉ dùng dữ liệu tài khoản
                  </button>
                  <button
                    type="button"
                    onClick={onConfirmMigrateLocal}
                    disabled={isProcessing}
                    className="min-h-[42px] px-4 rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm shadow-indigo-600/20 flex items-center justify-center gap-2 transition-colors disabled:opacity-60 disabled:cursor-not-allowed whitespace-nowrap"
                  >
                    {isProcessing ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin" />
                        <span>Đang tải lên &amp; đồng bộ...</span>
                      </>
                    ) : (
                      <>
                        <CloudUpload className="h-4 w-4" />
                        <span>Tải lên &amp; Đồng bộ ngay</span>
                      </>
                    )}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
