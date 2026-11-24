import { useEffect, useState } from 'react';
import {
  AlertCircle,
  BookOpen,
  Check,
  CheckCircle2,
  Cloud,
  CloudOff,
  Eye,
  EyeOff,
  KeyRound,
  Loader2,
  LogIn,
  LogOut,
  Moon,
  RefreshCw,
  Settings,
  Sparkles,
  Sun,
  CheckCheck,
  Copy,
  ShieldCheck,
  Trash2,
  User,
  UserPlus,
  Wifi,
  WifiOff,
  X,
} from 'lucide-react';
import type {
  CloudUserSummary,
  SyncStatus,
  UserSettings,
  WordPaperTheme,
} from '../../types/document';
import { formatRelativeVietnameseTime } from '../../services/cloudSync';
import {
  addOrApproveEmail,
  listAllowedEmails,
  removeAllowedEmail,
  type AllowedEmailRecord,
} from '../../services/accessControl';

export interface SettingsModalProps {
  isOpen: boolean;
  user: CloudUserSummary | null;
  settings: UserSettings;
  syncStatus: SyncStatus;
  cloudDocumentCount: number;
  lastSyncedAt: string | null;
  isSupabaseConfigured: boolean;
  onClose: () => void;
  onUpdateSettings: (patch: Partial<UserSettings>) => void;
  onManualSync: () => Promise<void> | void;
  onOpenAuth: () => void;
  onRequestLogout: () => void;
}

interface WordThemeOption {
  key: WordPaperTheme;
  label: string;
  description: string;
  previewBg: string;
  previewBorder: string;
  linePrimary: string;
  lineSecondary: string;
  icon: typeof Sun;
}

const WORD_THEME_OPTIONS: readonly WordThemeOption[] = [
  {
    key: 'light',
    label: 'Trắng sáng',
    description: 'Sắc nét ban ngày',
    previewBg: 'bg-white',
    previewBorder: 'border-slate-200',
    linePrimary: 'bg-slate-800/80',
    lineSecondary: 'bg-slate-300',
    icon: Sun,
  },
  {
    key: 'sepia',
    label: 'Giấy kem ấm',
    description: 'Dịu mắt đọc lâu',
    previewBg: 'bg-[#F7EFE5]',
    previewBorder: 'border-[#DECDB6]',
    linePrimary: 'bg-[#5A422B]/80',
    lineSecondary: 'bg-[#C9B397]',
    icon: BookOpen,
  },
  {
    key: 'dark',
    label: 'Nền tối',
    description: 'Êm dịu ban đêm',
    previewBg: 'bg-[#181D27]',
    previewBorder: 'border-[#2B3547]',
    linePrimary: 'bg-slate-100/85',
    lineSecondary: 'bg-slate-600',
    icon: Moon,
  },
];

interface SyncStatusMeta {
  label: string;
  badgeClass: string;
  dotClass: string;
}

const SYNC_STATUS_META: Record<SyncStatus, SyncStatusMeta> = {
  local: {
    label: 'Chế độ cục bộ (Local)',
    badgeClass: 'bg-slate-100 text-slate-700 border-slate-200',
    dotClass: 'bg-slate-400',
  },
  idle: {
    label: 'Sẵn sàng đồng bộ',
    badgeClass: 'bg-indigo-50 text-[#4F46E5] border-indigo-200',
    dotClass: 'bg-[#4F46E5]',
  },
  syncing: {
    label: 'Đang đồng bộ...',
    badgeClass: 'bg-amber-50 text-amber-800 border-amber-200',
    dotClass: 'bg-amber-500',
  },
  synced: {
    label: 'Đã đồng bộ',
    badgeClass: 'bg-emerald-50 text-emerald-700 border-emerald-200',
    dotClass: 'bg-emerald-500',
  },
  error: {
    label: 'Lỗi đồng bộ',
    badgeClass: 'bg-rose-50 text-rose-700 border-rose-200',
    dotClass: 'bg-rose-500',
  },
};

function getUserInitial(user: CloudUserSummary): string {
  const source = user.displayName.trim() || user.email.trim();
  return source ? source.charAt(0).toUpperCase() : 'U';
}

export function SettingsModal({
  isOpen,
  user,
  settings,
  syncStatus,
  cloudDocumentCount,
  lastSyncedAt,
  isSupabaseConfigured,
  onClose,
  onUpdateSettings,
  onManualSync,
  onOpenAuth,
  onRequestLogout,
}: SettingsModalProps) {

  const [allowedList, setAllowedList] = useState<AllowedEmailRecord[]>([]);
  const [newAllowedEmail, setNewAllowedEmail] = useState<string>("");
  const [newAllowedNote, setNewAllowedNote] = useState<string>("");
  const [isLoadingAllowed, setIsLoadingAllowed] = useState<boolean>(false);
  const [isSavingAllowed, setIsSavingAllowed] = useState<boolean>(false);
  const [allowedTableMissing, setAllowedTableMissing] = useState<boolean>(false);
  const [allowedFeedback, setAllowedFeedback] = useState<string | null>(null);
  const [copiedSql, setCopiedSql] = useState<boolean>(false);
  const [showGeminiKey, setShowGeminiKey] = useState<boolean>(false);

  const refreshAllowedEmails = async () => {
    if (!isSupabaseConfigured || !user) return;
    setIsLoadingAllowed(true);
    const res = await listAllowedEmails();
    setIsLoadingAllowed(false);
    setAllowedTableMissing(Boolean(res.tableMissing));
    if (!res.error) setAllowedList(res.records);
  };

  useEffect(() => {
    if (isOpen && user && isSupabaseConfigured) {
      void refreshAllowedEmails();
    }
  }, [isOpen, user?.email]);

  const handleAddAllowedEmail = async () => {
    if (!newAllowedEmail.trim() || isSavingAllowed) return;
    setIsSavingAllowed(true);
    setAllowedFeedback(null);
    const res = await addOrApproveEmail(newAllowedEmail, newAllowedNote, user?.email);
    setIsSavingAllowed(false);
    if (res.ok) {
      setNewAllowedEmail("");
      setNewAllowedNote("");
      setAllowedFeedback("Đã ép email vào danh sách! Người này có thể đăng nhập Google ngay mà không cần đóng gói lại app.");
      await refreshAllowedEmails();
    } else {
      setAllowedFeedback("Lỗi: " + (res.error || "Không thể thêm email."));
    }
  };

  const handleApprovePending = async (targetEmail: string, note?: string) => {
    setIsSavingAllowed(true);
    const res = await addOrApproveEmail(targetEmail, note, user?.email);
    setIsSavingAllowed(false);
    if (res.ok) {
      setAllowedFeedback(`Đã duyệt quyền truy cập cho ${targetEmail}!`);
      await refreshAllowedEmails();
    }
  };

  const handleDeleteAllowed = async (targetEmail: string) => {
    setIsSavingAllowed(true);
    const res = await removeAllowedEmail(targetEmail);
    setIsSavingAllowed(false);
    if (res.ok) {
      setAllowedFeedback(`Đã thu hồi quyền của ${targetEmail}.`);
      await refreshAllowedEmails();
    }
  };

  const [isManualSyncRunning, setIsManualSyncRunning] =
    useState<boolean>(false);

  useEffect(() => {
    if (!isOpen) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, onClose]);

  if (!isOpen) {
    return null;
  }

  const isSyncing = syncStatus === 'syncing' || isManualSyncRunning;
  const canManualSync = isSupabaseConfigured && user !== null && !isSyncing;
  const statusMeta = SYNC_STATUS_META[syncStatus];
  const lastSyncText = formatRelativeVietnameseTime(lastSyncedAt);

  const handleManualSyncClick = async () => {
    if (!canManualSync) {
      return;
    }
    setIsManualSyncRunning(true);
    try {
      await Promise.resolve(onManualSync());
    } finally {
      setIsManualSyncRunning(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="settings-modal-title"
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm pb-safe"
    >
      <div
        onClick={(event) => {
          if (event.target === event.currentTarget) {
            onClose();
          }
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        <div className="w-full max-w-xl max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
          {/* Header */}
          <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <span className="h-9 w-9 shrink-0 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-[#4F46E5]">
                <Settings className="h-4 w-4" />
              </span>
              <div className="min-w-0">
                <h2
                  id="settings-modal-title"
                  className="text-sm font-semibold text-slate-900 truncate"
                >
                  Cài đặt &amp; Đồng bộ tài khoản
                </h2>
                <p className="text-xs text-slate-500 truncate mt-0.5">
                  Quản lý tài khoản đám mây, giao diện đọc Word và trạng thái
                  lưu trữ
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng cài đặt"
              className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-white hover:bg-slate-100 border border-[#E6E0D6] text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Body */}
          <div className="flex-1 min-h-0 overflow-y-auto">
            <div className="p-4 sm:p-6 space-y-6">
              {/* 1. Account Section */}
              <section aria-labelledby="settings-account-heading" className="space-y-2.5">
                <h3
                  id="settings-account-heading"
                  className="text-xs font-semibold uppercase tracking-wider text-slate-500"
                >
                  Tài khoản đang sử dụng
                </h3>

                {user ? (
                  <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3.5">
                    <div className="flex items-center gap-3.5 min-w-0">
                      <div className="relative shrink-0">
                        <div className="h-11 w-11 rounded-full bg-[#4F46E5] text-white text-sm font-bold flex items-center justify-center shadow-sm shadow-indigo-600/20">
                          {getUserInitial(user)}
                        </div>
                        <span
                          title="Đang trực tuyến"
                          className="absolute -bottom-0.5 -right-0.5 h-3.5 w-3.5 rounded-full bg-emerald-500 ring-2 ring-[#FAF8F5]"
                        />
                      </div>

                      <div className="min-w-0">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-sm font-semibold text-slate-900 truncate">
                            {user.displayName || user.email}
                          </span>
                          <span className="inline-flex items-center gap-1 rounded-full bg-emerald-50 border border-emerald-200 px-2 py-0.5 text-[10.5px] font-semibold text-emerald-700">
                            <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
                            <span>Online</span>
                          </span>
                        </div>
                        <p className="text-xs text-slate-500 truncate mt-0.5">
                          {user.email}
                        </p>
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={onRequestLogout}
                      className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[38px] px-3.5 rounded-xl border border-rose-200 bg-white hover:bg-rose-50 text-xs font-semibold text-rose-600 transition-colors"
                    >
                      <LogOut className="h-3.5 w-3.5" />
                      <span>Đăng xuất</span>
                    </button>
                  </div>
                ) : (
                  <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3.5">
                    <div className="flex items-start gap-3 min-w-0">
                      <span className="h-10 w-10 rounded-full bg-[#EFECE6] border border-[#E2DCD2] text-slate-600 flex items-center justify-center shrink-0">
                        <User className="h-5 w-5" />
                      </span>
                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <p className="text-sm font-semibold text-slate-900">
                            Chế độ Khách (Local Mode)
                          </p>
                          <span className="inline-flex items-center rounded-full bg-slate-200/75 px-2 py-0.5 text-[10.5px] font-medium text-slate-600">
                            Chỉ lưu trên máy này
                          </span>
                        </div>
                        <p className="text-xs text-slate-500 leading-relaxed mt-0.5">
                          Đăng nhập tài khoản Supabase để đồng bộ tệp và chỉnh
                          sửa giữa nhiều thiết bị.
                        </p>
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={onOpenAuth}
                      className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[40px] px-4 rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm shadow-indigo-600/20 transition-colors whitespace-nowrap"
                    >
                      <LogIn className="h-3.5 w-3.5" />
                      <span>Đăng nhập / Đồng bộ</span>
                    </button>
                  </div>
                )}
              </section>

              {/* 2. Auto-sync Option */}
              <section aria-labelledby="settings-autosync-heading" className="space-y-2.5">
                <h3
                  id="settings-autosync-heading"
                  className="text-xs font-semibold uppercase tracking-wider text-slate-500"
                >
                  Tùy chọn đồng bộ
                </h3>

                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] px-4 py-3.5 flex items-center justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-xs sm:text-sm font-semibold text-slate-900">
                      Tự động đồng bộ khi sửa (Auto-sync)
                    </p>
                    <p className="text-xs text-slate-500 leading-relaxed mt-0.5">
                      Tự động lưu thay đổi trong Word, Excel &amp; Slide sau khi
                      ngừng gõ và đẩy lên tài khoản nếu đã đăng nhập.
                    </p>
                  </div>

                  <button
                    type="button"
                    role="switch"
                    aria-checked={settings.autoSyncOnEdit}
                    aria-label="Tự động đồng bộ khi sửa"
                    onClick={() =>
                      onUpdateSettings({
                        autoSyncOnEdit: !settings.autoSyncOnEdit,
                      })
                    }
                    className={`relative inline-flex h-6 w-11 shrink-0 items-center rounded-full transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40 ${
                      settings.autoSyncOnEdit ? 'bg-[#4F46E5]' : 'bg-slate-300'
                    }`}
                  >
                    <span
                      className={`inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform ${
                        settings.autoSyncOnEdit
                          ? 'translate-x-6'
                          : 'translate-x-1'
                      }`}
                    />
                  </button>
                </div>
              </section>

              {/* 3. Default Word Paper Theme */}
              <section aria-labelledby="settings-theme-heading" className="space-y-2.5">
                <h3
                  id="settings-theme-heading"
                  className="text-xs font-semibold uppercase tracking-wider text-slate-500"
                >
                  Giao diện mặc định khi đọc Word
                </h3>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                  {WORD_THEME_OPTIONS.map((option) => {
                    const Icon = option.icon;
                    const isSelected =
                      settings.defaultWordTheme === option.key;
                    return (
                      <button
                        key={option.key}
                        type="button"
                        aria-pressed={isSelected}
                        onClick={() =>
                          onUpdateSettings({ defaultWordTheme: option.key })
                        }
                        className={`text-left rounded-2xl border p-3 transition-all flex flex-col gap-2.5 ${
                          isSelected
                            ? 'border-[#4F46E5] bg-indigo-50/50 ring-2 ring-indigo-500/15'
                            : 'border-[#E6E0D6] bg-[#FAF8F5]/70 hover:bg-white hover:border-slate-300'
                        }`}
                      >
                        {/* Mini paper preview */}
                        <div
                          className={`h-14 w-full rounded-xl border ${option.previewBg} ${option.previewBorder} p-2.5 flex flex-col justify-between relative overflow-hidden`}
                        >
                          <div className="space-y-1.5">
                            <div
                              className={`h-1.5 w-3/4 rounded-full ${option.linePrimary}`}
                            />
                            <div
                              className={`h-1.5 w-full rounded-full ${option.lineSecondary}`}
                            />
                            <div
                              className={`h-1.5 w-2/3 rounded-full ${option.lineSecondary}`}
                            />
                          </div>
                          {isSelected && (
                            <span className="absolute top-1.5 right-1.5 h-5 w-5 rounded-full bg-[#4F46E5] text-white flex items-center justify-center shadow-sm">
                              <Check className="h-3 w-3" />
                            </span>
                          )}
                        </div>

                        <div className="flex items-center gap-2 min-w-0">
                          <Icon
                            className={`h-3.5 w-3.5 shrink-0 ${
                              isSelected ? 'text-[#4F46E5]' : 'text-slate-500'
                            }`}
                          />
                          <div className="min-w-0">
                            <p className="text-xs font-semibold text-slate-900 truncate">
                              {option.label}
                            </p>
                            <p className="text-[11px] text-slate-500 truncate">
                              {option.description}
                            </p>
                          </div>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </section>

              {/* 3.5 AI Tools / Gemini API Key */}
              <section aria-labelledby="settings-ai-heading" className="space-y-2.5">
                <div className="flex items-center justify-between gap-2">
                  <h3
                    id="settings-ai-heading"
                    className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-slate-500"
                  >
                    <Sparkles className="h-3.5 w-3.5 text-indigo-600" />
                    <span>Trợ lý AI Tools (Gemini API)</span>
                  </h3>
                  <span
                    className={`rounded-full px-2 py-0.5 text-[10.5px] font-semibold ${
                      settings.geminiApiKey?.trim()
                        ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                        : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {settings.geminiApiKey?.trim()
                      ? 'Đã gắn khóa Gemini'
                      : 'Chế độ cục bộ (Offline)'}
                  </span>
                </div>

                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 space-y-2.5">
                  <label
                    htmlFor="settings-gemini-key"
                    className="block text-xs font-semibold text-slate-800"
                  >
                    Khóa Gemini API cá nhân (tùy chọn)
                  </label>
                  <div className="flex items-center gap-2">
                    <div className="relative flex-1 min-w-0">
                      <KeyRound className="pointer-events-none absolute left-3 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-slate-400" />
                      <input
                        id="settings-gemini-key"
                        type={showGeminiKey ? 'text' : 'password'}
                        value={settings.geminiApiKey ?? ''}
                        onChange={(event) =>
                          onUpdateSettings({
                            geminiApiKey: event.target.value,
                          })
                        }
                        placeholder="AIzaSy... (để trống nếu dùng bộ xử lý cục bộ)"
                        autoComplete="off"
                        spellCheck={false}
                        className="w-full rounded-xl border border-[#DDD6CA] bg-white py-2 pl-8 pr-9 text-xs text-slate-900 placeholder:text-slate-400 focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                      />
                      <button
                        type="button"
                        onClick={() => setShowGeminiKey((prev) => !prev)}
                        aria-label={showGeminiKey ? 'Ẩn khóa API' : 'Hiện khóa API'}
                        className="absolute right-2 top-1/2 -translate-y-1/2 flex h-7 w-7 items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                      >
                        {showGeminiKey ? (
                          <EyeOff className="h-3.5 w-3.5" />
                        ) : (
                          <Eye className="h-3.5 w-3.5" />
                        )}
                      </button>
                    </div>
                    {settings.geminiApiKey?.trim() && (
                      <button
                        type="button"
                        onClick={() => onUpdateSettings({ geminiApiKey: '' })}
                        className="shrink-0 rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-600 hover:bg-rose-50 hover:text-rose-600"
                      >
                        Xóa
                      </button>
                    )}
                  </div>
                  <p className="text-[11px] leading-relaxed text-slate-500">
                    Khóa chỉ lưu trên trình duyệt/thiết bị này và được dùng trực tiếp khi chạy{' '}
                    <strong>AI Tools</strong> trong Word, Excel và PowerPoint. Khi ngoại tuyến hoặc không nhập khóa, LITE tự động chuyển sang bộ xử lý thông minh cục bộ.
                  </p>
                </div>
              </section>

              {/* 4. Cloud Stats & Manual Sync */}
              <section aria-labelledby="settings-sync-stats-heading" className="space-y-3">
                <h3
                  id="settings-sync-stats-heading"
                  className="text-xs font-semibold uppercase tracking-wider text-slate-500"
                >
                  Thống kê &amp; Đồng bộ thủ công
                </h3>

                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    {/* Cloud count */}
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3">
                      <div className="flex items-center gap-1.5 text-[11px] text-slate-500 font-medium">
                        {user ? (
                          <Cloud className="h-3.5 w-3.5 text-[#4F46E5]" />
                        ) : (
                          <CloudOff className="h-3.5 w-3.5 text-slate-400" />
                        )}
                        <span>Đã lưu đám mây</span>
                      </div>
                      <p className="mt-1 text-base font-bold text-slate-900 tabular-nums">
                        {cloudDocumentCount}{' '}
                        <span className="text-xs font-medium text-slate-500">
                          tệp
                        </span>
                      </p>
                    </div>

                    {/* Connection & sync status */}
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3">
                      <div className="flex items-center gap-1.5 text-[11px] text-slate-500 font-medium">
                        {isSupabaseConfigured && user ? (
                          <Wifi className="h-3.5 w-3.5 text-emerald-600" />
                        ) : (
                          <WifiOff className="h-3.5 w-3.5 text-slate-400" />
                        )}
                        <span>Trạng thái kết nối</span>
                      </div>
                      <div className="mt-1.5">
                        <span
                          className={`inline-flex items-center gap-1.5 rounded-full border px-2 py-0.5 text-[11px] font-semibold ${statusMeta.badgeClass}`}
                        >
                          {isSyncing ? (
                            <Loader2 className="h-3 w-3 animate-spin" />
                          ) : syncStatus === 'synced' ? (
                            <CheckCircle2 className="h-3 w-3" />
                          ) : (
                            <span
                              className={`h-1.5 w-1.5 rounded-full ${statusMeta.dotClass}`}
                            />
                          )}
                          <span>{statusMeta.label}</span>
                        </span>
                      </div>
                    </div>

                    {/* Last synced */}
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3">
                      <div className="flex items-center gap-1.5 text-[11px] text-slate-500 font-medium">
                        <RefreshCw className="h-3.5 w-3.5 text-slate-400" />
                        <span>Đồng bộ gần nhất</span>
                      </div>
                      <p className="mt-1 text-xs font-semibold text-slate-800 leading-snug">
                        {lastSyncText}
                      </p>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handleManualSyncClick}
                    disabled={!canManualSync}
                    className="w-full min-h-[44px] rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold shadow-sm shadow-indigo-600/20 flex items-center justify-center gap-2 transition-colors disabled:bg-slate-200 disabled:text-slate-500 disabled:shadow-none disabled:cursor-not-allowed"
                  >
                    {isSyncing ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin" />
                        <span>Đang đồng bộ dữ liệu...</span>
                      </>
                    ) : (
                      <>
                        <RefreshCw className="h-4 w-4" />
                        <span>Đồng bộ thủ công ngay</span>
                      </>
                    )}
                  </button>

                  {(!isSupabaseConfigured || !user) && (
                    <div className="rounded-xl border border-amber-200/80 bg-amber-50/70 px-3.5 py-2.5 flex items-start gap-2 text-xs text-amber-900">
                      <AlertCircle className="h-4 w-4 text-amber-600 shrink-0 mt-0.5" />
                      <span className="leading-relaxed">
                        {!isSupabaseConfigured
                          ? 'Chức năng đồng bộ đám mây đang khóa vì dự án chưa cấu hình biến môi trường Supabase (.env).'
                          : 'Vui lòng đăng nhập tài khoản để kích hoạt đồng bộ thủ công và đồng bộ đa thiết bị.'}
                      </span>
                    </div>
                  )}
                </div>
              </section>

              {/* Quản lý Email được phép dùng App (Ép mail từ xa không cần build lại) */}
              {user && isSupabaseConfigured && (
                <section className="rounded-2xl border border-indigo-200/80 bg-indigo-50/40 p-4 space-y-3">
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <ShieldCheck className="h-4 w-4 text-indigo-600 shrink-0" />
                      <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-950">
                        Ép Email &amp; Duyệt quyền dùng App từ xa
                      </h3>
                    </div>
                    <button
                      type="button"
                      onClick={() => void refreshAllowedEmails()}
                      disabled={isLoadingAllowed}
                      className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-white border border-indigo-200 text-[11px] font-medium text-indigo-700 hover:bg-indigo-50"
                    >
                      <RefreshCw className={`h-3 w-3 ${isLoadingAllowed ? "animate-spin" : ""}`} />
                      <span>Làm mới</span>
                    </button>
                  </div>

                  <p className="text-[11.5px] text-slate-600 leading-relaxed">
                    Nhập Gmail của người muốn dùng app bên dưới rồi bấm <strong>Ép Email</strong>. Họ có thể đăng nhập bằng Google ngay lập tức mà <strong>không cần đóng gói lại app</strong>.
                  </p>

                  {allowedTableMissing ? (
                    <div className="rounded-xl bg-amber-50 border border-amber-200 p-3 space-y-2 text-xs text-amber-900">
                      <p className="font-semibold">
                        Cần chạy 1 đoạn SQL ngắn trong Supabase SQL Editor để kích hoạt bảng danh sách email:
                      </p>
                      <button
                        type="button"
                        onClick={() => {
                          void navigator.clipboard.writeText(
                            `create table if not exists public.app_allowed_emails (email text primary key, status text not null default 'approved', note text, added_by text, created_at timestamptz default now()); alter table public.app_allowed_emails enable row level security; create policy "read_all" on public.app_allowed_emails for select using (true); create policy "insert_req" on public.app_allowed_emails for insert with check (true); create policy "manage_auth" on public.app_allowed_emails for all using (auth.role() = 'authenticated');`
                          );
                          setCopiedSql(true);
                          window.setTimeout(() => setCopiedSql(false), 2500);
                        }}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-amber-600 text-white font-semibold"
                      >
                        <Copy className="h-3.5 w-3.5" />
                        <span>{copiedSql ? "Đã sao chép lệnh SQL!" : "Sao chép lệnh SQL tạo bảng"}</span>
                      </button>
                    </div>
                  ) : (
                    <>
                      <div className="flex flex-col sm:flex-row gap-2">
                        <input
                          type="email"
                          value={newAllowedEmail}
                          onChange={(e) => setNewAllowedEmail(e.target.value)}
                          placeholder="nguoi-dung@gmail.com"
                          className="flex-1 h-9 rounded-xl border border-indigo-200 bg-white px-3 text-xs text-slate-900 outline-none focus:border-indigo-500"
                        />
                        <input
                          type="text"
                          value={newAllowedNote}
                          onChange={(e) => setNewAllowedNote(e.target.value)}
                          placeholder="Ghi chú (tên)..."
                          className="sm:w-32 h-9 rounded-xl border border-indigo-200 bg-white px-3 text-xs text-slate-900 outline-none focus:border-indigo-500"
                        />
                        <button
                          type="button"
                          disabled={isSavingAllowed || !newAllowedEmail.trim()}
                          onClick={() => void handleAddAllowedEmail()}
                          className="inline-flex items-center justify-center gap-1.5 h-9 px-3.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shrink-0 disabled:opacity-50"
                        >
                          <UserPlus className="h-3.5 w-3.5" />
                          <span>Ép Email</span>
                        </button>
                      </div>

                      {allowedFeedback && (
                        <p className="text-[11.5px] font-medium text-indigo-800 bg-white border border-indigo-200 rounded-xl px-3 py-2">
                          {allowedFeedback}
                        </p>
                      )}

                      {allowedList.length > 0 && (
                        <div className="max-h-44 overflow-y-auto space-y-1.5 pr-1">
                          {allowedList.map((item) => (
                            <div
                              key={item.email}
                              className="flex items-center justify-between gap-2 px-3 py-2 rounded-xl bg-white border border-slate-200/90 text-xs"
                            >
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-1.5">
                                  <span className="font-semibold text-slate-900 truncate">
                                    {item.email}
                                  </span>
                                  <span
                                    className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                      item.status === "approved"
                                        ? "bg-emerald-50 text-emerald-700 border border-emerald-200"
                                        : "bg-amber-50 text-amber-700 border border-amber-200"
                                    }`}
                                  >
                                    {item.status === "approved" ? "Đã cấp quyền" : "Chờ duyệt"}
                                  </span>
                                </div>
                                {item.note && (
                                  <p className="text-[11px] text-slate-500 truncate">{item.note}</p>
                                )}
                              </div>
                              <div className="flex items-center gap-1 shrink-0">
                                {item.status !== "approved" && (
                                  <button
                                    type="button"
                                    disabled={isSavingAllowed}
                                    onClick={() => void handleApprovePending(item.email, item.note)}
                                    title="Duyệt cho phép dùng app"
                                    className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold"
                                  >
                                    <CheckCheck className="h-3 w-3" />
                                    <span>Duyệt</span>
                                  </button>
                                )}
                                <button
                                  type="button"
                                  disabled={isSavingAllowed}
                                  onClick={() => void handleDeleteAllowed(item.email)}
                                  title="Xóa / Thu hồi quyền"
                                  className="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </>
                  )}
                </section>
              )}

            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
