import { useEffect, useRef, useState } from 'react';
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react';
import {
  ArrowLeft,
  BookOpen,
  Check,
  ChevronDown,
  CloudUpload,
  Download,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  HardDrive,
  Loader2,
  Pencil,
  Presentation,
  Share2,
  X,
} from 'lucide-react';
import { CollaboratorPresenceBar } from "./CollaboratorPresenceBar";
import type { DocumentItem, DocumentType } from '../../types/document';
import { isSupabaseConfigured, supabase } from '../../lib/supabase';
import { buildDocumentUploadPayload } from '../../utils/fileExporter';
import {
  GoogleDriveAuthError,
  getStoredGoogleDriveToken,
  uploadFileToGoogleDrive,
} from '../../services/googleDrive';

export interface WorkspaceTopBarProps {
  collaborators?: any[];
  selfSessionId?: string;
  collabPermission?: any;
  document: DocumentItem;
  theme?: 'light' | 'dark';
  onBackHome: () => void;
  onShare: () => void;
  onRenameTitle?: (newTitle: string) => void;
  leftExtra?: ReactNode;
  toolbar?: ReactNode;
}

interface FormatVisualConfig {
  label: string;
  Icon: typeof Presentation;
  iconBoxLight: string;
  iconBoxDark: string;
  badgeLight: string;
  badgeDark: string;
}

type DriveUploadStatus =
  | 'idle'
  | 'uploading'
  | 'success'
  | 'auth_required'
  | 'error';

const FORMAT_VISUALS: Record<DocumentType, FormatVisualConfig> = {
  pptx: {
    label: 'PPTX',
    Icon: Presentation,
    iconBoxLight: 'bg-orange-50 text-orange-600 border-orange-200/80',
    iconBoxDark: 'bg-orange-500/15 text-orange-300 border-orange-400/25',
    badgeLight: 'bg-orange-50 text-orange-700 border-orange-200/80',
    badgeDark: 'bg-orange-500/15 text-orange-300 border-orange-400/25',
  },
  docx: {
    label: 'DOCX',
    Icon: FileText,
    iconBoxLight: 'bg-blue-50 text-blue-600 border-blue-200/80',
    iconBoxDark: 'bg-blue-500/15 text-blue-300 border-blue-400/25',
    badgeLight: 'bg-blue-50 text-blue-700 border-blue-200/80',
    badgeDark: 'bg-blue-500/15 text-blue-300 border-blue-400/25',
  },
  xlsx: {
    label: 'XLSX',
    Icon: FileSpreadsheet,
    iconBoxLight: 'bg-green-50 text-green-700 border-green-200/80',
    iconBoxDark: 'bg-emerald-500/15 text-emerald-300 border-emerald-400/25',
    badgeLight: 'bg-green-50 text-green-800 border-green-200/80',
    badgeDark: 'bg-emerald-500/15 text-emerald-300 border-emerald-400/25',
  },
  pdf: {
    label: 'PDF',
    Icon: BookOpen,
    iconBoxLight: 'bg-pink-50 text-pink-600 border-pink-200/80',
    iconBoxDark: 'bg-pink-500/15 text-pink-300 border-pink-400/25',
    badgeLight: 'bg-pink-50 text-pink-700 border-pink-200/80',
    badgeDark: 'bg-pink-500/15 text-pink-300 border-pink-400/25',
  },
};

export function WorkspaceTopBar({
  collaborators,
  selfSessionId,
  collabPermission,
  document,
  theme = 'light',
  onBackHome,
  onShare,
  onRenameTitle,
  leftExtra,
  toolbar,
}: WorkspaceTopBarProps) {
  const isDark = theme === 'dark';
  const visual = FORMAT_VISUALS[document.type];
  const FormatIcon = visual.Icon;

  const [isEditingTitle, setIsEditingTitle] = useState<boolean>(false);
  const [titleDraft, setTitleDraft] = useState<string>(document.title);
  const titleInputRef = useRef<HTMLInputElement | null>(null);
  const cancelBlurRef = useRef<boolean>(false);

  // Storage / Google Drive dropdown state
  const [isStorageMenuOpen, setIsStorageMenuOpen] = useState<boolean>(false);
  const [driveStatus, setDriveStatus] = useState<DriveUploadStatus>('idle');
  const [driveMessage, setDriveMessage] = useState<string | null>(null);
  const [driveViewLink, setDriveViewLink] = useState<string | null>(null);
  const [isConnectingGoogle, setIsConnectingGoogle] = useState<boolean>(false);
  const [isDownloadingLocal, setIsDownloadingLocal] = useState<boolean>(false);
  const storageMenuRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    setTitleDraft(document.title);
    setIsEditingTitle(false);
    setIsStorageMenuOpen(false);
    setDriveStatus('idle');
    setDriveMessage(null);
    setDriveViewLink(null);
  }, [document.id, document.title]);

  useEffect(() => {
    if (isEditingTitle && titleInputRef.current) {
      titleInputRef.current.focus();
      titleInputRef.current.select();
    }
  }, [isEditingTitle]);

  // Close Storage dropdown on outside click or Escape
  useEffect(() => {
    if (!isStorageMenuOpen) {
      return;
    }
    const handlePointerDown = (event: MouseEvent) => {
      if (
        storageMenuRef.current &&
        !storageMenuRef.current.contains(event.target as Node)
      ) {
        setIsStorageMenuOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setIsStorageMenuOpen(false);
      }
    };
    window.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isStorageMenuOpen]);

  const startEditingTitle = () => {
    if (!onRenameTitle) {
      return;
    }
    cancelBlurRef.current = false;
    setTitleDraft(document.title);
    setIsEditingTitle(true);
  };

  const commitTitleRename = () => {
    if (cancelBlurRef.current) {
      cancelBlurRef.current = false;
      return;
    }
    const trimmed = titleDraft.trim();
    setIsEditingTitle(false);
    if (!trimmed) {
      setTitleDraft(document.title);
      return;
    }
    if (trimmed !== document.title) {
      onRenameTitle?.(trimmed);
    }
  };

  const cancelTitleRename = () => {
    cancelBlurRef.current = true;
    setTitleDraft(document.title);
    setIsEditingTitle(false);
  };

  const handleTitleInputKeyDown = (
    event: ReactKeyboardEvent<HTMLInputElement>
  ) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      commitTitleRename();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      cancelTitleRename();
    }
  };

  const handleConnectGoogleDrive = async () => {
    if (!isSupabaseConfigured || !supabase) {
      setDriveStatus('error');
      setDriveMessage(
        'Chế độ Local Mode chưa cấu hình Supabase (.env). Vui lòng thiết lập VITE_SUPABASE_URL & VITE_SUPABASE_ANON_KEY để kết nối Google Drive.'
      );
      return;
    }
    setIsConnectingGoogle(true);
    setDriveMessage(null);
    try {
      const { error } = await supabase.auth.signInWithOAuth({
        provider: 'google',
        options: {
          redirectTo: window.location.origin,
          scopes:
            'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly',
          queryParams: { access_type: 'offline', prompt: 'consent' },
        },
      });
      if (error) {
        setDriveStatus('error');
        setDriveMessage(
          error.message || 'Không thể khởi tạo đăng nhập Google Drive.'
        );
      }
    } catch (err: unknown) {
      setDriveStatus('error');
      setDriveMessage(
        err instanceof Error
          ? err.message
          : 'Đã xảy ra lỗi khi kết nối Google Drive.'
      );
    } finally {
      setIsConnectingGoogle(false);
    }
  };

  const handleSaveToGoogleDrive = async () => {
    const token = getStoredGoogleDriveToken();
    if (!token) {
      setDriveStatus('auth_required');
      setDriveMessage(
        'Bạn chưa kết nối tài khoản Google Drive. Kết nối ngay để tải tệp lên Drive.'
      );
      setDriveViewLink(null);
      return;
    }

    setDriveStatus('uploading');
    setDriveMessage(null);
    setDriveViewLink(null);

    try {
      const { blob, fileName, mimeType } =
        await buildDocumentUploadPayload(document);
      const uploaded = await uploadFileToGoogleDrive(blob, fileName, mimeType);
      setDriveStatus('success');
      setDriveMessage(`Đã lưu lên Drive: ${uploaded.name || fileName}`);
      setDriveViewLink(uploaded.webViewLink ?? null);
    } catch (err: unknown) {
      if (err instanceof GoogleDriveAuthError) {
        setDriveStatus('auth_required');
        setDriveMessage(
          err.message ||
            'Phiên truy cập Google Drive đã hết hạn. Vui lòng kết nối lại Google Drive.'
        );
      } else {
        setDriveStatus('error');
        setDriveMessage(
          err instanceof Error
            ? err.message
            : 'Không thể lưu tệp lên Google Drive. Vui lòng thử lại.'
        );
      }
    }
  };

  const handleDownloadToDevice = async () => {
    setIsDownloadingLocal(true);
    setDriveMessage(null);
    try {
      const { blob, fileName } = await buildDocumentUploadPayload(document);
      const url = URL.createObjectURL(blob);
      const anchor = window.document.createElement('a');
      anchor.href = url;
      anchor.download = fileName;
      anchor.style.display = 'none';
      window.document.body.appendChild(anchor);
      anchor.click();
      window.document.body.removeChild(anchor);
      window.setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 1500);
      setIsStorageMenuOpen(false);
    } catch (err: unknown) {
      setDriveStatus('error');
      setDriveMessage(
        err instanceof Error
          ? err.message
          : 'Không thể tải tệp xuống thiết bị.'
      );
    } finally {
      setIsDownloadingLocal(false);
    }
  };

  return (
    <header
      className={`h-14 shrink-0 w-full px-2.5 sm:px-4 flex items-center justify-between gap-2 sm:gap-3 border-b transition-colors duration-200 z-20 relative ${
        isDark
          ? 'bg-[#141923] border-white/10 text-slate-100'
          : 'bg-[#FFFFFF] border-[#EAE6DF] text-slate-900'
      }`}
    >
      {/* Left Cluster: Back Home + Optional Drawer Trigger + Format Icon + File Title + Badge */}
      <div className="flex-1 min-w-0 flex items-center gap-1.5 sm:gap-2">
        <button
          type="button"
          onClick={onBackHome}
          aria-label="Quay lại Trang chủ"
          title="Quay lại Trang chủ"
          className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[38px] min-w-[38px] sm:min-w-0 px-2 sm:px-2.5 rounded-xl text-xs font-medium border transition-colors ${
            isDark
              ? 'bg-white/[0.05] hover:bg-white/[0.12] border-white/10 text-slate-200 hover:text-white'
              : 'bg-[#FAF8F5] hover:bg-[#F2EFEA] border-[#E7E2DA] text-slate-700 hover:text-slate-950'
          }`}
        >
          <ArrowLeft className="h-4 w-4 shrink-0" />
          <span className="hidden sm:inline whitespace-nowrap">Trang chủ</span>
        </button>

        {leftExtra && (
          <div className="shrink-0 flex items-center">{leftExtra}</div>
        )}

        <span
          aria-hidden="true"
          className={`hidden sm:block h-5 w-px shrink-0 ${
            isDark ? 'bg-white/10' : 'bg-[#EAE6DF]'
          }`}
        />

        <div className="min-w-0 flex items-center gap-1.5 sm:gap-2">
          <span
            aria-hidden="true"
            className={`hidden sm:flex shrink-0 h-8 w-8 rounded-lg border items-center justify-center ${
              isDark ? visual.iconBoxDark : visual.iconBoxLight
            }`}
          >
            <FormatIcon className="h-4 w-4" />
          </span>

          {onRenameTitle && isEditingTitle ? (
            <input
              ref={titleInputRef}
              type="text"
              value={titleDraft}
              onChange={(event) => setTitleDraft(event.target.value)}
              onKeyDown={handleTitleInputKeyDown}
              onBlur={commitTitleRename}
              maxLength={120}
              aria-label="Đổi tên tệp"
              placeholder="Nhập tên tệp..."
              className={`min-w-[110px] w-[140px] sm:w-[210px] md:w-[260px] max-w-full h-8 px-2.5 rounded-lg border text-[13.5px] sm:text-sm font-semibold tracking-tight outline-none transition-all ${
                isDark
                  ? 'bg-[#1C2333] border-indigo-400 text-white placeholder:text-slate-500 focus:ring-2 focus:ring-indigo-400/25'
                  : 'bg-white border-indigo-500 text-slate-900 placeholder:text-slate-400 focus:ring-2 focus:ring-indigo-500/20'
              }`}
            />
          ) : onRenameTitle ? (
            <button
              type="button"
              onClick={startEditingTitle}
              title="Đổi tên tệp"
              aria-label={`Đổi tên tệp: ${document.title}`}
              className={`group min-w-0 inline-flex items-center gap-1.5 px-2 py-1 -mx-1 rounded-lg transition-colors ${
                isDark
                  ? 'hover:bg-white/[0.08] text-white'
                  : 'hover:bg-[#F3F1EC] text-slate-900'
              }`}
            >
              <h1 className="truncate text-[13.5px] sm:text-sm font-semibold tracking-tight">
                {document.title}
              </h1>
              <Pencil
                className={`h-3.5 w-3.5 shrink-0 transition-opacity ${
                  isDark
                    ? 'text-slate-400 opacity-75 group-hover:opacity-100 group-hover:text-indigo-300'
                    : 'text-slate-400 opacity-75 group-hover:opacity-100 group-hover:text-indigo-600'
                }`}
              />
            </button>
          ) : (
            <h1
              title={document.title}
              className={`truncate text-[13.5px] sm:text-sm font-semibold tracking-tight ${
                isDark ? 'text-white' : 'text-slate-900'
              }`}
            >
              {document.title}
            </h1>
          )}

          <span
            className={`hidden sm:inline-flex shrink-0 items-center px-1.5 py-0.5 rounded-md border text-[10px] font-bold tracking-wider ${
              isDark ? visual.badgeDark : visual.badgeLight
            }`}
          >
            {visual.label}
          </span>
        </div>
      </div>

      {/* Right Cluster: Format-Specific Toolbar + Storage/Google Drive Menu + Share Button */}
      <div className="flex items-center gap-1.5 sm:gap-2 min-w-0 max-w-[64%] sm:max-w-[72%]">
        {toolbar && (
          <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar py-0.5 min-w-0">
            {toolbar}
          </div>
        )}

        {/* Menu Lưu trữ / Google Drive */}
        <div ref={storageMenuRef} className="relative shrink-0">
          <button
            type="button"
            onClick={() => setIsStorageMenuOpen((prev) => !prev)}
            aria-expanded={isStorageMenuOpen}
            aria-haspopup="menu"
            aria-label="Lưu trữ và Google Drive"
            title="Lưu trữ / Lưu lên Google Drive"
            className={`shrink-0 inline-flex items-center justify-center gap-1 min-h-[38px] sm:min-h-[36px] px-2 sm:px-2.5 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
              driveStatus === 'success'
                ? isDark
                  ? 'bg-emerald-500/20 border-emerald-400/35 text-emerald-200'
                  : 'bg-emerald-50 border-emerald-200 text-emerald-700'
                : isDark
                  ? 'bg-white/[0.06] hover:bg-white/[0.14] border-white/12 text-slate-200 hover:text-white'
                  : 'bg-[#FAF8F5] hover:bg-[#F2EFEA] border-[#E5E0D8] text-slate-700 hover:text-slate-950'
            }`}
          >
            {driveStatus === 'uploading' ? (
              <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
            ) : driveStatus === 'success' ? (
              <Check className="h-3.5 w-3.5 shrink-0" />
            ) : (
              <CloudUpload className="h-3.5 w-3.5 shrink-0" />
            )}
            <span className="hidden lg:inline">
              {driveStatus === 'success' ? 'Đã lưu Drive' : 'Lưu trữ'}
            </span>
            <ChevronDown className="h-3 w-3 opacity-75 shrink-0 hidden sm:inline" />
          </button>

          {isStorageMenuOpen && (
            <div
              role="menu"
              aria-label="Tùy chọn lưu trữ"
              className={`absolute right-0 top-full mt-2 w-72 max-w-[calc(100vw-24px)] rounded-2xl border shadow-xl p-2 z-50 ${
                isDark
                  ? 'bg-[#1A2130] border-white/15 text-slate-100'
                  : 'bg-white border-[#E5E0D8] text-slate-900'
              }`}
            >
              <div
                className={`flex items-center justify-between px-2.5 py-1.5 mb-1 border-b ${
                  isDark ? 'border-white/10' : 'border-slate-100'
                }`}
              >
                <span
                  className={`text-[11px] font-bold uppercase tracking-wider ${
                    isDark ? 'text-slate-400' : 'text-slate-500'
                  }`}
                >
                  Lưu trữ &amp; Xuất tệp
                </span>
                <button
                  type="button"
                  onClick={() => setIsStorageMenuOpen(false)}
                  aria-label="Đóng menu lưu trữ"
                  className={`p-1 rounded-lg transition-colors ${
                    isDark
                      ? 'text-slate-400 hover:text-white hover:bg-white/10'
                      : 'text-slate-400 hover:text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              </div>

              <div className="space-y-1">
                <button
                  type="button"
                  role="menuitem"
                  disabled={driveStatus === 'uploading'}
                  onClick={() => void handleSaveToGoogleDrive()}
                  className={`w-full flex items-center gap-2.5 px-2.5 py-2.5 rounded-xl text-left transition-colors disabled:opacity-60 ${
                    isDark
                      ? 'hover:bg-white/[0.08] text-slate-100'
                      : 'hover:bg-[#F7F5F0] text-slate-800'
                  }`}
                >
                  <span
                    className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 border ${
                      isDark
                        ? 'bg-emerald-500/15 border-emerald-400/25 text-emerald-300'
                        : 'bg-emerald-50 border-emerald-200/80 text-emerald-600'
                    }`}
                  >
                    {driveStatus === 'uploading' ? (
                      <Loader2 className="h-4 w-4 animate-spin" />
                    ) : (
                      <HardDrive className="h-4 w-4" />
                    )}
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-xs font-semibold truncate">
                      {driveStatus === 'uploading'
                        ? 'Đang tải lên Google Drive...'
                        : 'Lưu lên Google Drive'}
                    </p>
                    <p
                      className={`text-[11px] truncate ${
                        isDark ? 'text-slate-400' : 'text-slate-500'
                      }`}
                    >
                      Đồng bộ trực tiếp vào Drive của bạn
                    </p>
                  </div>
                </button>

                <button
                  type="button"
                  role="menuitem"
                  disabled={isDownloadingLocal}
                  onClick={() => void handleDownloadToDevice()}
                  className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left transition-colors disabled:opacity-60 ${
                    isDark
                      ? 'hover:bg-white/[0.08] text-slate-200'
                      : 'hover:bg-[#F7F5F0] text-slate-700'
                  }`}
                >
                  <span
                    className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 border ${
                      isDark
                        ? 'bg-white/[0.06] border-white/10 text-slate-300'
                        : 'bg-slate-50 border-slate-200/80 text-slate-600'
                    }`}
                  >
                    {isDownloadingLocal ? (
                      <Loader2 className="h-4 w-4 animate-spin" />
                    ) : (
                      <Download className="h-4 w-4" />
                    )}
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-xs font-semibold truncate">
                      Tải xuống máy (.{document.type})
                    </p>
                    <p
                      className={`text-[11px] truncate ${
                        isDark ? 'text-slate-400' : 'text-slate-500'
                      }`}
                    >
                      Lưu bản sao ngoại tuyến trên thiết bị
                    </p>
                  </div>
                </button>
              </div>

              {/* Phản hồi trạng thái lưu Google Drive */}
              {driveStatus === 'success' && (
                <div
                  className={`mt-2 p-2.5 rounded-xl border text-xs ${
                    isDark
                      ? 'bg-emerald-500/15 border-emerald-400/25 text-emerald-200'
                      : 'bg-emerald-50 border-emerald-200 text-emerald-800'
                  }`}
                >
                  <div className="flex items-start gap-2">
                    <Check className="h-4 w-4 shrink-0 mt-0.5" />
                    <div className="min-w-0 flex-1">
                      <p className="font-semibold">Đã lưu lên Drive</p>
                      {driveMessage && (
                        <p className="text-[11px] opacity-85 break-words mt-0.5">
                          {driveMessage}
                        </p>
                      )}
                      {driveViewLink && (
                        <a
                          href={driveViewLink}
                          target="_blank"
                          rel="noopener noreferrer"
                          className={`mt-1.5 inline-flex items-center gap-1 font-semibold underline underline-offset-2 ${
                            isDark
                              ? 'text-emerald-300 hover:text-white'
                              : 'text-emerald-700 hover:text-emerald-950'
                          }`}
                        >
                          <span>Mở trên Google Drive</span>
                          <ExternalLink className="h-3 w-3" />
                        </a>
                      )}
                    </div>
                  </div>
                </div>
              )}

              {driveStatus === 'auth_required' && (
                <div
                  className={`mt-2 p-2.5 rounded-xl border text-xs space-y-2 ${
                    isDark
                      ? 'bg-amber-500/15 border-amber-400/25 text-amber-200'
                      : 'bg-amber-50 border-amber-200 text-amber-900'
                  }`}
                >
                  <p className="leading-relaxed">
                    {driveMessage ||
                      'Cần kết nối Google Drive để lưu tệp trực tiếp.'}
                  </p>
                  <button
                    type="button"
                    disabled={isConnectingGoogle}
                    onClick={() => void handleConnectGoogleDrive()}
                    className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm transition-colors disabled:opacity-60"
                  >
                    {isConnectingGoogle ? (
                      <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
                    ) : (
                      <HardDrive className="h-3.5 w-3.5 shrink-0" />
                    )}
                    <span>Kết nối Google Drive</span>
                  </button>
                </div>
              )}

              {driveStatus === 'error' && driveMessage && (
                <div
                  className={`mt-2 p-2.5 rounded-xl border text-xs ${
                    isDark
                      ? 'bg-rose-500/15 border-rose-400/25 text-rose-200'
                      : 'bg-rose-50 border-rose-200 text-rose-700'
                  }`}
                >
                  <p className="leading-relaxed break-words">{driveMessage}</p>
                </div>
              )}
            </div>
          )}
        </div>

        <CollaboratorPresenceBar
          collaborators={collaborators}
          selfSessionId={selfSessionId}
          collabPermission={collabPermission}
          theme={theme}
        />

        <button
          type="button"
          onClick={onShare}
          aria-label="Chia sẻ tài liệu"
          title="Chia sẻ tài liệu"
          className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[38px] sm:min-h-[36px] px-2.5 sm:px-3 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
            isDark
              ? 'bg-white/[0.1] hover:bg-white/[0.18] border-white/15 text-white'
              : 'bg-indigo-50 hover:bg-indigo-100 border-indigo-200/80 text-indigo-700'
          }`}
        >
          <Share2 className="h-3.5 w-3.5 shrink-0" />
          <span className="hidden sm:inline">Chia sẻ</span>
        </button>
      </div>
    </header>
  );
}
