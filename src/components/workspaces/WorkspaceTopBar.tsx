import { getStoredGoogleDriveToken, upsertDocumentToLiteDrive, GoogleDriveAuthError } from '../../services/googleDrive';
import { useEffect, useRef, useState } from 'react';
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react';
import {
  AlertCircle,
  ArrowLeft,
  BookOpen,
  Check,
  CloudOff,
  Download,
  ExternalLink,
  Eye,
  FileDown,
  FileOutput,
  FileSpreadsheet,
  FileText,
  HardDrive,
  History,
  Info,
  Loader2,
  MoreVertical,
  Pencil,
  Presentation,
  Save,
  Share2,
  X,
} from 'lucide-react';
import { CollaboratorPresenceBar } from './CollaboratorPresenceBar';
import type { CollaboratorPresence } from '../../services/realtimeCollab';
import type { DocumentItem, DocumentType } from '../../types/document';
import { requestGoogleDriveAccess } from '../../lib/supabase';
import { useWorkspaceSync } from '../../contexts/WorkspaceSyncContext';
import {
  buildDocumentUploadPayload,
  exportOriginalDocumentBinary,
} from '../../utils/fileExporter';
import { downloadDocumentAsPdf } from '../../utils/quickConverter';
import { saveFileUniversal } from '../../utils/saveFileUniversal';

export type TopBarEditMode = 'original' | 'edit';

export interface WorkspaceTopBarModeSwitch {
  visible?: boolean;
  mode: TopBarEditMode;
  onChange: (mode: TopBarEditMode) => void;
  editDisabled?: boolean;
  originalTitle?: string;
  editTitle?: string;
}

export interface WorkspaceTopBarExtraMenuItem {
  id: string;
  label: string;
  icon?: ReactNode;
  onClick: () => void;
  disabled?: boolean;
}

export interface WorkspaceTopBarProps {
  collaborators?: CollaboratorPresence[];
  selfSessionId?: string;
  collabPermission?: 'owner' | 'edit' | 'view';
  document: DocumentItem;
  theme?: 'light' | 'dark';
  onBackHome: () => void;
  onShare: () => void;
  onRenameTitle?: (newTitle: string) => void;
  leftExtra?: ReactNode;
  toolbar?: ReactNode;
  modeSwitch?: WorkspaceTopBarModeSwitch;
  onSave?: () => Promise<void> | void;
  isSaving?: boolean;
  isSaved?: boolean;
  isDirty?: boolean;
  onDownloadEdited?: () => Promise<void> | void;
  onExportPdf?: () => Promise<void> | void;
  extraMenuItems?: WorkspaceTopBarExtraMenuItem[];
}

interface FormatVisualConfig {
  label: string;
  Icon: typeof Presentation;
  iconBoxLight: string;
  iconBoxDark: string;
  badgeLight: string;
  badgeDark: string;
  modeActiveLight: string;
  modeActiveDark: string;
  saveDirtyClass: string;
  shareButtonLight: string;
  focusInputLight: string;
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
    iconBoxLight: 'bg-[#FFF4ED] text-[#C43E1C] border-[#C43E1C]/25',
    iconBoxDark: 'bg-[#C43E1C]/20 text-orange-300 border-[#C43E1C]/40',
    badgeLight: 'bg-[#FFF4ED] text-[#C43E1C] border-[#C43E1C]/25',
    badgeDark: 'bg-[#C43E1C]/20 text-orange-300 border-[#C43E1C]/40',
    modeActiveLight: 'bg-white text-[#C43E1C] shadow-sm ring-1 ring-[#C43E1C]/20',
    modeActiveDark: 'bg-[#C43E1C]/30 text-orange-100 shadow-sm ring-1 ring-[#C43E1C]/55',
    saveDirtyClass: 'bg-[#C43E1C] hover:bg-[#9E3014] border-[#C43E1C] text-white shadow-sm',
    shareButtonLight: 'bg-[#FFF4ED] hover:bg-[#FFE6D5] border-[#C43E1C]/25 text-[#C43E1C]',
    focusInputLight: 'bg-white border-[#C43E1C] text-slate-900 placeholder:text-slate-400 focus:ring-2 focus:ring-[#C43E1C]/20',
  },
  docx: {
    label: 'DOCX',
    Icon: FileText,
    iconBoxLight: 'bg-[#EEF4FF] text-[#185ABD] border-[#185ABD]/25',
    iconBoxDark: 'bg-[#185ABD]/20 text-blue-300 border-[#185ABD]/40',
    badgeLight: 'bg-[#EEF4FF] text-[#185ABD] border-[#185ABD]/25',
    badgeDark: 'bg-[#185ABD]/20 text-blue-300 border-[#185ABD]/40',
    modeActiveLight: 'bg-white text-[#185ABD] shadow-sm ring-1 ring-[#185ABD]/20',
    modeActiveDark: 'bg-[#185ABD]/30 text-blue-100 shadow-sm ring-1 ring-[#185ABD]/55',
    saveDirtyClass: 'bg-[#185ABD] hover:bg-[#10489A] border-[#185ABD] text-white shadow-sm',
    shareButtonLight: 'bg-[#EEF4FF] hover:bg-[#DFEAFE] border-[#185ABD]/25 text-[#185ABD]',
    focusInputLight: 'bg-white border-[#185ABD] text-slate-900 placeholder:text-slate-400 focus:ring-2 focus:ring-[#185ABD]/20',
  },
  xlsx: {
    label: 'XLSX',
    Icon: FileSpreadsheet,
    iconBoxLight: 'bg-[#ECFDF5] text-[#107C41] border-[#107C41]/25',
    iconBoxDark: 'bg-[#107C41]/20 text-emerald-300 border-[#107C41]/40',
    badgeLight: 'bg-[#ECFDF5] text-[#107C41] border-[#107C41]/25',
    badgeDark: 'bg-[#107C41]/20 text-emerald-300 border-[#107C41]/40',
    modeActiveLight: 'bg-white text-[#107C41] shadow-sm ring-1 ring-[#107C41]/20',
    modeActiveDark: 'bg-[#107C41]/30 text-emerald-100 shadow-sm ring-1 ring-[#107C41]/55',
    saveDirtyClass: 'bg-[#107C41] hover:bg-[#0B5E31] border-[#107C41] text-white shadow-sm',
    shareButtonLight: 'bg-[#ECFDF5] hover:bg-[#D1FAE5] border-[#107C41]/25 text-[#107C41]',
    focusInputLight: 'bg-white border-[#107C41] text-slate-900 placeholder:text-slate-400 focus:ring-2 focus:ring-[#107C41]/20',
  },
  pdf: {
    label: 'PDF',
    Icon: BookOpen,
    iconBoxLight: 'bg-[#FDF2F8] text-[#C81E68] border-[#C81E68]/25',
    iconBoxDark: 'bg-[#C81E68]/20 text-pink-300 border-[#C81E68]/40',
    badgeLight: 'bg-[#FDF2F8] text-[#C81E68] border-[#C81E68]/25',
    badgeDark: 'bg-[#C81E68]/20 text-pink-300 border-[#C81E68]/40',
    modeActiveLight: 'bg-white text-[#C81E68] shadow-sm ring-1 ring-[#C81E68]/20',
    modeActiveDark: 'bg-[#C81E68]/30 text-pink-100 shadow-sm ring-1 ring-[#C81E68]/55',
    saveDirtyClass: 'bg-[#C81E68] hover:bg-[#9D174D] border-[#C81E68] text-white shadow-sm',
    shareButtonLight: 'bg-[#FDF2F8] hover:bg-[#FCE7F3] border-[#C81E68]/25 text-[#C81E68]',
    focusInputLight: 'bg-white border-[#C81E68] text-slate-900 placeholder:text-slate-400 focus:ring-2 focus:ring-[#C81E68]/20',
  },
};

function resolveDocumentCountSummary(doc: DocumentItem): {
  label: string;
  value: string;
} {
  if (doc.type === 'pptx') {
    return {
      label: 'Số slide',
      value: `${doc.slidesPreview.length} slide`,
    };
  }
  if (doc.type === 'xlsx') {
    return {
      label: 'Số sheet',
      value: `${doc.sheets?.length ?? 0} trang tính`,
    };
  }
  if (doc.type === 'pdf') {
    return {
      label: 'Số trang',
      value: `${doc.pageCount ?? doc.pdfPages?.length ?? 1} trang`,
    };
  }
  return {
    label: 'Số trang',
    value: `${doc.pageCount ?? 1} trang`,
  };
}

function resolveOriginalStatusLabel(doc: DocumentItem): string {
  const hasOrigin = Boolean(
    (doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) ||
      doc.storagePath ||
      doc.r2ObjectKey ||
      doc.fileUrl
  );
  if (!hasOrigin) {
    return 'Tạo mới trực tiếp trên LITE';
  }
  return doc.hasLocalEdits
    ? 'Có bản gốc (đã chỉnh sửa cục bộ)'
    : 'Nguyên vẹn 100% từ tệp gốc';
}

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
  modeSwitch,
  onSave,
  isSaving = false,
  isSaved = false,
  isDirty = false,
  onDownloadEdited,
  onExportPdf,
  extraMenuItems,
}: WorkspaceTopBarProps) {
  const isDark = theme === 'dark';
  const visual = FORMAT_VISUALS[document.type];
  const FormatIcon = visual.Icon;
  const isReadOnlyCollab = collabPermission === 'view';
  const showModeSwitch = Boolean(modeSwitch && modeSwitch.visible !== false);

  const {
    driveSyncBadge,
    driveSyncLabel,
    pendingOfflineCount,
    onOpenVersionHistory,
    onOpenQuickConverter,
    onDownloadOriginalFile,
  } = useWorkspaceSync();

  const [isEditingTitle, setIsEditingTitle] = useState<boolean>(false);
  const [titleDraft, setTitleDraft] = useState<string>(document.title);
  const titleInputRef = useRef<HTMLInputElement | null>(null);
  const cancelBlurRef = useRef<boolean>(false);

  // Overflow menu (`⋮`) & file info modal state
  const [isMoreMenuOpen, setIsMoreMenuOpen] = useState<boolean>(false);
  const [isFileInfoOpen, setIsFileInfoOpen] = useState<boolean>(false);
  const [isSavingLocal, setIsSavingLocal] = useState<boolean>(false);
  const [driveStatus, setDriveStatus] = useState<DriveUploadStatus>('idle');
  const [driveMessage, setDriveMessage] = useState<string | null>(null);
  const [driveViewLink, setDriveViewLink] = useState<string | null>(null);
  const [isConnectingGoogle, setIsConnectingGoogle] = useState<boolean>(false);
  const [isDownloadingLocal, setIsDownloadingLocal] = useState<boolean>(false);
  const [isDownloadingOriginal, setIsDownloadingOriginal] =
    useState<boolean>(false);
  const [isExportingPdf, setIsExportingPdf] = useState<boolean>(false);
  const [localFileNotice, setLocalFileNotice] = useState<{
    type: 'success' | 'error';
    title: string;
    message: string;
  } | null>(null);
  const localFileNoticeTimerRef = useRef<number | null>(null);
  const moreMenuRef = useRef<HTMLDivElement | null>(null);

  const showLocalFileNotice = (
    type: 'success' | 'error',
    title: string,
    message: string
  ) => {
    setLocalFileNotice({ type, title, message });
    if (localFileNoticeTimerRef.current !== null) {
      window.clearTimeout(localFileNoticeTimerRef.current);
      localFileNoticeTimerRef.current = null;
    }
    if (type === 'success') {
      localFileNoticeTimerRef.current = window.setTimeout(() => {
        setLocalFileNotice(null);
        localFileNoticeTimerRef.current = null;
      }, 5500);
    }
  };

  useEffect(() => {
    return () => {
      if (localFileNoticeTimerRef.current !== null) {
        window.clearTimeout(localFileNoticeTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    setTitleDraft(document.title);
    setIsEditingTitle(false);
    setIsMoreMenuOpen(false);
    setIsFileInfoOpen(false);
    setDriveStatus('idle');
    setDriveMessage(null);
    setDriveViewLink(null);
    setLocalFileNotice(null);
  }, [document.id, document.title]);

  useEffect(() => {
    if (isEditingTitle && titleInputRef.current) {
      titleInputRef.current.focus();
      titleInputRef.current.select();
    }
  }, [isEditingTitle]);

  // Close `⋮` dropdown on outside click or Escape
  useEffect(() => {
    if (!isMoreMenuOpen) {
      return;
    }
    const handlePointerDown = (event: MouseEvent) => {
      if (
        moreMenuRef.current &&
        !moreMenuRef.current.contains(event.target as Node)
      ) {
        setIsMoreMenuOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setIsMoreMenuOpen(false);
      }
    };
    window.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isMoreMenuOpen]);

  // Close file info dialog on Escape
  useEffect(() => {
    if (!isFileInfoOpen) {
      return;
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setIsFileInfoOpen(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isFileInfoOpen]);

  const startEditingTitle = () => {
    if (!onRenameTitle || isReadOnlyCollab) {
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

  const handlePrimarySave = async () => {
    if (!onSave || isSaving || isSavingLocal) {
      return;
    }
    setIsSavingLocal(true);
    try {
      await onSave();
    } finally {
      setIsSavingLocal(false);
    }
  };

  const handleConnectGoogleDrive = async () => {
    setIsConnectingGoogle(true);
    setDriveMessage(null);
    try {
      const result = await requestGoogleDriveAccess({
        userId: document.userId,
      });
      if (result.error) {
        setDriveStatus('error');
        setDriveMessage(
          result.error.message || 'Không thể khởi tạo kết nối Google Drive.'
        );
      } else if (result.driveConnected) {
        setDriveStatus('idle');
        setDriveMessage(
          'Đã kết nối Google Drive thành công! Đang sẵn sàng lưu tệp.'
        );
        await handleSaveToGoogleDrive();
      } else if (result.noticeMessage) {
        setDriveStatus('error');
        setDriveMessage(result.noticeMessage);
      } else if (result.openedExternally) {
        setDriveMessage(
          'Đã mở trình duyệt để cấp quyền Google Drive. Sau khi xác thực, hệ thống sẽ tự quay lại LITE.'
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
    let token = getStoredGoogleDriveToken(document.userId);
    if (!token) {
      setIsConnectingGoogle(true);
      setDriveMessage(null);
      try {
        const grantResult = await requestGoogleDriveAccess({
          userId: document.userId,
        });
        if (grantResult.openedExternally) {
          setDriveStatus('auth_required');
          setDriveMessage(
            'Đã mở trình duyệt để cấp quyền Google Drive. Sau khi xác nhận, bấm Lưu lên Google Drive lại.'
          );
          return;
        }
        if (grantResult.error || !grantResult.driveConnected) {
          setDriveStatus('auth_required');
          setDriveMessage(
            grantResult.error?.message ||
              'Bạn chưa cấp quyền Google Drive. Bấm Kết nối Google Drive để tải tệp lên Drive.'
          );
          setDriveViewLink(null);
          return;
        }
        token =
          grantResult.driveAccessToken ??
          getStoredGoogleDriveToken(document.userId);
      } finally {
        setIsConnectingGoogle(false);
      }
    }

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
      const uploaded = await upsertDocumentToLiteDrive(
        document.id,
        blob,
        fileName,
        mimeType,
        document.userId
      );
      setDriveStatus('success');
      setDriveMessage(
        uploaded.updatedExisting
          ? `Đã cập nhật trên Drive (LITE Workspace): ${uploaded.name || fileName}`
          : `Đã lưu vào LITE Workspace: ${uploaded.name || fileName}`
      );
      setDriveViewLink(uploaded.webViewLink ?? null);
    } catch (err: unknown) {
      if (err instanceof GoogleDriveAuthError) {
        setDriveStatus('auth_required');
        setDriveMessage(
          (err instanceof Error ? err.message : String(err)) ||
            'Phiên truy cập Google Drive đã hết hạn. Vui lòng kết nối lại Google Drive.'
        );
      } else {
        setDriveStatus('error');
        setDriveMessage(
          err instanceof Error
            ? (err instanceof Error ? err.message : String(err))
            : 'Không thể lưu tệp lên Google Drive. Vui lòng thử lại.'
        );
      }
    }
  };

  const isAnyFileBusy =
    isDownloadingOriginal || isDownloadingLocal || isExportingPdf;

  const handleDownloadOriginalFile = async () => {
    if (isAnyFileBusy) {
      return;
    }
    setIsDownloadingOriginal(true);
    setDriveMessage(null);
    setLocalFileNotice(null);
    try {
      if (onDownloadOriginalFile) {
        await onDownloadOriginalFile();
        setIsMoreMenuOpen(false);
        showLocalFileNotice(
          'success',
          'Đã tải xuống bản gốc',
          'Tệp nhị phân gốc ban đầu đã được xử lý tải xuống thiết bị.'
        );
      } else {
        const result = await exportOriginalDocumentBinary(document);
        setIsMoreMenuOpen(false);
        const isCanceled = result.method === 'electron' && !result.savedPath;
        showLocalFileNotice(
          'success',
          isCanceled ? 'Thao tác lưu tệp' : 'Đã tải xuống bản gốc',
          result.message
        );
      }
    } catch (err: unknown) {
      showLocalFileNotice(
        'error',
        'Không thể tải bản gốc',
        err instanceof Error
          ? (err instanceof Error ? err.message : String(err))
          : 'Không thể tải xuống bản gốc của tệp.'
      );
    } finally {
      setIsDownloadingOriginal(false);
    }
  };

  const handleDownloadEditedFile = async () => {
    if (isAnyFileBusy) {
      return;
    }
    setIsDownloadingLocal(true);
    setDriveMessage(null);
    setLocalFileNotice(null);
    try {
      if (onDownloadEdited) {
        await onDownloadEdited();
        setIsMoreMenuOpen(false);
      } else {
        const { blob, fileName, mimeType } =
          await buildDocumentUploadPayload(document);
        const result = await saveFileUniversal(blob, fileName, {
          mimeType,
          shareTitle: document.title,
        });
        setIsMoreMenuOpen(false);
        const isCanceled = result.method === 'electron' && !result.savedPath;
        showLocalFileNotice(
          'success',
          isCanceled ? 'Thao tác lưu tệp' : 'Đã lưu về thiết bị',
          result.message
        );
      }
    } catch (err: unknown) {
      showLocalFileNotice(
        'error',
        'Không thể tải tệp xuống',
        err instanceof Error
          ? (err instanceof Error ? err.message : String(err))
          : 'Không thể tải tệp xuống thiết bị.'
      );
    } finally {
      setIsDownloadingLocal(false);
    }
  };

  const handleExportPdfAction = async () => {
    if (isAnyFileBusy) {
      return;
    }
    setIsExportingPdf(true);
    setDriveMessage(null);
    setLocalFileNotice(null);
    try {
      if (onExportPdf) {
        await onExportPdf();
        setIsMoreMenuOpen(false);
      } else {
        const result = await downloadDocumentAsPdf(document);
        setIsMoreMenuOpen(false);
        const isCanceled = result.method === 'electron' && !result.savedPath;
        showLocalFileNotice(
          'success',
          isCanceled ? 'Thao tác xuất PDF' : 'Đã xuất sang PDF',
          result.message
        );
      }
    } catch (err: unknown) {
      showLocalFileNotice(
        'error',
        'Không thể xuất PDF',
        err instanceof Error
          ? (err instanceof Error ? err.message : String(err))
          : 'Không thể xuất tài liệu sang định dạng PDF.'
      );
    } finally {
      setIsExportingPdf(false);
    }
  };

  // Resolve compact state of the primary Save button (`Lưu`)
  const saveVisualState = (() => {
    if (isSaving || isSavingLocal || driveSyncBadge === 'saving') {
      return 'saving' as const;
    }
    if (driveSyncBadge === 'error') {
      return 'error' as const;
    }
    if (isDirty) {
      return 'dirty' as const;
    }
    if (driveSyncBadge === 'offline' || pendingOfflineCount > 0) {
      return 'offline' as const;
    }
    if (isSaved || driveSyncBadge === 'saved') {
      return 'saved' as const;
    }
    return 'idle' as const;
  })();

  const saveButtonLabel =
    saveVisualState === 'saving'
      ? 'Đang lưu'
      : saveVisualState === 'saved'
        ? 'Đã lưu'
        : saveVisualState === 'offline'
          ? 'Chờ đồng bộ'
          : saveVisualState === 'error'
            ? 'Lỗi lưu'
            : 'Lưu';

  const saveButtonTooltip =
    driveSyncLabel.trim() ||
    (saveVisualState === 'saving'
      ? 'Đang lưu & đồng bộ...'
      : saveVisualState === 'saved'
        ? 'Đã lưu đồng bộ'
        : saveVisualState === 'offline'
          ? pendingOfflineCount > 0
            ? `Đang chờ đồng bộ ngoại tuyến (${pendingOfflineCount})`
            : 'Đang chờ đồng bộ ngoại tuyến'
          : saveVisualState === 'error'
            ? 'Lưu / đồng bộ gặp lỗi — bấm để lưu lại'
            : isDirty
              ? 'Có thay đổi chưa lưu (Ctrl+S)'
              : 'Lưu thay đổi (Ctrl+S)');

  const saveButtonClass =
    saveVisualState === 'saved'
      ? isDark
        ? 'bg-emerald-500/20 border-emerald-400/40 text-emerald-200 hover:bg-emerald-500/25'
        : 'bg-emerald-50 border-emerald-200 text-emerald-700 hover:bg-emerald-100/80'
      : saveVisualState === 'offline'
        ? isDark
          ? 'bg-amber-500/20 border-amber-400/40 text-amber-200 hover:bg-amber-500/25'
          : 'bg-amber-50 border-amber-200 text-amber-800 hover:bg-amber-100/80'
        : saveVisualState === 'error'
          ? isDark
            ? 'bg-rose-500/20 border-rose-400/40 text-rose-200 hover:bg-rose-500/25'
            : 'bg-rose-50 border-rose-200 text-rose-700 hover:bg-rose-100/80'
          : saveVisualState === 'dirty'
            ? visual.saveDirtyClass
            : isDark
              ? 'bg-white/[0.06] hover:bg-white/[0.12] border-white/12 text-slate-200'
              : 'bg-slate-50/90 hover:bg-slate-100 border-slate-200/80 text-slate-700';

  const countSummary = resolveDocumentCountSummary(document);

  return (
    <>
      <header
        className={`h-[52px] shrink-0 w-full px-1.5 sm:px-3.5 flex items-center justify-between gap-1 sm:gap-2 overflow-visible border-b transition-colors duration-200 z-20 relative ${
          isDark
            ? 'bg-[#141923] border-white/10 text-slate-100'
            : 'bg-white border-slate-200/80 shadow-[0_1px_2px_rgba(15,23,42,0.03)] text-slate-900'
        }`}
      >
        {/* 1. LEFT CLUSTER: Back + leftExtra + Truncated File Title (no pencil icon) */}
        <div className="flex-1 min-w-[64px] flex items-center gap-1 sm:gap-1.5 overflow-hidden">
          <button
            type="button"
            onClick={onBackHome}
            aria-label="Quay lại Trang chủ"
            title="Quay lại Trang chủ"
            className={`shrink-0 inline-flex items-center justify-center gap-1.5 h-8 sm:h-9 w-8 sm:w-auto sm:px-2.5 rounded-xl text-xs font-medium border transition-colors whitespace-nowrap ${
              isDark
                ? 'bg-white/[0.05] hover:bg-white/[0.12] border-white/10 text-slate-200 hover:text-white'
                : 'bg-slate-50/90 hover:bg-slate-100 border-slate-200/80 text-slate-700 hover:text-slate-950'
            }`}
          >
            <ArrowLeft className="h-4 w-4 shrink-0" />
            <span className="hidden sm:inline">Trang chủ</span>
          </button>

          {leftExtra && (
            <div className="shrink-0 flex items-center">{leftExtra}</div>
          )}

          <div className="min-w-0 flex-1 flex items-center gap-1.5">
            <span
              aria-hidden="true"
              className={`hidden md:flex shrink-0 h-7 w-7 rounded-lg border items-center justify-center ${
                isDark ? visual.iconBoxDark : visual.iconBoxLight
              }`}
            >
              <FormatIcon className="h-3.5 w-3.5" />
            </span>

            {onRenameTitle && !isReadOnlyCollab && isEditingTitle ? (
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
                className={`min-w-[100px] w-full max-w-[260px] h-8 px-2.5 rounded-lg border text-[13.5px] sm:text-sm font-semibold tracking-tight outline-none transition-all ${
                  isDark
                    ? 'bg-[#1C2333] border-indigo-400 text-white placeholder:text-slate-500 focus:ring-2 focus:ring-indigo-400/25'
                    : visual.focusInputLight
                }`}
              />
            ) : onRenameTitle && !isReadOnlyCollab ? (
              <button
                type="button"
                onClick={startEditingTitle}
                title={`${document.title} (Bấm để đổi tên)`}
                aria-label={`Đổi tên tệp: ${document.title}`}
                className={`min-w-0 max-w-full inline-flex items-center px-1.5 py-1 -mx-1 rounded-lg text-left transition-colors ${
                  isDark
                    ? 'hover:bg-white/[0.08] text-white'
                    : 'hover:bg-slate-100/80 text-slate-900'
                }`}
              >
                <h1 className="truncate text-[13.5px] sm:text-sm font-semibold tracking-tight">
                  {document.title}
                </h1>
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
              className={`hidden md:inline-flex shrink-0 items-center px-1.5 py-0.5 rounded-md border text-[10px] font-bold tracking-wider ${
                isDark ? visual.badgeDark : visual.badgeLight
              }`}
            >
              {visual.label}
            </span>
          </div>
        </div>

        {/* 2. CENTER CLUSTER: Compact Gốc | Sửa pill OR slim workspace toolbar (PDF page pill, etc.) */}
        {(showModeSwitch || toolbar) && (
          <div className="shrink-0 flex items-center gap-1.5">
            {showModeSwitch && modeSwitch && (
              <div
                role="group"
                aria-label="Chế độ xem và chỉnh sửa"
                className={`shrink-0 inline-flex items-center p-0.5 rounded-xl border ${
                  isDark
                    ? 'bg-white/[0.04] border-white/10'
                    : 'bg-slate-100/90 border-slate-200/80'
                }`}
              >
                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => modeSwitch.onChange('original')}
                  aria-pressed={modeSwitch.mode === 'original'}
                  title={
                    modeSwitch.originalTitle ||
                    'Xem nguyên bản (khóa sửa nhầm, giữ chuẩn định dạng gốc)'
                  }
                  className={`inline-flex items-center gap-1 h-7 sm:h-[30px] px-1.5 sm:px-2.5 rounded-[9px] text-[11px] sm:text-xs font-semibold transition-all whitespace-nowrap shrink-0 ${
                    modeSwitch.mode === 'original'
                      ? isDark
                        ? visual.modeActiveDark
                        : visual.modeActiveLight
                      : isDark
                        ? 'text-slate-300 hover:text-white'
                        : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <Eye className="h-3.5 w-3.5 shrink-0" />
                  <span>Gốc</span>
                </button>
                <button
                  type="button"
                  disabled={isReadOnlyCollab}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => modeSwitch.onChange('edit')}
                  aria-pressed={modeSwitch.mode === 'edit'}
                  title={
                    modeSwitch.editTitle ||
                    (isReadOnlyCollab
                      ? 'Tài liệu đang ở chế độ chỉ xem'
                      : 'Chuyển sang Chỉnh sửa nhanh')
                  }
                  className={`inline-flex items-center gap-1 h-7 sm:h-[30px] px-1.5 sm:px-2.5 rounded-[9px] text-[11px] sm:text-xs font-semibold transition-all whitespace-nowrap shrink-0 disabled:opacity-45 disabled:cursor-not-allowed ${
                    modeSwitch.mode === 'edit' && !isReadOnlyCollab
                      ? isDark
                        ? visual.modeActiveDark
                        : visual.modeActiveLight
                      : isDark
                        ? 'text-slate-300 hover:text-white'
                        : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <Pencil className="h-3.5 w-3.5 shrink-0" />
                  <span>Sửa</span>
                </button>
              </div>
            )}

            {toolbar && (
              <div className="hidden sm:flex shrink-0 items-center gap-1">
                {toolbar}
              </div>
            )}
          </div>
        )}

        {/* 3. RIGHT CLUSTER: Collaborators (md+) + Save (`Lưu`) + Share (`Chia sẻ`) + More (`⋮`) */}
        <div className="shrink-0 flex items-center gap-1 sm:gap-1.5 whitespace-nowrap">
          <div className="hidden md:flex items-center shrink-0">
            <CollaboratorPresenceBar
              collaborators={collaborators}
              selfSessionId={selfSessionId}
              collabPermission={collabPermission}
              theme={theme}
            />
          </div>

          {onSave && (
            <button
              type="button"
              onClick={() => void handlePrimarySave()}
              disabled={isSaving || isSavingLocal}
              title={saveButtonTooltip}
              aria-label={saveButtonTooltip}
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 h-8 w-8 p-0 overflow-hidden md:h-9 md:w-auto md:px-3 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap disabled:opacity-70 ${saveButtonClass}`}
            >
              {saveVisualState === 'saving' ? (
                <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
              ) : saveVisualState === 'saved' ? (
                <Check className="h-3.5 w-3.5 shrink-0" />
              ) : saveVisualState === 'offline' ? (
                <CloudOff className="h-3.5 w-3.5 shrink-0" />
              ) : saveVisualState === 'error' ? (
                <AlertCircle className="h-3.5 w-3.5 shrink-0" />
              ) : (
                <span className="relative flex items-center justify-center shrink-0">
                  <Save className="h-3.5 w-3.5 shrink-0" />
                  {isDirty && (
                    <span className="md:hidden absolute -top-1 -right-1 h-2 w-2 rounded-full bg-amber-400 ring-1 ring-white" />
                  )}
                </span>
              )}
              <span className="hidden md:inline">{saveButtonLabel}</span>
            </button>
          )}

          <button
            type="button"
            onClick={onShare}
            aria-label="Chia sẻ"
            title="Chia sẻ"
            className={`shrink-0 inline-flex items-center justify-center gap-1.5 h-8 w-8 p-0 overflow-hidden md:h-9 md:w-auto md:px-3 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
              isDark
                ? 'bg-white/[0.1] hover:bg-white/[0.18] border-white/15 text-white'
                : visual.shareButtonLight
            }`}
          >
            <Share2 className="h-3.5 w-3.5 shrink-0" />
            <span className="hidden md:inline">Chia sẻ</span>
          </button>

          {/* Single Overflow Menu `⋮` */}
          <div ref={moreMenuRef} className="relative shrink-0">
            <button
              type="button"
              onClick={() => setIsMoreMenuOpen((prev) => !prev)}
              aria-expanded={isMoreMenuOpen}
              aria-haspopup="menu"
              aria-label="Thêm tùy chọn"
              title="Thêm tùy chọn"
              className={`shrink-0 inline-flex items-center justify-center h-8 sm:h-9 w-8 sm:w-9 rounded-xl border transition-colors whitespace-nowrap ${
                isMoreMenuOpen
                  ? isDark
                    ? 'bg-white/[0.16] border-white/25 text-white'
                    : 'bg-slate-200/80 border-slate-300 text-slate-950'
                  : isDark
                    ? 'bg-white/[0.06] hover:bg-white/[0.14] border-white/12 text-slate-200 hover:text-white'
                    : 'bg-slate-50/90 hover:bg-slate-100 border-slate-200/80 text-slate-700 hover:text-slate-950'
              }`}
            >
              <MoreVertical className="h-4 w-4 shrink-0" />
            </button>

            {isMoreMenuOpen && (
              <div
                role="menu"
                aria-label="Tùy chọn tài liệu"
                className={`absolute right-0 top-full mt-2 w-72 max-w-[calc(100vw-1rem)] max-h-[calc(100dvh-4.5rem)] overflow-y-auto rounded-2xl border shadow-2xl p-2 z-50 ${
                  isDark
                    ? 'bg-[#19202D] border-white/15 text-slate-100'
                    : 'bg-white border-[#E7E2DA] text-slate-900'
                }`}
              >
                <div className="space-y-0.5">
                  <button
                    type="button"
                    role="menuitem"
                    disabled={isAnyFileBusy}
                    onClick={() => void handleDownloadOriginalFile()}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-55 ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-100'
                        : 'hover:bg-[#F7F5F0] text-slate-800'
                    }`}
                  >
                    {isDownloadingOriginal ? (
                      <Loader2 className="h-4 w-4 animate-spin text-sky-500 shrink-0" />
                    ) : (
                      <FileDown className="h-4 w-4 text-sky-500 shrink-0" />
                    )}
                    <span className="truncate">
                      {isDownloadingOriginal
                        ? 'Đang chuẩn bị bản gốc...'
                        : 'Tải bản gốc'}
                    </span>
                  </button>

                  <button
                    type="button"
                    role="menuitem"
                    disabled={isAnyFileBusy}
                    onClick={() => void handleDownloadEditedFile()}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-55 ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-100'
                        : 'hover:bg-[#F7F5F0] text-slate-800'
                    }`}
                  >
                    {isDownloadingLocal ? (
                      <Loader2 className="h-4 w-4 animate-spin text-indigo-500 shrink-0" />
                    ) : (
                      <Download className="h-4 w-4 text-indigo-500 shrink-0" />
                    )}
                    <span className="truncate">
                      {isDownloadingLocal
                        ? 'Đang tải xuống...'
                        : `Tải xuống bản sửa (.${document.type})`}
                    </span>
                  </button>

                  <button
                    type="button"
                    role="menuitem"
                    disabled={isAnyFileBusy}
                    onClick={() => void handleExportPdfAction()}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-55 ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-100'
                        : 'hover:bg-[#F7F5F0] text-slate-800'
                    }`}
                  >
                    {isExportingPdf ? (
                      <Loader2 className="h-4 w-4 animate-spin text-rose-500 shrink-0" />
                    ) : (
                      <FileText className="h-4 w-4 text-rose-500 shrink-0" />
                    )}
                    <span className="truncate">
                      {isExportingPdf ? 'Đang xuất PDF...' : 'Xuất PDF'}
                    </span>
                  </button>

                  <button
                    type="button"
                    role="menuitem"
                    disabled={driveStatus === 'uploading'}
                    onClick={() => void handleSaveToGoogleDrive()}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-55 ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-100'
                        : 'hover:bg-[#F7F5F0] text-slate-800'
                    }`}
                  >
                    {driveStatus === 'uploading' ? (
                      <Loader2 className="h-4 w-4 animate-spin text-emerald-500 shrink-0" />
                    ) : (
                      <HardDrive className="h-4 w-4 text-emerald-500 shrink-0" />
                    )}
                    <span className="truncate">
                      {driveStatus === 'uploading'
                        ? 'Đang lưu lên Drive...'
                        : 'Lưu lên Google Drive'}
                    </span>
                  </button>

                  <div
                    className={`my-1 border-t ${
                      isDark ? 'border-white/10' : 'border-slate-100'
                    }`}
                  />

                  <button
                    type="button"
                    role="menuitem"
                    disabled={!onOpenVersionHistory}
                    onClick={() => {
                      setIsMoreMenuOpen(false);
                      onOpenVersionHistory?.();
                    }}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-45 disabled:cursor-not-allowed ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-200'
                        : 'hover:bg-[#F7F5F0] text-slate-700'
                    }`}
                  >
                    <History className="h-4 w-4 text-indigo-500 shrink-0" />
                    <span className="truncate">Lịch sử phiên bản</span>
                  </button>

                  <button
                    type="button"
                    role="menuitem"
                    disabled={!onOpenQuickConverter}
                    onClick={() => {
                      setIsMoreMenuOpen(false);
                      onOpenQuickConverter?.();
                    }}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-45 disabled:cursor-not-allowed ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-200'
                        : 'hover:bg-[#F7F5F0] text-slate-700'
                    }`}
                  >
                    <FileOutput className="h-4 w-4 text-purple-500 shrink-0" />
                    <span className="truncate">Chuyển đổi định dạng</span>
                  </button>

                  {extraMenuItems && extraMenuItems.length > 0 && (
                    <>
                      <div
                        className={`my-1 border-t ${
                          isDark ? 'border-white/10' : 'border-slate-100'
                        }`}
                      />
                      {extraMenuItems.map((item) => (
                        <button
                          key={item.id}
                          type="button"
                          role="menuitem"
                          disabled={item.disabled}
                          onClick={() => {
                            setIsMoreMenuOpen(false);
                            item.onClick();
                          }}
                          className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors disabled:opacity-45 disabled:cursor-not-allowed ${
                            isDark
                              ? 'hover:bg-white/[0.08] text-slate-200'
                              : 'hover:bg-[#F7F5F0] text-slate-700'
                          }`}
                        >
                          {item.icon && (
                            <span className="shrink-0 flex items-center justify-center">
                              {item.icon}
                            </span>
                          )}
                          <span className="truncate">{item.label}</span>
                        </button>
                      ))}
                    </>
                  )}

                  <div
                    className={`my-1 border-t ${
                      isDark ? 'border-white/10' : 'border-slate-100'
                    }`}
                  />

                  <button
                    type="button"
                    role="menuitem"
                    onClick={() => {
                      setIsMoreMenuOpen(false);
                      setIsFileInfoOpen(true);
                    }}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-left text-xs font-semibold transition-colors ${
                      isDark
                        ? 'hover:bg-white/[0.08] text-slate-200'
                        : 'hover:bg-[#F7F5F0] text-slate-700'
                    }`}
                  >
                    <Info className="h-4 w-4 text-slate-400 shrink-0" />
                    <span className="truncate">Thông tin tệp</span>
                  </button>
                </div>

                {/* Inline Google Drive status feedback inside `⋮` menu */}
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
                        <p className="font-semibold">Đã lưu lên Google Drive</p>
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

                {localFileNotice && (
                  <div
                    role={localFileNotice.type === 'error' ? 'alert' : 'status'}
                    className={`mt-2 p-2.5 rounded-xl border text-xs ${
                      localFileNotice.type === 'error'
                        ? isDark
                          ? 'bg-rose-500/15 border-rose-400/25 text-rose-200'
                          : 'bg-rose-50 border-rose-200 text-rose-700'
                        : isDark
                          ? 'bg-emerald-500/15 border-emerald-400/25 text-emerald-200'
                          : 'bg-emerald-50 border-emerald-200 text-emerald-800'
                    }`}
                  >
                    <p className="font-semibold">{localFileNotice.title}</p>
                    <p className="text-[11px] opacity-90 break-words mt-0.5">
                      {localFileNotice.message}
                    </p>
                  </div>
                )}
              </div>
            )}

            {/* Toast phản hồi lưu/tải/xuất PDF hiển thị sau khi menu `⋮` đóng */}
            {!isMoreMenuOpen && localFileNotice && (
              <div
                role={localFileNotice.type === 'error' ? 'alert' : 'status'}
                className={`absolute right-0 top-full mt-2 w-72 max-w-[calc(100vw-1rem)] rounded-2xl border shadow-xl p-3 z-50 text-xs flex items-start gap-2.5 ${
                  localFileNotice.type === 'error'
                    ? isDark
                      ? 'bg-[#2A1720] border-rose-400/30 text-rose-100'
                      : 'bg-rose-50 border-rose-200 text-rose-800'
                    : isDark
                      ? 'bg-[#172824] border-emerald-400/30 text-emerald-100'
                      : 'bg-emerald-50 border-emerald-200 text-emerald-900'
                }`}
              >
                {localFileNotice.type === 'error' ? (
                  <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
                ) : (
                  <Check className="h-4 w-4 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0 flex-1">
                  <p className="font-semibold">{localFileNotice.title}</p>
                  <p className="text-[11px] opacity-90 break-words mt-0.5">
                    {localFileNotice.message}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setLocalFileNotice(null)}
                  aria-label="Đóng thông báo"
                  className="p-0.5 rounded-md opacity-70 hover:opacity-100 transition-opacity shrink-0"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* File Info Modal (`Thông tin tệp`) */}
      {isFileInfoOpen && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="topbar-file-info-title"
          onClick={() => setIsFileInfoOpen(false)}
          className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-[2px] flex items-center justify-center p-4 pb-safe"
        >
          <div
            onClick={(event) => event.stopPropagation()}
            className={`w-full max-w-md rounded-2xl sm:rounded-3xl border shadow-2xl overflow-hidden ${
              isDark
                ? 'bg-[#181F2C] border-white/15 text-slate-100'
                : 'bg-[#FFFDF9] border-[#E7E2DA] text-slate-900'
            }`}
          >
            <div
              className={`px-5 py-4 border-b flex items-center justify-between gap-3 ${
                isDark ? 'border-white/10' : 'border-[#EDE8DF]'
              }`}
            >
              <div className="flex items-center gap-2.5 min-w-0">
                <span
                  className={`w-9 h-9 rounded-xl border flex items-center justify-center shrink-0 ${
                    isDark ? visual.iconBoxDark : visual.iconBoxLight
                  }`}
                >
                  <Info className="h-4 w-4" />
                </span>
                <div className="min-w-0">
                  <h2
                    id="topbar-file-info-title"
                    className="text-sm font-bold truncate"
                  >
                    Thông tin tệp
                  </h2>
                  <p
                    className={`text-xs truncate ${
                      isDark ? 'text-slate-400' : 'text-slate-500'
                    }`}
                  >
                    Chi tiết tài liệu hiện tại
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsFileInfoOpen(false)}
                aria-label="Đóng thông tin tệp"
                className={`p-1.5 rounded-xl transition-colors ${
                  isDark
                    ? 'text-slate-400 hover:text-white hover:bg-white/10'
                    : 'text-slate-400 hover:text-slate-700 hover:bg-slate-100'
                }`}
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="p-5 space-y-3 text-xs">
              <div className="flex items-start justify-between gap-4">
                <span className={isDark ? 'text-slate-400' : 'text-slate-500'}>
                  Tên tệp
                </span>
                <span className="font-semibold text-right break-words max-w-[65%]">
                  {document.title}
                </span>
              </div>
              <div className="flex items-center justify-between gap-4">
                <span className={isDark ? 'text-slate-400' : 'text-slate-500'}>
                  Loại
                </span>
                <span
                  className={`inline-flex items-center px-2 py-0.5 rounded-md border text-[11px] font-bold ${
                    isDark ? visual.badgeDark : visual.badgeLight
                  }`}
                >
                  {visual.label} (.{document.type})
                </span>
              </div>
              <div className="flex items-center justify-between gap-4">
                <span className={isDark ? 'text-slate-400' : 'text-slate-500'}>
                  Dung lượng
                </span>
                <span className="font-semibold">
                  {document.fileSize || 'Chưa có thông tin'}
                </span>
              </div>
              <div className="flex items-center justify-between gap-4">
                <span className={isDark ? 'text-slate-400' : 'text-slate-500'}>
                  {countSummary.label}
                </span>
                <span className="font-semibold">{countSummary.value}</span>
              </div>
              <div className="flex items-center justify-between gap-4">
                <span className={isDark ? 'text-slate-400' : 'text-slate-500'}>
                  Cập nhật
                </span>
                <span className="font-semibold">{document.updatedAt}</span>
              </div>
              <div
                className={`flex items-start justify-between gap-4 pt-2 border-t border-dashed ${
                  isDark ? 'border-white/10' : 'border-slate-200'
                }`}
              >
                <span className={isDark ? 'text-slate-400' : 'text-slate-500'}>
                  Trạng thái bản gốc
                </span>
                <span
                  className={`font-semibold text-right max-w-[65%] ${
                    isDark ? 'text-indigo-300' : 'text-indigo-600'
                  }`}
                >
                  {resolveOriginalStatusLabel(document)}
                </span>
              </div>
            </div>

            <div
              className={`px-5 py-3 border-t flex justify-end ${
                isDark ? 'border-white/10' : 'border-[#EDE8DF]'
              }`}
            >
              <button
                type="button"
                onClick={() => setIsFileInfoOpen(false)}
                className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm transition-colors"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
