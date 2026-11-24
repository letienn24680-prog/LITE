import { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  BookOpen,
  Check,
  CheckCircle2,
  Copy,
  Download,
  ExternalLink,
  Eye,
  FileText,
  FolderOpen,
  Globe,
  KeyRound,
  Layers,
  Link2,
  Loader2,
  Lock,
  Maximize2,
  Presentation,
  QrCode,
  Save,
  Share2,
  ShieldOff,
  Smartphone,
  X,
} from 'lucide-react';
import { RealQrCode } from '../common/RealQrCode';
import { PresenterStageView } from './PresenterStageView';
import {
  buildPublicShareUrl,
  getExistingShareForDocument,
  revokeDocumentShare,
  saveDocumentShareSettings,
} from '../../services/shareService';
import { resolvePublicWebOrigin } from '../../services/realtimeCollab';
import type {
  DocumentItem,
  DocumentType,
  ModalType,
  ShareAccessLevel,
} from '../../types/document';

export interface ShareConfigUpdatePayload {
  shareToken?: string;
  shareAccess: ShareAccessLevel;
  shareAllowDownload: boolean;
  isShared: boolean;
}

export interface StageActionModalProps {
  modal: ModalType;
  document: DocumentItem | null;
  onClose: () => void;
  onSwitchModal?: (modal: Exclude<ModalType, null>) => void;
  onShareConfigUpdated?: (
    docId: string,
    config: ShareConfigUpdatePayload
  ) => void;
}

const TYPE_BADGE_STYLES: Record<
  DocumentType,
  { label: string; className: string }
> = {
  pptx: {
    label: 'PPTX',
    className: 'bg-orange-50 text-orange-700 border-orange-200/80',
  },
  docx: {
    label: 'DOCX',
    className: 'bg-blue-50 text-blue-700 border-blue-200/80',
  },
  xlsx: {
    label: 'XLSX',
    className: 'bg-green-100 text-green-800 border-green-300/80',
  },
  pdf: {
    label: 'PDF',
    className: 'bg-pink-50 text-pink-700 border-pink-200/80',
  },
};

function getDocumentUnit(type: DocumentType): string {
  if (type === 'pptx') {
    return 'slide';
  }
  if (type === 'xlsx') {
    return 'sheet';
  }
  return 'trang';
}

function buildPairingCode(docId: string): string {
  let hash = 1357;
  for (let i = 0; i < docId.length; i += 1) {
    hash = (hash * 31 + docId.charCodeAt(i)) % 9000;
  }
  const numeric = 1000 + hash;
  return `LITE-${numeric}`;
}

export function StageActionModal({
  modal,
  document,
  onClose,
  onSwitchModal,
  onShareConfigUpdated,
}: StageActionModalProps) {
  const [outlineIdx, setOutlineIdx] = useState<number>(0);
  const [codeCopied, setCodeCopied] = useState<boolean>(false);
  const [remoteUrlCopied, setRemoteUrlCopied] = useState<boolean>(false);

  // Real Share Modal states
  const [shareAccess, setShareAccess] = useState<ShareAccessLevel>('private');
  const [savedShareAccess, setSavedShareAccess] =
    useState<ShareAccessLevel>('private');
  const [sharePassword, setSharePassword] = useState<string>('');
  const [allowDownload, setAllowDownload] = useState<boolean>(true);
  const [shareToken, setShareToken] = useState<string>('');
  const [shareViewCount, setShareViewCount] = useState<number>(0);
  const [shareRevokedAt, setShareRevokedAt] = useState<string | null>(null);
  const [isLoadingShare, setIsLoadingShare] = useState<boolean>(false);
  const [isSavingShare, setIsSavingShare] = useState<boolean>(false);
  const [isRevokingShare, setIsRevokingShare] = useState<boolean>(false);
  const [shareError, setShareError] = useState<string | null>(null);
  const [shareFeedback, setShareFeedback] = useState<string | null>(null);
  const [linkCopied, setLinkCopied] = useState<boolean>(false);

  const codeCopyTimerRef = useRef<number | null>(null);
  const remoteUrlCopyTimerRef = useRef<number | null>(null);
  const shareCopyTimerRef = useRef<number | null>(null);

  const docId = document?.id ?? 'default';
  const slides = document?.slidesPreview ?? [];
  const outline = document?.outline ?? [];

  // Reset transient feedback & timers when modal or document changes
  useEffect(() => {
    setOutlineIdx(0);
    setCodeCopied(false);
    setRemoteUrlCopied(false);
    setLinkCopied(false);
    setSharePassword('');
    setShareError(null);
    setShareFeedback(null);

    return () => {
      if (codeCopyTimerRef.current !== null) {
        window.clearTimeout(codeCopyTimerRef.current);
        codeCopyTimerRef.current = null;
      }
      if (remoteUrlCopyTimerRef.current !== null) {
        window.clearTimeout(remoteUrlCopyTimerRef.current);
        remoteUrlCopyTimerRef.current = null;
      }
      if (shareCopyTimerRef.current !== null) {
        window.clearTimeout(shareCopyTimerRef.current);
        shareCopyTimerRef.current = null;
      }
    };
  }, [modal, docId]);

  // Hydrate real share configuration from shareService when Share modal opens
  useEffect(() => {
    if (modal !== 'share' || !document) {
      return;
    }

    let cancelled = false;
    const initialAccess: ShareAccessLevel =
      document.shareAccess ?? (document.isShared ? 'link' : 'private');
    const initialAllowDownload = document.shareAllowDownload ?? true;
    const initialToken = document.shareToken ?? '';

    setShareAccess(initialAccess);
    setSavedShareAccess(initialAccess);
    setAllowDownload(initialAllowDownload);
    setShareToken(initialToken);
    setShareViewCount(0);
    setShareRevokedAt(null);
    setShareError(null);
    setShareFeedback(null);
    setIsLoadingShare(true);

    void getExistingShareForDocument(document.id, document)
      .then((existing) => {
        if (cancelled) {
          return;
        }
        if (existing) {
          const effectiveAccess: ShareAccessLevel = existing.revokedAt
            ? 'private'
            : existing.accessLevel;
          setShareAccess(effectiveAccess);
          setSavedShareAccess(effectiveAccess);
          setAllowDownload(existing.allowDownload);
          setShareToken(existing.shareToken);
          setShareViewCount(existing.viewCount);
          setShareRevokedAt(existing.revokedAt ?? null);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setShareError(
            'Không thể tải trạng thái chia sẻ từ máy chủ. Bạn vẫn có thể cấu hình và lưu mới.'
          );
        }
      })
      .finally(() => {
        if (!cancelled) {
          setIsLoadingShare(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [modal, document]);

  // Global Escape close for non-Presenter modals (PresenterStageView handles its own Escape)
  useEffect(() => {
    if (!modal || (modal === 'present' && document)) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !isSavingShare && !isRevokingShare) {
        onClose();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [modal, document, isSavingShare, isRevokingShare, onClose]);

  const pairingCode = useMemo(() => buildPairingCode(docId), [docId]);

  const remoteUrl = useMemo(() => {
    const origin = resolvePublicWebOrigin();
    const base = `${origin}${window.location.pathname === '/' ? '/' : window.location.pathname}`;
    return `${base}#remote/${encodeURIComponent(pairingCode)}`;
  }, [pairingCode]);

  const publicShareUrl = useMemo(() => {
    if (!shareToken) {
      return '';
    }
    return buildPublicShareUrl(shareToken);
  }, [shareToken]);

  if (!modal) {
    return null;
  }

  // Fullscreen Presenter Stage View (handles Audience + Presenter View, Annotations, Laser, Spotlight, Realtime Remote)
  if (modal === 'present') {
    if (!document) {
      return (
        <div
          role="dialog"
          aria-modal="true"
          className="fixed inset-0 z-50 bg-slate-950/85 backdrop-blur-md pb-safe"
        >
          <div
            onClick={(event) => {
              if (event.target === event.currentTarget) {
                onClose();
              }
            }}
            className="h-full w-full flex items-center justify-center p-4"
          >
            <div className="w-full max-w-md rounded-2xl bg-[#151B28] border border-white/15 p-6 text-center text-slate-200 space-y-4 shadow-2xl">
              <Presentation className="h-8 w-8 text-indigo-300 mx-auto" />
              <p className="text-sm font-medium">
                Chưa chọn tài liệu để mở sân khấu trình chiếu.
              </p>
              <button
                type="button"
                onClick={onClose}
                className="inline-flex items-center justify-center min-h-[40px] px-4 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white transition-colors"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      );
    }

    return (
      <div
        role="dialog"
        aria-modal="true"
        aria-label={`Trình chiếu ${document.title}`}
        className="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md pb-safe"
      >
        <div className="h-full w-full p-1.5 sm:p-3">
          <PresenterStageView
            document={document}
            pairingCode={pairingCode}
            onClose={onClose}
            onOpenQrModal={() => onSwitchModal?.('qr')}
          />
        </div>
      </div>
    );
  }

  const safeOutlineIdx =
    outline.length > 0 ? Math.min(outlineIdx, outline.length - 1) : 0;
  const activeOutline = outline[safeOutlineIdx] ?? null;

  const badgeStyle = document ? TYPE_BADGE_STYLES[document.type] : null;
  const isBusyShare = isSavingShare || isRevokingShare;
  const hasLivePublicLink =
    savedShareAccess !== 'private' &&
    Boolean(shareToken) &&
    !shareRevokedAt;

  const copyTextSafely = (text: string) => {
    if (navigator.clipboard?.writeText) {
      void navigator.clipboard.writeText(text).catch(() => {
        // Ignore clipboard errors in restricted environments
      });
    }
  };

  const handleCopyPairingCode = () => {
    copyTextSafely(pairingCode);
    setCodeCopied(true);
    if (codeCopyTimerRef.current !== null) {
      window.clearTimeout(codeCopyTimerRef.current);
    }
    codeCopyTimerRef.current = window.setTimeout(() => {
      setCodeCopied(false);
      codeCopyTimerRef.current = null;
    }, 1800);
  };

  const handleCopyRemoteUrl = () => {
    copyTextSafely(remoteUrl);
    setRemoteUrlCopied(true);
    if (remoteUrlCopyTimerRef.current !== null) {
      window.clearTimeout(remoteUrlCopyTimerRef.current);
    }
    remoteUrlCopyTimerRef.current = window.setTimeout(() => {
      setRemoteUrlCopied(false);
      remoteUrlCopyTimerRef.current = null;
    }, 1800);
  };

  const handleOpenRemoteTab = () => {
    window.open(remoteUrl, '_blank', 'noopener,noreferrer');
  };

  const triggerLinkCopiedIndicator = () => {
    setLinkCopied(true);
    if (shareCopyTimerRef.current !== null) {
      window.clearTimeout(shareCopyTimerRef.current);
    }
    shareCopyTimerRef.current = window.setTimeout(() => {
      setLinkCopied(false);
      shareCopyTimerRef.current = null;
    }, 2000);
  };

  const handleSaveShare = async (overrideAccess?: ShareAccessLevel) => {
    if (!document || isBusyShare) {
      return;
    }

    const targetAccess = overrideAccess ?? shareAccess;
    setShareError(null);
    setShareFeedback(null);
    setIsSavingShare(true);

    try {
      const saved = await saveDocumentShareSettings({
        document,
        accessLevel: targetAccess,
        password: targetAccess === 'password' ? sharePassword : undefined,
        allowDownload,
      });

      const isSharedNow = saved.accessLevel !== 'private' && !saved.revokedAt;
      setShareAccess(saved.accessLevel);
      setSavedShareAccess(saved.accessLevel);
      setAllowDownload(saved.allowDownload);
      setShareToken(saved.shareToken);
      setShareViewCount(saved.viewCount);
      setShareRevokedAt(saved.revokedAt ?? null);
      setSharePassword('');

      onShareConfigUpdated?.(document.id, {
        shareToken: saved.shareToken,
        shareAccess: saved.accessLevel,
        shareAllowDownload: saved.allowDownload,
        isShared: isSharedNow,
      });

      if (saved.accessLevel === 'private') {
        setShareFeedback('Đã lưu: tài liệu hiện ở chế độ Riêng tư.');
      } else if (saved.accessLevel === 'password') {
        setShareFeedback(
          'Đã bật liên kết chia sẻ có bảo vệ mật khẩu và cập nhật bản xem.'
        );
      } else {
        setShareFeedback(
          'Đã lưu cài đặt chia sẻ. Bất kỳ ai có liên kết đều có thể mở xem.'
        );
      }

      return saved;
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể lưu cài đặt chia sẻ. Vui lòng thử lại.';
      setShareError(message);
      return null;
    } finally {
      setIsSavingShare(false);
    }
  };

  const handleCopyShareLink = async () => {
    if (!document || isBusyShare) {
      return;
    }

    // Avoid copying a dead link when document is still Private or unsaved
    if (!hasLivePublicLink) {
      if (shareAccess === 'password' && !sharePassword.trim()) {
        setShareError(
          'Vui lòng nhập mật khẩu bảo vệ rồi bấm "Lưu cài đặt chia sẻ" trước khi sao chép liên kết.'
        );
        return;
      }
      const nextAccess: ShareAccessLevel =
        shareAccess === 'password' ? 'password' : 'link';
      const saved = await handleSaveShare(nextAccess);
      if (saved && saved.shareToken) {
        const liveUrl = buildPublicShareUrl(saved.shareToken);
        copyTextSafely(liveUrl);
        triggerLinkCopiedIndicator();
        setShareFeedback(
          'Đã bật chia sẻ và sao chép liên kết công khai vào bộ nhớ tạm.'
        );
      }
      return;
    }

    copyTextSafely(publicShareUrl);
    triggerLinkCopiedIndicator();
  };

  const handleRevokeShare = async () => {
    if (!document || !shareToken || isBusyShare) {
      return;
    }

    setShareError(null);
    setShareFeedback(null);
    setIsRevokingShare(true);

    try {
      await revokeDocumentShare(shareToken, document.id);
      const nowIso = new Date().toISOString();
      setShareAccess('private');
      setSavedShareAccess('private');
      setShareRevokedAt(nowIso);
      setSharePassword('');

      onShareConfigUpdated?.(document.id, {
        shareToken,
        shareAccess: 'private',
        shareAllowDownload: allowDownload,
        isShared: false,
      });

      setShareFeedback(
        'Đã thu hồi quyền xem ngay lập tức. Liên kết chia sẻ cũ đã bị vô hiệu hóa.'
      );
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể thu hồi liên kết chia sẻ lúc này.';
      setShareError(message);
    } finally {
      setIsRevokingShare(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-md pb-safe"
    >
      <div
        onClick={(event) => {
          if (event.target === event.currentTarget && !isBusyShare) {
            onClose();
          }
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        {/* QR REMOTE CONTROL MODAL (Real Scannable QR -> #remote/<roomCode>) */}
        {modal === 'qr' && (
          <div className="w-full max-w-md max-h-full rounded-2xl sm:rounded-3xl bg-[#1A202D] border border-white/15 shadow-2xl overflow-hidden text-slate-100 flex flex-col">
            <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 border-b border-white/10 flex items-center justify-between gap-3">
              <div className="flex items-center gap-2.5 min-w-0">
                <span className="h-8 w-8 shrink-0 rounded-xl bg-indigo-500/20 border border-indigo-400/30 flex items-center justify-center text-indigo-300">
                  <QrCode className="h-4 w-4" />
                </span>
                <div className="min-w-0">
                  <h3 className="text-sm font-semibold text-white">
                    Điều khiển trình chiếu qua QR
                  </h3>
                  <p className="text-[11px] text-slate-400 truncate max-w-[220px] sm:max-w-[240px]">
                    {document?.title ?? 'Phiên trình chiếu trực tiếp'}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={onClose}
                aria-label="Đóng hộp thoại mã QR"
                className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[34px] sm:min-w-[34px] rounded-xl bg-white/5 hover:bg-white/15 border border-white/10 flex items-center justify-center text-slate-300 hover:text-white transition-colors"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="flex-1 min-h-0 overflow-y-auto">
              <div className="p-4 sm:p-6 flex flex-col items-center text-center space-y-4">
                {/* Real Scannable QR Code */}
                <div className="p-3.5 sm:p-4 rounded-2xl bg-white shadow-xl ring-4 ring-white/10">
                  <RealQrCode
                    value={remoteUrl}
                    sizeClassName="h-44 w-44 sm:h-48 sm:w-48"
                    darkHex="#0F172A"
                    lightHex="#FFFFFF"
                    ariaLabel={`Mã QR điều khiển từ xa phòng ${pairingCode}`}
                  />
                </div>

                {/* Pairing Room Code */}
                <div className="w-full rounded-2xl bg-[#121722] border border-white/10 px-4 py-3 flex items-center justify-between gap-3">
                  <div className="text-left">
                    <p className="text-[10px] uppercase tracking-wider text-slate-400 font-medium">
                      Mã phòng ghép đôi
                    </p>
                    <p className="text-base font-mono font-bold tracking-widest text-white mt-0.5">
                      {pairingCode}
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={handleCopyPairingCode}
                    className="inline-flex items-center gap-1.5 min-h-[40px] sm:min-h-[34px] px-3 py-1.5 rounded-xl bg-white/10 hover:bg-white/15 text-xs font-medium text-slate-100 transition-colors"
                  >
                    {codeCopied ? (
                      <>
                        <Check className="h-3.5 w-3.5 text-emerald-400" />
                        <span className="text-emerald-300">Đã chép</span>
                      </>
                    ) : (
                      <>
                        <Copy className="h-3.5 w-3.5" />
                        <span>Sao chép</span>
                      </>
                    )}
                  </button>
                </div>

                {/* Real Remote URL */}
                <div className="w-full rounded-2xl bg-[#121722] border border-white/10 px-3.5 py-2.5 flex items-center gap-2">
                  <Link2
                    className="h-3.5 w-3.5 shrink-0 text-indigo-300"
                    aria-hidden="true"
                  />
                  <input
                    type="text"
                    readOnly
                    value={remoteUrl}
                    aria-label="Đường dẫn điều khiển từ xa"
                    onFocus={(event) => event.currentTarget.select()}
                    className="min-w-0 flex-1 bg-transparent text-[11px] font-mono text-slate-300 focus:outline-none truncate"
                  />
                  <button
                    type="button"
                    onClick={handleCopyRemoteUrl}
                    className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white/10 hover:bg-white/15 text-[11px] font-medium text-slate-100 transition-colors"
                  >
                    {remoteUrlCopied ? (
                      <>
                        <Check className="h-3 w-3 text-emerald-400" />
                        <span className="text-emerald-300">Đã chép</span>
                      </>
                    ) : (
                      <span>Chép URL</span>
                    )}
                  </button>
                </div>

                {/* Action Buttons: Open Remote Tab + Launch Presenter Stage */}
                <div className="w-full grid grid-cols-1 gap-2">
                  <button
                    type="button"
                    onClick={handleOpenRemoteTab}
                    className="w-full inline-flex items-center justify-center gap-2 min-h-[44px] px-4 rounded-xl bg-white/10 hover:bg-white/15 border border-white/15 text-xs font-semibold text-white transition-colors"
                  >
                    <ExternalLink className="h-3.5 w-3.5 text-indigo-300" />
                    <span>Mở điều khiển từ xa (Tab mới)</span>
                  </button>

                  {onSwitchModal && document && (
                    <button
                      type="button"
                      onClick={() => onSwitchModal('present')}
                      className="w-full inline-flex items-center justify-center gap-2 min-h-[44px] px-4 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white shadow-lg shadow-indigo-950/40 transition-colors"
                    >
                      <Presentation className="h-4 w-4" />
                      <span>Mở sân khấu để nhận điều khiển</span>
                    </button>
                  )}
                </div>

                <div className="w-full text-left space-y-1.5 text-xs text-slate-300 bg-white/[0.04] border border-white/10 rounded-2xl p-3.5">
                  <div className="flex items-center gap-2 font-medium text-white">
                    <Smartphone className="h-3.5 w-3.5 text-indigo-300 shrink-0" />
                    <span>Chuyển slide, tia Laser &amp; Spotlight thời gian thực</span>
                  </div>
                  <p className="text-[11.5px] text-slate-400 leading-relaxed">
                    Quét mã QR bằng điện thoại hoặc mở tab mới trên cùng máy để
                    điều khiển sân khấu qua Supabase Realtime / kết nối cục bộ.
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* SHARE MODAL (Connected to shareService.ts) */}
        {modal === 'share' && (
          <div className="w-full max-w-lg max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
            <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 border-b border-[#EAE5DC] flex items-center justify-between gap-3">
              <div className="flex items-center gap-3 min-w-0">
                <span className="h-9 w-9 shrink-0 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                  <Share2 className="h-4 w-4" />
                </span>
                <div className="min-w-0">
                  <div className="flex items-center gap-2">
                    <h3 className="text-sm font-semibold text-slate-900">
                      Chia sẻ tài liệu
                    </h3>
                    {badgeStyle && (
                      <span
                        className={`inline-flex items-center rounded-md border px-1.5 py-0.5 text-[10px] font-semibold tracking-wide ${badgeStyle.className}`}
                      >
                        {badgeStyle.label}
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-slate-500 truncate mt-0.5">
                    {document?.title ?? 'Tài liệu hiện tại'}
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={onClose}
                disabled={isBusyShare}
                aria-label="Đóng hộp thoại chia sẻ"
                className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors disabled:opacity-50"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="flex-1 min-h-0 overflow-y-auto">
              <div className="p-4 sm:p-6 space-y-5">
                {isLoadingShare && (
                  <div className="flex items-center gap-2 rounded-xl border border-indigo-100 bg-indigo-50/70 px-3.5 py-2.5 text-xs text-indigo-700">
                    <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
                    <span>Đang tải cấu hình chia sẻ hiện tại...</span>
                  </div>
                )}

                {shareError && (
                  <div
                    role="alert"
                    className="flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3.5 py-2.5 text-xs text-rose-700"
                  >
                    <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
                    <span>{shareError}</span>
                  </div>
                )}

                {shareFeedback && (
                  <div
                    role="status"
                    className="flex items-start gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-3.5 py-2.5 text-xs text-emerald-700"
                  >
                    <CheckCircle2 className="h-4 w-4 shrink-0 mt-0.5" />
                    <span>{shareFeedback}</span>
                  </div>
                )}

                {/* Access Permission Selector */}
                <div className="space-y-2">
                  <label className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                    Quyền truy cập
                  </label>
                  <div className="grid grid-cols-1 gap-2">
                    {[
                      {
                        key: 'private' as const,
                        title: 'Riêng tư',
                        desc: 'Chỉ mình bạn có thể mở và chỉnh sửa tài liệu này',
                        icon: Lock,
                      },
                      {
                        key: 'link' as const,
                        title: 'Bất kỳ ai có liên kết',
                        desc: 'Mở xem trực tiếp trên trình duyệt không cần đăng nhập',
                        icon: Globe,
                      },
                      {
                        key: 'password' as const,
                        title: 'Bảo vệ bằng mật khẩu',
                        desc: 'Yêu cầu nhập mã bảo vệ trước khi hiển thị nội dung',
                        icon: KeyRound,
                      },
                    ].map((option) => {
                      const OptionIcon = option.icon;
                      const isSelected = shareAccess === option.key;
                      return (
                        <button
                          key={option.key}
                          type="button"
                          disabled={isBusyShare}
                          onClick={() => {
                            setShareAccess(option.key);
                            setShareError(null);
                          }}
                          aria-pressed={isSelected}
                          className={`w-full text-left rounded-2xl border px-3.5 py-2.5 flex items-center gap-3 transition-all min-h-[54px] disabled:opacity-60 ${
                            isSelected
                              ? 'border-indigo-500 bg-indigo-50/70 ring-2 ring-indigo-500/15'
                              : 'border-[#E6E0D6] bg-[#FAF8F5]/70 hover:bg-white hover:border-slate-300'
                          }`}
                        >
                          <span
                            className={`h-8 w-8 rounded-xl shrink-0 flex items-center justify-center ${
                              isSelected
                                ? 'bg-indigo-600 text-white'
                                : 'bg-white border border-[#E6E0D6] text-slate-500'
                            }`}
                          >
                            <OptionIcon className="h-4 w-4" />
                          </span>
                          <div className="min-w-0 flex-1">
                            <p className="text-xs font-semibold text-slate-900">
                              {option.title}
                            </p>
                            <p className="text-[11px] text-slate-500 leading-snug mt-0.5">
                              {option.desc}
                            </p>
                          </div>
                          <span
                            className={`h-4 w-4 rounded-full border shrink-0 flex items-center justify-center ${
                              isSelected
                                ? 'border-indigo-600 bg-indigo-600'
                                : 'border-slate-300 bg-white'
                            }`}
                          >
                            {isSelected && (
                              <span className="h-1.5 w-1.5 rounded-full bg-white" />
                            )}
                          </span>
                        </button>
                      );
                    })}
                  </div>

                  {shareAccess === 'password' && (
                    <div className="pt-1">
                      <input
                        type="password"
                        value={sharePassword}
                        disabled={isBusyShare}
                        onChange={(event) => {
                          setSharePassword(event.target.value);
                          setShareError(null);
                        }}
                        placeholder={
                          savedShareAccess === 'password'
                            ? 'Nhập mật khẩu mới (để trống nếu giữ mật khẩu cũ)...'
                            : 'Đặt mật khẩu truy cập (ví dụ: LITE2026)...'
                        }
                        aria-label="Mật khẩu chia sẻ"
                        className="w-full h-10 rounded-xl border border-[#DDD6CA] bg-white px-3.5 text-xs text-slate-800 placeholder:text-slate-400 focus:border-indigo-500 focus:outline-none focus:ring-2 focus:ring-indigo-500/15"
                      />
                    </div>
                  )}
                </div>

                {/* Allow Original File Download Toggle */}
                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] px-4 py-3 flex items-center justify-between gap-3">
                  <div className="flex items-start gap-2.5 min-w-0">
                    <Download
                      className="h-4 w-4 text-slate-500 mt-0.5 shrink-0"
                      aria-hidden="true"
                    />
                    <div className="min-w-0">
                      <p className="text-xs font-semibold text-slate-800">
                        Cho phép tải tệp gốc
                      </p>
                      <p className="text-[11px] text-slate-500 leading-snug mt-0.5">
                        {allowDownload
                          ? `Người xem có thể tải tệp .${document?.type ?? 'pptx'} gốc về máy.`
                          : 'Chế độ chỉ xem trực tuyến, ẩn nút tải tệp nguồn.'}
                      </p>
                    </div>
                  </div>
                  <button
                    type="button"
                    role="switch"
                    disabled={isBusyShare}
                    aria-checked={allowDownload}
                    aria-label="Cho phép tải tệp gốc"
                    onClick={() => setAllowDownload((prev) => !prev)}
                    className={`relative inline-flex h-6 w-11 shrink-0 items-center rounded-full transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40 disabled:opacity-60 ${
                      allowDownload ? 'bg-indigo-600' : 'bg-slate-300'
                    }`}
                  >
                    <span
                      className={`inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform ${
                        allowDownload ? 'translate-x-6' : 'translate-x-1'
                      }`}
                    />
                  </button>
                </div>

                {/* Save Share Settings Primary CTA */}
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    disabled={isBusyShare || !document}
                    onClick={() => {
                      void handleSaveShare();
                    }}
                    className="flex-1 inline-flex items-center justify-center gap-2 min-h-[44px] px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-xs font-semibold text-white shadow-sm shadow-indigo-600/20 transition-all disabled:opacity-60"
                  >
                    {isSavingShare ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin" />
                        <span>Đang lưu cài đặt...</span>
                      </>
                    ) : (
                      <>
                        <Save className="h-4 w-4" />
                        <span>Lưu cài đặt chia sẻ</span>
                      </>
                    )}
                  </button>

                  {hasLivePublicLink && (
                    <button
                      type="button"
                      disabled={isBusyShare}
                      onClick={() => {
                        void handleRevokeShare();
                      }}
                      className="inline-flex items-center justify-center gap-1.5 min-h-[44px] px-3.5 rounded-xl border border-rose-200 bg-rose-50 hover:bg-rose-100 text-xs font-semibold text-rose-700 transition-colors disabled:opacity-60 whitespace-nowrap"
                    >
                      {isRevokingShare ? (
                        <Loader2 className="h-3.5 w-3.5 animate-spin" />
                      ) : (
                        <ShieldOff className="h-3.5 w-3.5" />
                      )}
                      <span>Thu hồi quyền xem</span>
                    </button>
                  )}
                </div>

                {/* Active Public Link + Real QR + Views OR Private Guard */}
                {hasLivePublicLink ? (
                  <div className="space-y-4 pt-2 border-t border-[#EAE5DC]">
                    <div className="flex items-center justify-between gap-2">
                      <label
                        htmlFor="share-doc-link"
                        className="text-xs font-semibold uppercase tracking-wider text-slate-500"
                      >
                        Liên kết công khai đang hoạt động
                      </label>
                      <span className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-2.5 py-0.5 text-[11px] font-medium text-slate-600 tabular-nums">
                        <Eye className="h-3 w-3 text-indigo-500" />
                        <span>{shareViewCount} lượt xem</span>
                      </span>
                    </div>

                    <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                      <div className="flex-1 min-w-0 flex items-center gap-2 rounded-xl border border-[#E2DCD2] bg-[#FAF8F5] px-3 h-11">
                        <Link2
                          className="h-4 w-4 shrink-0 text-slate-400"
                          aria-hidden="true"
                        />
                        <input
                          id="share-doc-link"
                          type="text"
                          readOnly
                          value={publicShareUrl}
                          onFocus={(event) => event.currentTarget.select()}
                          className="w-full bg-transparent text-xs font-mono text-slate-700 focus:outline-none truncate"
                        />
                      </div>
                      <button
                        type="button"
                        disabled={isBusyShare}
                        onClick={() => {
                          void handleCopyShareLink();
                        }}
                        className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[44px] px-4 rounded-xl text-xs font-semibold transition-all whitespace-nowrap ${
                          linkCopied
                            ? 'bg-emerald-600 text-white shadow-sm shadow-emerald-600/20'
                            : 'bg-slate-900 hover:bg-slate-800 text-white shadow-sm'
                        }`}
                      >
                        {linkCopied ? (
                          <>
                            <Check className="h-4 w-4" />
                            <span>Đã sao chép</span>
                          </>
                        ) : (
                          <>
                            <Copy className="h-3.5 w-3.5" />
                            <span>Sao chép liên kết</span>
                          </>
                        )}
                      </button>
                    </div>

                    {/* Real QR + Test Link Card */}
                    <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5]/70 p-3.5 flex items-center gap-4">
                      <div className="p-2 rounded-xl bg-white border border-[#E6E0D6] shrink-0 shadow-sm">
                        <RealQrCode
                          value={publicShareUrl}
                          sizeClassName="h-24 w-24"
                          darkHex="#0F172A"
                          lightHex="#FFFFFF"
                          ariaLabel={`Mã QR chia sẻ ${document?.title ?? 'tài liệu'}`}
                        />
                      </div>
                      <div className="min-w-0 flex-1 space-y-2">
                        <div className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-900">
                          <Smartphone className="h-3.5 w-3.5 text-indigo-600 shrink-0" />
                          <span>Mở nhanh bằng QR hoặc Tab mới</span>
                        </div>
                        <p className="text-[11.5px] text-slate-500 leading-relaxed">
                          Người nhận quét mã QR hoặc mở đường dẫn để xem ngay mà
                          không cần tài khoản LITE.
                        </p>
                        <a
                          href={publicShareUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-600 hover:text-indigo-700"
                        >
                          <ExternalLink className="h-3.5 w-3.5" />
                          <span>Mở thử link chia sẻ</span>
                        </a>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="rounded-2xl border border-dashed border-[#DDD6CA] bg-[#FAF8F5] p-4 space-y-3">
                    <p className="text-xs text-slate-600 leading-relaxed">
                      Tài liệu hiện đang ở chế độ <strong>Riêng tư</strong> nên
                      chưa có liên kết công khai. Chọn quyền chia sẻ và bấm{' '}
                      <strong>Lưu cài đặt chia sẻ</strong>, hoặc bật nhanh bên
                      dưới để tránh sao chép liên kết không truy cập được.
                    </p>
                    <button
                      type="button"
                      disabled={isBusyShare || !document}
                      onClick={() => {
                        void handleCopyShareLink();
                      }}
                      className="inline-flex items-center justify-center gap-1.5 min-h-[40px] px-3.5 rounded-xl bg-white hover:bg-indigo-50 border border-[#DFD8CC] hover:border-indigo-300 text-xs font-semibold text-indigo-700 transition-colors disabled:opacity-60"
                    >
                      <Copy className="h-3.5 w-3.5" />
                      <span>
                        {shareAccess === 'password'
                          ? 'Lưu mật khẩu & Sao chép liên kết'
                          : 'Bật liên kết công khai & Sao chép'}
                      </span>
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* FULLSCREEN READER / INSPECTOR MODAL */}
        {modal === 'fullscreen' && (
          <div className="w-full max-w-5xl h-full max-h-full sm:h-[86dvh] rounded-2xl sm:rounded-3xl bg-[#FAF8F5] text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
            <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-white border-b border-[#EAE5DC] flex items-center justify-between gap-3">
              <div className="flex items-center gap-3 min-w-0">
                <span className="h-9 w-9 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600 shrink-0">
                  <Maximize2 className="h-4 w-4" />
                </span>
                <div className="min-w-0">
                  <h3 className="text-sm font-semibold text-slate-900 truncate">
                    {document?.title ?? 'Xem mở rộng tài liệu'}
                  </h3>
                  <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-[11px] sm:text-xs text-slate-500 mt-0.5">
                    {document && (
                      <>
                        {badgeStyle && (
                          <span
                            className={`inline-flex items-center rounded border px-1.5 py-0 text-[10px] font-semibold ${badgeStyle.className}`}
                          >
                            {badgeStyle.label}
                          </span>
                        )}
                        <span>•</span>
                        <span>{document.updatedAt}</span>
                        {document.pageCount !== undefined &&
                          document.pageCount > 0 && (
                            <>
                              <span>•</span>
                              <span>
                                {document.pageCount}{' '}
                                {getDocumentUnit(document.type)}
                              </span>
                            </>
                          )}
                        {document.fileSize && (
                          <>
                            <span>•</span>
                            <span>{document.fileSize}</span>
                          </>
                        )}
                      </>
                    )}
                  </div>
                </div>
              </div>

              <div className="shrink-0 flex items-center gap-2">
                {onSwitchModal && document && (
                  <button
                    type="button"
                    onClick={() => onSwitchModal('present')}
                    className="inline-flex items-center gap-1.5 min-h-[38px] px-3 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-xs font-semibold text-white transition-colors"
                  >
                    <Presentation className="h-3.5 w-3.5" />
                    <span className="hidden sm:inline">Trình chiếu</span>
                  </button>
                )}
                <button
                  type="button"
                  onClick={onClose}
                  aria-label="Đóng chế độ mở rộng"
                  className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors"
                >
                  <X className="h-4 w-4" />
                </button>
              </div>
            </div>

            {document && activeOutline ? (
              <div className="flex-1 min-h-0 flex flex-col md:grid md:grid-cols-12">
                {/* Mobile Horizontal Scrollable Outline Strip (< md) */}
                <div className="md:hidden shrink-0 border-b border-[#EAE5DC] bg-[#F5F1EA] px-3 py-2.5 flex items-center gap-2 overflow-x-auto no-scrollbar">
                  {outline.map((item, idx) => {
                    const isActive = idx === safeOutlineIdx;
                    return (
                      <button
                        key={`${item.id}-fs-mobile`}
                        type="button"
                        onClick={() => setOutlineIdx(idx)}
                        className={`shrink-0 min-h-[38px] px-3 py-1.5 rounded-xl text-xs whitespace-nowrap transition-all ${
                          isActive
                            ? 'bg-white text-indigo-700 font-semibold shadow-sm ring-1 ring-indigo-500/25'
                            : 'bg-white/60 text-slate-700 hover:bg-white'
                        }`}
                      >
                        <span className="font-semibold mr-1">
                          {item.romanIndex}.
                        </span>
                        {item.title}
                      </button>
                    );
                  })}
                </div>

                {/* Desktop / Tablet Outline Sidebar (md+) */}
                <aside className="hidden md:flex md:col-span-4 border-r border-[#EAE5DC] bg-[#F5F1EA] p-5 overflow-y-auto flex-col justify-between gap-6">
                  <div>
                    <div className="flex items-center gap-2 text-xs font-semibold text-slate-700 mb-3">
                      <BookOpen className="h-4 w-4 text-indigo-600" />
                      <span>Mục lục thông minh</span>
                    </div>
                    <div className="space-y-1.5">
                      {outline.map((item, idx) => {
                        const isActive = idx === safeOutlineIdx;
                        return (
                          <button
                            key={item.id}
                            type="button"
                            onClick={() => setOutlineIdx(idx)}
                            className={`w-full text-left px-3.5 py-2.5 rounded-xl transition-all flex items-start justify-between gap-2 ${
                              isActive
                                ? 'bg-white text-indigo-700 font-semibold shadow-sm ring-1 ring-indigo-500/25'
                                : 'text-slate-700 hover:bg-white/60'
                            }`}
                          >
                            <span className="text-xs leading-snug">
                              <span className="font-semibold mr-1.5">
                                {item.romanIndex}.
                              </span>
                              {item.title}
                            </span>
                            <span className="text-[11px] text-slate-400 shrink-0 tabular-nums">
                              Tr. {item.pageNumber}
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  <div className="rounded-2xl bg-white/80 border border-[#E7E1D6] p-3.5 space-y-2 text-xs text-slate-600">
                    <div className="flex items-center gap-2 text-slate-800 font-medium">
                      <FolderOpen className="h-3.5 w-3.5 text-indigo-500" />
                      <span>{document.folder ?? 'Thư mục chung'}</span>
                    </div>
                    <div className="flex items-center gap-2 text-slate-500">
                      <Layers className="h-3.5 w-3.5" />
                      <span>
                        {document.type === 'xlsx' && document.sheets
                          ? `${document.sheets.length} sheet bảng tính · ${outline.length} chương mục`
                          : `${slides.length} slide xem trước · ${outline.length} chương mục`}
                      </span>
                    </div>
                  </div>
                </aside>

                {/* Wide Document Reading Sheet */}
                <div className="flex-1 min-h-0 md:col-span-8 overflow-y-auto">
                  <div className="p-4 sm:p-8">
                    <div className="max-w-2xl mx-auto rounded-2xl bg-white border border-[#EAE4D9] shadow-sm p-5 sm:p-9 space-y-5 sm:space-y-6">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3.5 sm:pb-4">
                        <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-indigo-50 text-indigo-700 text-xs font-semibold">
                          <FileText className="h-3.5 w-3.5" />
                          <span>
                            Mục {activeOutline.romanIndex} · Trang{' '}
                            {activeOutline.pageNumber}
                          </span>
                        </span>
                        <span className="text-xs text-slate-400">
                          Smart Reader View
                        </span>
                      </div>

                      <div>
                        <h2 className="text-lg sm:text-2xl font-bold text-slate-900 leading-snug">
                          {activeOutline.romanIndex}. {activeOutline.title}
                        </h2>
                        <p className="mt-2.5 sm:mt-3 text-sm sm:text-[15px] text-slate-700 leading-relaxed">
                          {activeOutline.summary}
                        </p>
                      </div>

                      <div className="space-y-2.5">
                        <h4 className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                          Ý chính nổi bật
                        </h4>
                        <ul className="space-y-2">
                          {activeOutline.keyPoints.map((point, idx) => (
                            <li
                              key={`${activeOutline.id}-fs-kp-${idx}`}
                              className="flex items-start gap-2.5 text-sm text-slate-800 leading-relaxed"
                            >
                              <span className="mt-2 h-1.5 w-1.5 rounded-full bg-indigo-600 shrink-0" />
                              <span>{point}</span>
                            </li>
                          ))}
                        </ul>
                      </div>

                      <div className="rounded-xl bg-[#FAF8F5] border border-[#EBE5DA] p-4 space-y-1.5">
                        <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
                          Trích đoạn tài liệu
                        </span>
                        <p className="text-sm text-slate-700 leading-relaxed">
                          {document.contentPreview}
                        </p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <div className="flex-1 flex items-center justify-center text-sm text-slate-500">
                Chọn một tài liệu để mở rộng toàn màn hình.
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
