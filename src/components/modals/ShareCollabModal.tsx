import { useEffect, useState } from "react";
import {
  Check,
  Edit3,
  Eye,
  FileDown,
  FileOutput,
  Globe,
  Link2,
  Loader2,
  Lock,
  MessageCircle,
  Send,
  Share2,
  Users,
  X,
} from "lucide-react";
import type { CloudUserSummary, DocumentItem } from "../../types/document";
import {
  buildCollabShareUrl,
  createOrUpdateShareLink,
  getShareLinkInfo,
  revokeShareLink,
  type CollabPermission,
  type CollaboratorPresence,
} from "../../services/realtimeCollab";
import { loadOriginalFileBinary } from "../../services/offlineSyncQueue";
import {
  shareFileUniversal,
  shareLinkUniversal,
} from "../../utils/saveFileUniversal";

const ORIGINAL_EXT_LABEL: Record<DocumentItem["type"], string> = {
  docx: ".docx",
  xlsx: ".xlsx",
  pptx: ".pptx",
  pdf: ".pdf",
};

export interface ShareCollabModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: DocumentItem | null;
  currentUser?: CloudUserSummary | null;
  collaborators?: CollaboratorPresence[];
  selfSessionId?: string;
  onShareStateChanged?: (info: {
    documentId: string;
    permission: CollabPermission;
    shareToken?: string;
    isShared: boolean;
  }) => void;
}

export function ShareCollabModal({
  isOpen,
  onClose,
  document,
  currentUser,
  collaborators = [],
  selfSessionId,
  onShareStateChanged,
}: ShareCollabModalProps) {
  const [permission, setPermission] = useState<CollabPermission>("private");
  const [shareUrl, setShareUrl] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);
  const [shareFeedback, setShareFeedback] = useState<string | null>(null);
  const [sendingKind, setSendingKind] = useState<"original" | "pdf" | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const ownerId =
    currentUser?.id || document?.sharedOwnerId || document?.userId || "local-owner";

  useEffect(() => {
    if (!isOpen || !document) {
      return;
    }
    let cancelled = false;
    setErrorMsg(null);
    setCopied(false);
    setShareFeedback(null);
    setSendingKind(null);

    if (document.shareToken && document.shareAccess === "link") {
      const initialPerm: "view" | "edit" =
        document.collabPermission === "view" ? "view" : "edit";
      setPermission(initialPerm);
      setShareUrl(buildCollabShareUrl(document.shareToken, initialPerm));
    } else {
      setPermission("private");
      setShareUrl("");
    }

    setIsLoading(true);
    void getShareLinkInfo(ownerId, document.id)
      .then((info) => {
        if (cancelled) return;
        setPermission(info.permission);
        setShareUrl(
          info.shareUrl ||
            (info.shareToken && info.permission !== "private"
              ? buildCollabShareUrl(info.shareToken, info.permission)
              : "")
        );
      })
      .catch(() => {
        // Giữ trạng thái hiện tại nếu lỗi mạng
      })
      .finally(() => {
        if (!cancelled) setIsLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [isOpen, document, ownerId]);

  useEffect(() => {
    if (!isOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && !sendingKind) onClose();
    };
    const onHardwareBack = (e: Event) => {
      if (!sendingKind) {
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener("keydown", onKey);
    window.addEventListener("lite:hardware-back", onHardwareBack);
    return () => {
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("lite:hardware-back", onHardwareBack);
    };
  }, [isOpen, sendingKind, onClose]);

  if (!isOpen || !document) {
    return null;
  }

  const handleSelectPermission = async (nextPerm: CollabPermission) => {
    if (isSaving || nextPerm === permission) {
      return;
    }
    setIsSaving(true);
    setErrorMsg(null);
    try {
      if (nextPerm === "private") {
        await revokeShareLink(ownerId, document.id);
        setPermission("private");
        setShareUrl("");
        onShareStateChanged?.({
          documentId: document.id,
          permission: "private",
          isShared: false,
        });
      } else {
        const res = await createOrUpdateShareLink(
          ownerId,
          document.id,
          nextPerm,
          document
        );
        setPermission(res.permission);
        setShareUrl(res.shareUrl);
        onShareStateChanged?.({
          documentId: document.id,
          permission: res.permission,
          shareToken: res.shareToken,
          isShared: true,
        });
      }
    } catch (err: unknown) {
      setErrorMsg(
        err instanceof Error
          ? err.message
          : "Không thể cập nhật quyền chia sẻ. Vui lòng thử lại."
      );
    } finally {
      setIsSaving(false);
    }
  };

  const copyTextToClipboard = async (value: string) => {
    try {
      await navigator.clipboard.writeText(value);
    } catch {
      const input = window.document.createElement("input");
      input.value = value;
      window.document.body.appendChild(input);
      input.select();
      window.document.execCommand("copy");
      window.document.body.removeChild(input);
    }
  };

  const handleCopyLink = async () => {
    if (!shareUrl) return;
    await copyTextToClipboard(shareUrl);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 2200);
  };

  const ensurePublicShareUrl = async (): Promise<string> => {
    if (permission !== "private" && shareUrl) {
      return shareUrl;
    }
    setIsSaving(true);
    setErrorMsg(null);
    try {
      const res = await createOrUpdateShareLink(
        ownerId,
        document.id,
        "view",
        document
      );
      setPermission(res.permission);
      setShareUrl(res.shareUrl);
      onShareStateChanged?.({
        documentId: document.id,
        permission: res.permission,
        shareToken: res.shareToken,
        isShared: true,
      });
      return res.shareUrl;
    } finally {
      setIsSaving(false);
    }
  };

  const handleQuickAppLinkShare = async (
    channel: "zalo" | "messenger" | "facebook" | "telegram" | "system"
  ) => {
    if (isSaving) return;
    setErrorMsg(null);
    setShareFeedback(null);
    try {
      const url = await ensurePublicShareUrl();
      const shareText = `Mở tài liệu "${document.title}" trên LITE Workspace`;

      if (channel === "system") {
        const res = await shareLinkUniversal({
          title: document.title,
          text: shareText,
          url,
        });
        if (res === "copied") {
          setCopied(true);
          setShareFeedback("Đã sao chép liên kết công khai vào khay nhớ tạm.");
          window.setTimeout(() => setCopied(false), 2200);
        }
        return;
      }

      if (channel === "facebook") {
        await copyTextToClipboard(url);
        window.open(
          `https://www.facebook.com/sharer/sharer.php?u=${encodeURIComponent(url)}`,
          "_blank",
          "noopener,noreferrer"
        );
        setShareFeedback("Đã mở Facebook Share (liên kết cũng đã được sao chép).");
        return;
      }

      if (channel === "telegram") {
        window.open(
          `https://t.me/share/url?url=${encodeURIComponent(url)}&text=${encodeURIComponent(shareText)}`,
          "_blank",
          "noopener,noreferrer"
        );
        setShareFeedback("Đã mở chia sẻ qua Telegram.");
        return;
      }

      // Zalo or Messenger: prefer native system share sheet on mobile/Android, else copy + open web app
      const res = await shareLinkUniversal({
        title: document.title,
        text: shareText,
        url,
      });
      if (res === "copied") {
        setCopied(true);
        window.setTimeout(() => setCopied(false), 2200);
        const appUrl =
          channel === "zalo" ? "https://chat.zalo.me/" : "https://www.messenger.com/";
        const appName = channel === "zalo" ? "Zalo" : "Messenger";
        window.open(appUrl, "_blank", "noopener,noreferrer");
        setShareFeedback(
          `Đã sao chép liên kết chuẩn — hãy dán (Ctrl+V) vào khung chat ${appName}.`
        );
      }
    } catch (err: unknown) {
      setErrorMsg(
        err instanceof Error
          ? err.message
          : "Không thể chia sẻ liên kết. Vui lòng thử lại."
      );
    }
  };

  const handleSendAsOriginalFile = async () => {
    if (sendingKind) return;
    setSendingKind("original");
    setErrorMsg(null);
    setShareFeedback(null);
    try {
      let docForExport = document;
      if (!docForExport.hasLocalEdits && !docForExport.rawArrayBuffer) {
        const cached = await loadOriginalFileBinary(
          currentUser?.id ?? null,
          docForExport.id
        ).catch(() => null);
        if (cached && cached.buffer.byteLength > 0) {
          docForExport = {
            ...docForExport,
            rawArrayBuffer: cached.buffer,
            originalMimeType: docForExport.originalMimeType ?? cached.mimeType,
            originalFileName: docForExport.originalFileName ?? cached.fileName,
          };
        }
      }

      const { buildDocumentUploadPayload } = await import(
        "../../utils/fileExporter"
      );
      const { blob, fileName, mimeType } =
        await buildDocumentUploadPayload(docForExport);
      const res = await shareFileUniversal(blob, fileName, {
        mimeType,
        title: docForExport.title,
        text: `Gửi tệp "${fileName}" từ LITE Workspace`,
      });
      setShareFeedback(res.message);
    } catch (err: unknown) {
      setErrorMsg(
        err instanceof Error
          ? err.message
          : "Không thể xuất và gửi tệp gốc. Vui lòng thử lại."
      );
    } finally {
      setSendingKind(null);
    }
  };

  const handleSendAsPdfFile = async () => {
    if (sendingKind) return;
    setSendingKind("pdf");
    setErrorMsg(null);
    setShareFeedback(null);
    try {
      const { exportDocumentToPdf } = await import("../../utils/quickConverter");
      const { blob, fileName } = await exportDocumentToPdf(document);
      const res = await shareFileUniversal(blob, fileName, {
        mimeType: "application/pdf",
        title: `${document.title} (PDF)`,
        text: `Gửi bản PDF "${fileName}" từ LITE Workspace`,
      });
      setShareFeedback(res.message);
    } catch (err: unknown) {
      setErrorMsg(
        err instanceof Error
          ? err.message
          : "Không thể tạo và gửi bản PDF. Vui lòng thử lại."
      );
    } finally {
      setSendingKind(null);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/55 backdrop-blur-xs p-4"
      onClick={onClose}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={"Chia sẻ " + document.title}
        className="w-full max-w-md max-h-[92dvh] flex flex-col rounded-2xl bg-white border border-slate-200 shadow-2xl overflow-hidden text-slate-900"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="shrink-0 flex items-center justify-between px-5 py-4 border-b border-slate-100">
          <div className="min-w-0 pr-3">
            <h3 className="text-base font-semibold text-slate-900 truncate">
              Chia sẻ "{document.title}"
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Cộng tác chỉnh sửa thời gian thực trên LITE
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng hộp thoại chia sẻ"
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="flex-1 min-h-0 overflow-y-auto p-5 space-y-4">
          {collaborators.length > 0 && (
            <div className="space-y-2 rounded-xl bg-slate-50 border border-slate-200/80 p-3">
              <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-700">
                <Users className="w-3.5 h-3.5 text-indigo-600" />
                <span>Đang trực tuyến trong tệp ({collaborators.length})</span>
              </div>
              <div className="flex flex-wrap gap-1.5">
                {collaborators.map((person) => {
                  const isSelf =
                    Boolean(selfSessionId) &&
                    person.sessionId === selfSessionId;
                  return (
                    <span
                      key={person.sessionId}
                      className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-white border border-slate-200 text-slate-700"
                    >
                      <span
                        className="w-2 h-2 rounded-full shrink-0"
                        style={{ backgroundColor: person.color || "#2563EB" }}
                      />
                      <span className="truncate max-w-[140px]">
                        {person.displayName}
                        {isSelf ? " (Bạn)" : ""}
                      </span>
                      {person.cellCoord && (
                        <span className="text-[10px] text-slate-400">
                          • {person.cellCoord}
                        </span>
                      )}
                    </span>
                  );
                })}
              </div>
            </div>
          )}

          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <label className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                Quyền truy cập liên kết
              </label>
              {(isLoading || isSaving) && (
                <span className="inline-flex items-center gap-1 text-[11px] text-indigo-600 font-medium">
                  <Loader2 className="w-3 h-3 animate-spin" />
                  Đang đồng bộ...
                </span>
              )}
            </div>

            <div className="space-y-2">
              <button
                type="button"
                disabled={isSaving}
                onClick={() => void handleSelectPermission("private")}
                className={
                  "w-full flex items-center gap-3 p-3 rounded-xl border text-left transition " +
                  (permission === "private"
                    ? "border-indigo-600 bg-indigo-50/70 text-indigo-900"
                    : "border-slate-200 hover:bg-slate-50 text-slate-700")
                }
              >
                <span
                  className={
                    "w-8 h-8 rounded-lg flex items-center justify-center shrink-0 " +
                    (permission === "private"
                      ? "bg-indigo-600 text-white"
                      : "bg-slate-100 text-slate-500")
                  }
                >
                  <Lock className="w-4 h-4" />
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-xs font-semibold">
                    Hạn chế (Chỉ mình tôi)
                  </p>
                  <p className="text-[11px] text-slate-500 truncate">
                    Thu hồi liên kết, chỉ chủ sở hữu mở được tệp
                  </p>
                </div>
              </button>

              <button
                type="button"
                disabled={isSaving}
                onClick={() => void handleSelectPermission("view")}
                className={
                  "w-full flex items-center gap-3 p-3 rounded-xl border text-left transition " +
                  (permission === "view"
                    ? "border-indigo-600 bg-indigo-50/70 text-indigo-900"
                    : "border-slate-200 hover:bg-slate-50 text-slate-700")
                }
              >
                <span
                  className={
                    "w-8 h-8 rounded-lg flex items-center justify-center shrink-0 " +
                    (permission === "view"
                      ? "bg-indigo-600 text-white"
                      : "bg-slate-100 text-slate-500")
                  }
                >
                  <Eye className="w-4 h-4" />
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-xs font-semibold">
                    Bất kỳ ai có liên kết — Chỉ xem
                  </p>
                  <p className="text-[11px] text-slate-500 truncate">
                    Người nhận link chỉ được xem trực tiếp, không thể sửa
                  </p>
                </div>
              </button>

              <button
                type="button"
                disabled={isSaving}
                onClick={() => void handleSelectPermission("edit")}
                className={
                  "w-full flex items-center gap-3 p-3 rounded-xl border text-left transition " +
                  (permission === "edit"
                    ? "border-indigo-600 bg-indigo-50/70 text-indigo-900"
                    : "border-slate-200 hover:bg-slate-50 text-slate-700")
                }
              >
                <span
                  className={
                    "w-8 h-8 rounded-lg flex items-center justify-center shrink-0 " +
                    (permission === "edit"
                      ? "bg-indigo-600 text-white"
                      : "bg-slate-100 text-slate-500")
                  }
                >
                  <Edit3 className="w-4 h-4" />
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-xs font-semibold">
                    Bất kỳ ai có liên kết — Được phép chỉnh sửa
                  </p>
                  <p className="text-[11px] text-slate-500 truncate">
                    Nhiều người cùng gõ sửa trực tiếp theo thời gian thực
                  </p>
                </div>
              </button>
            </div>
          </div>

          {permission !== "private" && shareUrl && (
            <div className="pt-1 space-y-2.5">
              <div className="flex items-center gap-1.5 text-[11px] font-medium text-emerald-700">
                <Globe className="w-3.5 h-3.5" />
                <span>Liên kết công khai chuẩn (mở trên mọi trình duyệt)</span>
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  readOnly
                  value={shareUrl}
                  onFocus={(e) => e.currentTarget.select()}
                  className="flex-1 text-xs px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-slate-700 outline-none"
                />
                <button
                  type="button"
                  onClick={() => void handleCopyLink()}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shrink-0 transition"
                >
                  {copied ? (
                    <>
                      <Check className="w-3.5 h-3.5" />
                      <span>Đã sao chép</span>
                    </>
                  ) : (
                    <>
                      <Link2 className="w-3.5 h-3.5" />
                      <span>Sao chép liên kết</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          )}

          {/* Chia sẻ nhanh qua ứng dụng & Gửi dưới dạng tệp */}
          <div className="pt-2 border-t border-slate-100 space-y-3">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                Gửi trực tiếp sang ứng dụng khác
              </p>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Chia sẻ liên kết chuẩn hoặc gửi tệp đính kèm qua Zalo, Messenger, Facebook...
              </p>
            </div>

            <div className="grid grid-cols-5 gap-1.5">
              <button
                type="button"
                disabled={isSaving || sendingKind !== null}
                onClick={() => void handleQuickAppLinkShare("zalo")}
                className="flex flex-col items-center justify-center gap-1 rounded-xl border border-sky-200 bg-sky-50/70 py-2 px-1 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 transition disabled:opacity-50"
              >
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-[#0068FF] text-[10px] font-extrabold text-white">
                  Zalo
                </span>
                <span className="truncate">Zalo</span>
              </button>

              <button
                type="button"
                disabled={isSaving || sendingKind !== null}
                onClick={() => void handleQuickAppLinkShare("messenger")}
                className="flex flex-col items-center justify-center gap-1 rounded-xl border border-indigo-200 bg-indigo-50/60 py-2 px-1 text-[11px] font-semibold text-indigo-700 hover:bg-indigo-100 transition disabled:opacity-50"
              >
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-gradient-to-tr from-[#00B2FF] to-[#A033FF] text-white">
                  <MessageCircle className="h-3.5 w-3.5" />
                </span>
                <span className="truncate">Messenger</span>
              </button>

              <button
                type="button"
                disabled={isSaving || sendingKind !== null}
                onClick={() => void handleQuickAppLinkShare("facebook")}
                className="flex flex-col items-center justify-center gap-1 rounded-xl border border-blue-200 bg-blue-50/60 py-2 px-1 text-[11px] font-semibold text-blue-700 hover:bg-blue-100 transition disabled:opacity-50"
              >
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-[#1877F2] text-xs font-extrabold text-white">
                  f
                </span>
                <span className="truncate">Facebook</span>
              </button>

              <button
                type="button"
                disabled={isSaving || sendingKind !== null}
                onClick={() => void handleQuickAppLinkShare("telegram")}
                className="flex flex-col items-center justify-center gap-1 rounded-xl border border-cyan-200 bg-cyan-50/60 py-2 px-1 text-[11px] font-semibold text-cyan-700 hover:bg-cyan-100 transition disabled:opacity-50"
              >
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-[#229ED9] text-white">
                  <Send className="h-3.5 w-3.5" />
                </span>
                <span className="truncate">Telegram</span>
              </button>

              <button
                type="button"
                disabled={isSaving || sendingKind !== null}
                onClick={() => void handleQuickAppLinkShare("system")}
                className="flex flex-col items-center justify-center gap-1 rounded-xl border border-slate-200 bg-slate-50 py-2 px-1 text-[11px] font-semibold text-slate-700 hover:bg-slate-100 transition disabled:opacity-50"
              >
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-slate-700 text-white">
                  <Share2 className="h-3.5 w-3.5" />
                </span>
                <span className="truncate">Hệ thống</span>
              </button>
            </div>

            <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
              <button
                type="button"
                disabled={sendingKind !== null}
                onClick={() => void handleSendAsOriginalFile()}
                className="inline-flex min-h-[42px] items-center justify-center gap-1.5 rounded-xl bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-xs hover:bg-emerald-700 transition disabled:opacity-60"
              >
                {sendingKind === "original" ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin shrink-0" />
                    <span>Đang đóng gói tệp...</span>
                  </>
                ) : (
                  <>
                    <FileOutput className="h-4 w-4 shrink-0" />
                    <span>
                      Gửi tệp gốc ({ORIGINAL_EXT_LABEL[document.type]})
                    </span>
                  </>
                )}
              </button>

              <button
                type="button"
                disabled={sendingKind !== null}
                onClick={() => void handleSendAsPdfFile()}
                className="inline-flex min-h-[42px] items-center justify-center gap-1.5 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 text-xs font-semibold text-rose-700 hover:bg-rose-600 hover:text-white transition disabled:opacity-60"
              >
                {sendingKind === "pdf" ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin shrink-0" />
                    <span>Đang tạo PDF...</span>
                  </>
                ) : (
                  <>
                    <FileDown className="h-4 w-4 shrink-0" />
                    <span>Gửi dưới dạng bản PDF</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {shareFeedback && (
            <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-2">
              {shareFeedback}
            </p>
          )}

          {errorMsg && (
            <p className="text-xs text-rose-600 bg-rose-50 border border-rose-200 rounded-xl px-3 py-2">
              {errorMsg}
            </p>
          )}
        </div>
      </div>
    </div>
  );
}

export default ShareCollabModal;
