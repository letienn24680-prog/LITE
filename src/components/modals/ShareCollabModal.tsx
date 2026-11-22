import { useEffect, useState } from "react";
import {
  Check,
  Edit3,
  Eye,
  Globe,
  Link2,
  Loader2,
  Lock,
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

    if (document.shareToken && document.shareAccess === "link") {
      const initialPerm: CollabPermission =
        document.collabPermission === "view" ? "view" : "edit";
      setPermission(initialPerm);
      setShareUrl(buildCollabShareUrl(document.shareToken));
    } else {
      setPermission("private");
      setShareUrl("");
    }

    setIsLoading(true);
    void getShareLinkInfo(ownerId, document.id)
      .then((info) => {
        if (cancelled) return;
        setPermission(info.permission);
        setShareUrl(info.shareUrl || "");
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
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [isOpen, onClose]);

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

  const handleCopyLink = async () => {
    if (!shareUrl) return;
    try {
      await navigator.clipboard.writeText(shareUrl);
    } catch {
      const input = window.document.createElement("input");
      input.value = shareUrl;
      window.document.body.appendChild(input);
      input.select();
      window.document.execCommand("copy");
      window.document.body.removeChild(input);
    }
    setCopied(true);
    window.setTimeout(() => setCopied(false), 2200);
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
        className="w-full max-w-md rounded-2xl bg-white border border-slate-200 shadow-2xl overflow-hidden text-slate-900"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-100">
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

        <div className="p-5 space-y-4">
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
            <div className="pt-1 space-y-2">
              <div className="flex items-center gap-1.5 text-[11px] font-medium text-emerald-700">
                <Globe className="w-3.5 h-3.5" />
                <span>Liên kết cộng tác đang hoạt động</span>
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
