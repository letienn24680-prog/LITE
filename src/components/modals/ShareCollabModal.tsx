import { useState } from "react";
import { X, Link2, Check, Users, Globe, Lock, Eye, Edit3 } from "lucide-react";

export function ShareCollabModal(props: any) {
  if (props.isOpen === false || props.open === false) return null;
  const handleClose = props.onClose || props.onDismiss || (() => {});
  const doc = props.document || props.doc || props.targetDoc || null;
  const docTitle = (doc && (doc.title || doc.name)) || "Tài liệu LITE";
  const initialPerm = props.permission || (doc && doc.permission) || "edit";
  const [accessMode, setAccessMode] = useState<"restricted" | "link">("link");
  const [permission, setPermission] = useState<"view" | "edit">(initialPerm === "view" ? "view" : "edit");
  const [copied, setCopied] = useState(false);
  const shareToken = props.shareToken || (doc && (doc.share_token || doc.id)) || "shared";
  const shareUrl = props.shareUrl || window.location.origin + "/?share=" + encodeURIComponent(shareToken);
  const collaborators: any[] = Array.isArray(props.collaborators) ? props.collaborators : [];

  const handleCopy = async () => {
    try { await navigator.clipboard.writeText(shareUrl); } catch {
      const input = document.createElement("input");
      input.value = shareUrl; document.body.appendChild(input);
      input.select(); document.execCommand("copy"); document.body.removeChild(input);
    }
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
    const updater = props.onCreateOrUpdateLink || props.onUpdatePermission;
    if (typeof updater === "function") updater(permission);
  };

  const handlePermissionChange = (nextPerm: "view" | "edit") => {
    setPermission(nextPerm);
    if (typeof props.onPermissionChange === "function") props.onPermissionChange(nextPerm);
    const updater = props.onCreateOrUpdateLink || props.onUpdatePermission;
    if (typeof updater === "function") updater(nextPerm);
  };

  const handleAccessChange = (mode: "restricted" | "link") => {
    setAccessMode(mode);
    const revoker = props.onRevokeLink || props.onRevoke;
    const updater = props.onCreateOrUpdateLink || props.onUpdatePermission;
    if (mode === "restricted" && typeof revoker === "function") revoker();
    else if (mode === "link" && typeof updater === "function") updater(permission);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-xs p-4" onClick={handleClose}>
      <div className="w-full max-w-md rounded-2xl bg-white border border-slate-200 shadow-2xl overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-100">
          <div className="min-w-0 pr-3">
            <h3 className="text-base font-semibold text-slate-900 truncate">Chia sẻ "{docTitle}"</h3>
            <p className="text-xs text-slate-500 mt-0.5">Cộng tác xem và chỉnh sửa trực tuyến</p>
          </div>
          <button type="button" onClick={handleClose} className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition"><X className="w-4 h-4" /></button>
        </div>
        <div className="p-5 space-y-4">
          {collaborators.length > 0 && (
            <div className="space-y-2">
              <div className="flex items-center gap-1.5 text-xs font-medium text-slate-600"><Users className="w-3.5 h-3.5 text-indigo-600" /><span>Đang trực tuyến ({collaborators.length})</span></div>
              <div className="flex flex-wrap gap-2">
                {collaborators.map((c: any, idx: number) => (
                  <span key={(c && (c.sessionId || c.userId)) || idx} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-700">
                    <span className="w-2 h-2 rounded-full" style={{ backgroundColor: (c && c.color) || "#4F46E5" }} />
                    {(c && (c.displayName || c.email)) || "Thành viên"}
                  </span>
                ))}
              </div>
            </div>
          )}
          <div className="space-y-2">
            <label className="text-xs font-semibold uppercase tracking-wider text-slate-500">Quyền truy cập chung</label>
            <div className="grid grid-cols-2 gap-2">
              <button type="button" onClick={() => handleAccessChange("restricted")} className={"flex items-center gap-2 p-2.5 rounded-xl border text-left text-xs font-medium transition " + (accessMode === "restricted" ? "border-indigo-600 bg-indigo-50/70 text-indigo-700" : "border-slate-200 text-slate-600 hover:bg-slate-50")}><Lock className="w-4 h-4 shrink-0" /><span>Chỉ mình tôi</span></button>
              <button type="button" onClick={() => handleAccessChange("link")} className={"flex items-center gap-2 p-2.5 rounded-xl border text-left text-xs font-medium transition " + (accessMode === "link" ? "border-indigo-600 bg-indigo-50/70 text-indigo-700" : "border-slate-200 text-slate-600 hover:bg-slate-50")}><Globe className="w-4 h-4 shrink-0" /><span>Bất kỳ ai có link</span></button>
            </div>
          </div>
          {accessMode === "link" && (
            <>
              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-50 border border-slate-200/80">
                <span className="text-xs font-medium text-slate-700">Vai trò người nhận link</span>
                <div className="flex items-center gap-1 bg-white p-1 rounded-lg border border-slate-200">
                  <button type="button" onClick={() => handlePermissionChange("view")} className={"inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition " + (permission === "view" ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-100")}><Eye className="w-3.5 h-3.5" />Chỉ xem</button>
                  <button type="button" onClick={() => handlePermissionChange("edit")} className={"inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition " + (permission === "edit" ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100")}><Edit3 className="w-3.5 h-3.5" />Chỉnh sửa</button>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <input type="text" readOnly value={shareUrl} className="flex-1 text-xs px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-slate-600 outline-none" />
                <button type="button" onClick={handleCopy} className="inline-flex items-center gap-1.5 px-3.5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shrink-0 transition">
                  {copied ? (<><Check className="w-3.5 h-3.5" />Đã chép</>) : (<><Link2 className="w-3.5 h-3.5" />Sao chép</>)}
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
export default ShareCollabModal;
