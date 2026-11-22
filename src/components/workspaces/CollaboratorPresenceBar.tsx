import { useState } from "react";
import { Eye } from "lucide-react";
import type { CollaboratorPresence } from "../../services/realtimeCollab";

export interface CollaboratorPresenceBarProps {
  collaborators?: CollaboratorPresence[];
  selfSessionId?: string;
  collabPermission?: string;
  theme?: "light" | "dark";
}

function formatPresenceLocation(c: CollaboratorPresence): string {
  const parts: string[] = [];
  if (c.cellCoord) {
    parts.push("Đang ở ô " + c.cellCoord);
  } else if (typeof c.slideIndex === "number") {
    parts.push("Slide " + (c.slideIndex + 1));
  } else if (c.cursorLabel) {
    parts.push(c.cursorLabel);
  } else {
    parts.push("Đang trực tuyến");
  }
  if (c.isTyping) {
    parts.push("đang nhập...");
  }
  return parts.join(" • ");
}

function getInitials(name: string): string {
  const clean = (name || "U").trim();
  const words = clean.split(/\s+/);
  if (words.length >= 2) {
    return (words[0][0] + words[words.length - 1][0]).toUpperCase();
  }
  return clean.slice(0, 2).toUpperCase();
}

export function CollaboratorPresenceBar({
  collaborators = [],
  selfSessionId,
  collabPermission,
  theme = "light",
}: CollaboratorPresenceBarProps) {
  const isDark = theme === "dark";
  const [activeTipId, setActiveTipId] = useState<string | null>(null);

  const remoteUsers = collaborators.filter(
    (c) => !selfSessionId || c.sessionId !== selfSessionId
  );
  const isViewOnly = collabPermission === "view";

  if (remoteUsers.length === 0 && !isViewOnly) {
    return null;
  }

  const visibleUsers = remoteUsers.slice(0, 4);
  const extraCount = Math.max(0, remoteUsers.length - visibleUsers.length);

  return (
    <div className="flex items-center gap-1.5 shrink-0">
      {isViewOnly && (
        <span
          title="Bạn đang mở tệp ở chế độ Chỉ xem"
          className={
            "inline-flex items-center gap-1 px-2 py-1 rounded-lg border text-[11px] font-semibold whitespace-nowrap " +
            (isDark
              ? "bg-amber-500/15 border-amber-400/30 text-amber-200"
              : "bg-amber-50 border-amber-200 text-amber-800")
          }
        >
          <Eye className="h-3 w-3 shrink-0" />
          <span>Chỉ xem</span>
        </span>
      )}

      {visibleUsers.length > 0 && (
        <div
          className="flex items-center -space-x-1.5"
          aria-label="Người dùng đang cùng chỉnh sửa"
        >
          {visibleUsers.map((person) => {
            const locationText = formatPresenceLocation(person);
            const isOpen = activeTipId === person.sessionId;
            return (
              <div
                key={person.sessionId}
                className="relative group"
                onMouseEnter={() => setActiveTipId(person.sessionId)}
                onMouseLeave={() =>
                  setActiveTipId((prev) =>
                    prev === person.sessionId ? null : prev
                  )
                }
              >
                <button
                  type="button"
                  onClick={() =>
                    setActiveTipId((prev) =>
                      prev === person.sessionId ? null : person.sessionId
                    )
                  }
                  aria-label={person.displayName + " • " + locationText}
                  style={{
                    backgroundColor: person.color || "#2563EB",
                    borderColor: isDark ? "#141923" : "#FFFFFF",
                  }}
                  className="relative h-7 w-7 rounded-full border-2 text-[10px] font-bold text-white flex items-center justify-center shadow-sm transition-transform hover:scale-110 hover:z-20 focus:outline-none"
                >
                  {getInitials(person.displayName)}
                  <span
                    className={
                      "absolute -bottom-0.5 -right-0.5 h-2 w-2 rounded-full bg-emerald-400 ring-1 " +
                      (isDark ? "ring-[#141923]" : "ring-white") +
                      (person.isTyping ? " animate-ping" : "")
                    }
                  />
                </button>

                <div
                  role="tooltip"
                  className={
                    "pointer-events-none absolute right-0 top-full mt-2 z-50 w-max max-w-[230px] rounded-xl px-2.5 py-1.5 text-[11px] shadow-xl border transition-opacity duration-150 " +
                    (isOpen
                      ? "opacity-100"
                      : "opacity-0 group-hover:opacity-100") +
                    " " +
                    (isDark
                      ? "bg-[#1A2130] border-white/15 text-slate-100"
                      : "bg-slate-900 border-slate-800 text-white")
                  }
                >
                  <div className="flex items-center gap-1.5 font-semibold truncate">
                    <span
                      className="h-2 w-2 rounded-full shrink-0"
                      style={{ backgroundColor: person.color || "#2563EB" }}
                    />
                    <span className="truncate">{person.displayName}</span>
                  </div>
                  <p className="text-[10px] opacity-80 truncate mt-0.5">
                    {locationText}
                  </p>
                </div>
              </div>
            );
          })}

          {extraCount > 0 && (
            <span
              className={
                "h-7 w-7 rounded-full border-2 text-[10px] font-bold flex items-center justify-center " +
                (isDark
                  ? "bg-slate-800 border-[#141923] text-slate-200"
                  : "bg-slate-100 border-white text-slate-700")
              }
            >
              +{extraCount}
            </span>
          )}
        </div>
      )}
    </div>
  );
}
