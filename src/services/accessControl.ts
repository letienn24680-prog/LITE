import { isSupabaseConfigured, supabase } from "../lib/supabase";

export type AccessStatus = "approved" | "pending" | "denied" | "unconfigured";

export interface AllowedEmailRecord {
  email: string;
  status: "approved" | "pending";
  note?: string;
  added_by?: string;
  created_at?: string;
}

const ALLOWED_EMAILS_TABLE = "app_allowed_emails";
const OWNER_STORAGE_KEY = "lite_app_owner_email_v1";
const DEFAULT_OWNER_EMAILS = [
  (import.meta.env.VITE_APP_OWNER_EMAIL as string | undefined)?.trim().toLowerCase() || "",
  "letienn24680@gmail.com",
].filter(Boolean);

export function normalizeEmail(raw: string): string {
  return raw.trim().toLowerCase();
}

export function getKnownOwnerEmails(): string[] {
  const set = new Set<string>(DEFAULT_OWNER_EMAILS);
  try {
    const saved = window.localStorage.getItem(OWNER_STORAGE_KEY);
    if (saved) set.add(normalizeEmail(saved));
  } catch {
    // Ignore storage errors
  }
  return Array.from(set);
}

export function rememberOwnerEmail(email: string): void {
  const norm = normalizeEmail(email);
  if (!norm) return;
  try {
    window.localStorage.setItem(OWNER_STORAGE_KEY, norm);
  } catch {
    // Ignore storage errors
  }
}

export function isAppOwner(email?: string | null): boolean {
  if (!email) return false;
  const norm = normalizeEmail(email);
  return getKnownOwnerEmails().includes(norm);
}

export async function checkEmailAccess(
  _email: string
): Promise<{ allowed: boolean; status: AccessStatus; reason?: string }> {
  return { allowed: true, status: "approved" };
}

export async function requestEmailAccess(
  email: string,
  note?: string
): Promise<{ ok: boolean; message: string }> {
  const norm = normalizeEmail(email);
  if (!norm || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(norm)) {
    return { ok: false, message: "Vui lòng nhập đúng định dạng địa chỉ Email / Gmail." };
  }

  if (isAppOwner(norm)) {
    return { ok: true, message: "Email này là chủ sở hữu ứng dụng và đã có toàn quyền truy cập." };
  }

  if (!isSupabaseConfigured || !supabase) {
    return { ok: false, message: "Ứng dụng chưa kết nối Supabase nên không thể gửi yêu cầu." };
  }

  try {
    const existing = await supabase
      .from(ALLOWED_EMAILS_TABLE)
      .select("email, status")
      .eq("email", norm)
      .maybeSingle();

    if (!existing.error && existing.data) {
      const st = (existing.data as { status?: string }).status;
      if (st === "approved") {
        return {
          ok: true,
          message: `Email ${norm} đã được duyệt sẵn! Bạn có thể bấm Đăng nhập hoặc Tiếp tục với Google ngay bây giờ.`,
        };
      }
      return {
        ok: true,
        message: `Yêu cầu của ${norm} đã được ghi nhận từ trước và đang chờ chủ sở hữu duyệt.`,
      };
    }

    const { error } = await supabase.from(ALLOWED_EMAILS_TABLE).insert({
      email: norm,
      status: "pending",
      note: note?.trim() || "Yêu cầu từ màn hình đăng nhập",
    });

    if (error) {
      return {
        ok: false,
        message: `Không thể gửi yêu cầu (${error.message}). Hãy báo chủ sở hữu thêm trực tiếp email ${norm}.`,
      };
    }

    return {
      ok: true,
      message: `Đã gửi yêu cầu cấp quyền cho ${norm}! Khi chủ sở hữu bấm Duyệt trên hệ thống, bạn chỉ cần đăng nhập lại là dùng được ngay.`,
    };
  } catch (err) {
    return {
      ok: false,
      message: err instanceof Error ? err.message : "Lỗi kết nối mạng khi gửi yêu cầu.",
    };
  }
}

export async function listAllowedEmails(): Promise<{
  records: AllowedEmailRecord[];
  tableMissing?: boolean;
  error?: string;
}> {
  if (!isSupabaseConfigured || !supabase) {
    return { records: [], error: "Chưa cấu hình Supabase." };
  }

  try {
    const { data, error } = await supabase
      .from(ALLOWED_EMAILS_TABLE)
      .select("email, status, note, added_by, created_at")
      .order("created_at", { ascending: false });

    if (error) {
      const msg = (error.message || "").toLowerCase();
      if (msg.includes("does not exist") || msg.includes("relation") || error.code === "42P01") {
        return { records: [], tableMissing: true };
      }
      return { records: [], error: error.message };
    }

    return { records: (data ?? []) as AllowedEmailRecord[] };
  } catch (err) {
    return {
      records: [],
      error: err instanceof Error ? err.message : "Không thể tải danh sách email.",
    };
  }
}

export async function addOrApproveEmail(
  email: string,
  note?: string,
  addedBy?: string
): Promise<{ ok: boolean; error?: string }> {
  const norm = normalizeEmail(email);
  if (!norm || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(norm)) {
    return { ok: false, error: "Địa chỉ email không hợp lệ." };
  }
  if (!isSupabaseConfigured || !supabase) {
    return { ok: false, error: "Chưa kết nối Supabase." };
  }

  const { error } = await supabase.from(ALLOWED_EMAILS_TABLE).upsert(
    {
      email: norm,
      status: "approved",
      note: note?.trim() || "Được chủ sở hữu cấp quyền",
      added_by: addedBy || "owner",
    },
    { onConflict: "email" }
  );

  if (error) {
    return { ok: false, error: error.message };
  }
  return { ok: true };
}

export async function removeAllowedEmail(
  email: string
): Promise<{ ok: boolean; error?: string }> {
  const norm = normalizeEmail(email);
  if (!isSupabaseConfigured || !supabase) {
    return { ok: false, error: "Chưa kết nối Supabase." };
  }
  const { error } = await supabase
    .from(ALLOWED_EMAILS_TABLE)
    .delete()
    .eq("email", norm);

  if (error) {
    return { ok: false, error: error.message };
  }
  return { ok: true };
}
