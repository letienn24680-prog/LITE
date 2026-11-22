import { useEffect, useState } from "react";
import type { FormEvent } from "react";
import {
  AlertCircle,
  CheckCircle2,
  Database,
  Eye,
  EyeOff,
  Info,
  KeyRound,
  Loader2,
  Lock,
  LogIn,
  Mail,
  Send,
  Terminal,
  UserPlus,
  X,
} from "lucide-react";
import { isSupabaseConfigured, supabase } from "../../lib/supabase";
import { mapAuthErrorToVietnamese } from "../../services/cloudSync";
import {
  checkEmailAccess,
  requestEmailAccess,
} from "../../services/accessControl";

declare global {
  interface Window {
    liteDesktop?: {
      isDesktop: boolean;
      openExternal: (url: string) => Promise<void>;
      onAuthCallback: (cb: (callbackUrl: string) => void) => () => void;
    };
  }
}

export interface AuthModalProps {
  isOpen: boolean;
  initialMode?: "login" | "register";
  externalNotice?: string | null;
  onClose: () => void;
  onAuthSuccess?: () => void;
}

type AuthMode = "login" | "register";

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function AuthModal({
  isOpen,
  initialMode = "login",
  externalNotice = null,
  onClose,
  onAuthSuccess,
}: AuthModalProps) {
  const [mode, setMode] = useState<AuthMode>(initialMode);
  const [email, setEmail] = useState<string>("");
  const [password, setPassword] = useState<string>("");
  const [showPassword, setShowPassword] = useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [isGoogleLoading, setIsGoogleLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // Xin cấp quyền truy cập cho email mới
  const [showRequestBox, setShowRequestBox] = useState<boolean>(false);
  const [requestEmailInput, setRequestEmailInput] = useState<string>("");
  const [requestNoteInput, setRequestNoteInput] = useState<string>("");
  const [isRequestingAccess, setIsRequestingAccess] = useState<boolean>(false);

  const isBusy = isSubmitting || isGoogleLoading || isRequestingAccess;

  useEffect(() => {
    if (!isOpen) return;
    setMode(initialMode);
    setEmail("");
    setPassword("");
    setShowPassword(false);
    setIsSubmitting(false);
    setIsGoogleLoading(false);
    setErrorMessage(externalNotice || null);
    setSuccessMessage(null);
    setShowRequestBox(Boolean(externalNotice));
  }, [isOpen, initialMode, externalNotice]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape" && !isBusy) onClose();
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isOpen, isBusy, onClose]);

  if (!isOpen) return null;

  const isCloudReady = isSupabaseConfigured && supabase !== null;
  const isDesktopApp = Boolean(window.liteDesktop?.isDesktop);

  const handleSwitchMode = (nextMode: AuthMode) => {
    if (isBusy || nextMode === mode) return;
    setMode(nextMode);
    setErrorMessage(null);
    setSuccessMessage(null);
    setShowPassword(false);
  };

  const handleGoogleAuth = async () => {
    if (isBusy) return;
    if (!isSupabaseConfigured || !supabase) {
      setErrorMessage("Ứng dụng đang ở Chế độ cục bộ, chưa cấu hình Supabase.");
      return;
    }

    setIsGoogleLoading(true);
    setErrorMessage(null);
    setSuccessMessage(null);

    try {
      const redirectTarget = isDesktopApp
        ? "liteapp://auth/callback"
        : window.location.origin;

      const { data, error } = await supabase.auth.signInWithOAuth({
        provider: "google",
        options: {
          redirectTo: redirectTarget,
          skipBrowserRedirect: isDesktopApp,
          scopes:
            "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly",
          queryParams: { access_type: "offline", prompt: "consent" },
        },
      });

      if (error) {
        setErrorMessage(mapAuthErrorToVietnamese(error.message));
        setIsGoogleLoading(false);
        return;
      }

      if (isDesktopApp && data?.url && window.liteDesktop?.openExternal) {
        await window.liteDesktop.openExternal(data.url);
        setSuccessMessage(
          "Đã mở trình duyệt để đăng nhập Google. Sau khi chọn tài khoản Gmail, trình duyệt sẽ tự động chuyển bạn quay lại ứng dụng LITE."
        );
        setIsGoogleLoading(false);
      }
    } catch (error) {
      const rawMessage =
        error instanceof Error ? error.message : "Network request failed";
      setErrorMessage(mapAuthErrorToVietnamese(rawMessage));
      setIsGoogleLoading(false);
    }
  };

  const handleRequestAccessSubmit = async () => {
    if (isBusy) return;
    const targetMail = (requestEmailInput || email).trim();
    if (!targetMail || !EMAIL_PATTERN.test(targetMail)) {
      setErrorMessage("Vui lòng nhập địa chỉ Gmail/Email hợp lệ để gửi yêu cầu.");
      return;
    }
    setIsRequestingAccess(true);
    setErrorMessage(null);
    setSuccessMessage(null);
    const res = await requestEmailAccess(targetMail, requestNoteInput);
    setIsRequestingAccess(false);
    if (res.ok) {
      setSuccessMessage(res.message);
      setRequestEmailInput("");
      setRequestNoteInput("");
    } else {
      setErrorMessage(res.message);
    }
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (isBusy) return;
    if (!isSupabaseConfigured || !supabase) {
      setErrorMessage("Chưa cấu hình biến môi trường Supabase nên không thể đăng nhập.");
      return;
    }

    const trimmedEmail = email.trim().toLowerCase();
    if (!trimmedEmail || !EMAIL_PATTERN.test(trimmedEmail)) {
      setErrorMessage("Vui lòng nhập địa chỉ email hợp lệ.");
      return;
    }
    if (!password) {
      setErrorMessage("Vui lòng nhập mật khẩu.");
      return;
    }
    if (mode === "register" && password.length < 6) {
      setErrorMessage("Mật khẩu đăng ký phải có ít nhất 6 ký tự.");
      return;
    }

    setIsSubmitting(true);
    setErrorMessage(null);
    setSuccessMessage(null);

    try {
      // Kiểm tra email có trong danh sách cho phép hay chưa
      const access = await checkEmailAccess(trimmedEmail);
      if (!access.allowed) {
        await requestEmailAccess(trimmedEmail, "Tự động ghi nhận khi thử đăng ký/đăng nhập");
        setRequestEmailInput(trimmedEmail);
        setShowRequestBox(true);
        setErrorMessage(
          access.reason ||
            `Email ${trimmedEmail} chưa được cấp quyền truy cập. Yêu cầu đã được gửi tới chủ sở hữu để duyệt từ xa.`
        );
        setIsSubmitting(false);
        return;
      }

      if (mode === "login") {
        const { data, error } = await supabase.auth.signInWithPassword({
          email: trimmedEmail,
          password,
        });
        if (error) {
          setErrorMessage(mapAuthErrorToVietnamese(error.message));
          return;
        }
        if (data.session || data.user) {
          if (onAuthSuccess) onAuthSuccess();
          else onClose();
          return;
        }
        setErrorMessage("Không thể khởi tạo phiên đăng nhập. Vui lòng thử lại.");
      } else {
        const { data, error } = await supabase.auth.signUp({
          email: trimmedEmail,
          password,
        });
        if (error) {
          setErrorMessage(mapAuthErrorToVietnamese(error.message));
          return;
        }
        if (data.session) {
          if (onAuthSuccess) onAuthSuccess();
          else onClose();
          return;
        }
        setPassword("");
        setSuccessMessage(
          `Đăng ký thành công! Vui lòng kiểm tra hộp thư ${trimmedEmail} và bấm vào liên kết xác nhận để hoàn tất.`
        );
      }
    } catch (error) {
      const rawMessage =
        error instanceof Error ? error.message : "Network request failed";
      setErrorMessage(mapAuthErrorToVietnamese(rawMessage));
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="auth-modal-title"
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm pb-safe"
    >
      <div
        onClick={(event) => {
          if (event.target === event.currentTarget && !isBusy) onClose();
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        <div className="w-full max-w-md max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
          <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <span className="h-9 w-9 shrink-0 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-[#4F46E5]">
                {!isCloudReady ? (
                  <Database className="h-4 w-4" />
                ) : mode === "login" ? (
                  <LogIn className="h-4 w-4" />
                ) : (
                  <UserPlus className="h-4 w-4" />
                )}
              </span>
              <div className="min-w-0">
                <h2
                  id="auth-modal-title"
                  className="text-sm font-semibold text-slate-900 truncate"
                >
                  {!isCloudReady
                    ? "Kết nối Supabase (Chế độ Local)"
                    : mode === "login"
                      ? "Đăng nhập tài khoản LITE"
                      : "Đăng ký tài khoản đồng bộ LITE"}
                </h2>
                <p className="text-xs text-slate-500 truncate mt-0.5">
                  {!isCloudReady
                    ? "Ứng dụng đang chạy cục bộ trên thiết bị này"
                    : "Đồng bộ tệp & cộng tác đa thiết bị"}
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              disabled={isBusy}
              aria-label="Đóng hộp thoại tài khoản"
              className="shrink-0 min-h-[38px] min-w-[38px] rounded-xl bg-white hover:bg-slate-100 border border-[#E6E0D6] text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors disabled:opacity-50"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          <div className="flex-1 min-h-0 overflow-y-auto">
            {!isCloudReady ? (
              <div className="p-4 sm:p-6 space-y-4">
                <div className="rounded-2xl border border-amber-200 bg-amber-50/75 p-4 flex items-start gap-3">
                  <Info className="h-5 w-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="space-y-1 text-xs leading-relaxed text-amber-950">
                    <p className="font-semibold">Đang hoạt động ở Chế độ Cục bộ</p>
                    <p className="text-amber-800">
                      Chưa tìm thấy cấu hình VITE_SUPABASE_URL hoặc VITE_SUPABASE_ANON_KEY.
                    </p>
                  </div>
                </div>
                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 space-y-2 text-xs text-slate-700">
                  <div className="flex items-center gap-2 font-semibold text-slate-800">
                    <Terminal className="h-4 w-4 text-[#4F46E5]" />
                    <span>Cấu hình .env trước khi đóng gói App</span>
                  </div>
                  <pre className="rounded-xl bg-slate-900 text-slate-100 p-3 text-[11px] font-mono overflow-x-auto">
                    {`VITE_SUPABASE_URL=https://your-project.supabase.co\nVITE_SUPABASE_ANON_KEY=your-anon-key`}
                  </pre>
                </div>
                <button
                  type="button"
                  onClick={onClose}
                  className="w-full min-h-[44px] rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-sm font-semibold"
                >
                  Đóng & Tiếp tục dùng Local Mode
                </button>
              </div>
            ) : (
              <div className="p-4 sm:p-6 space-y-4">
                <div
                  role="tablist"
                  aria-label="Chế độ xác thực"
                  className="grid grid-cols-2 gap-1.5 p-1 rounded-2xl bg-[#FAF8F5] border border-[#EAE5DC]"
                >
                  <button
                    type="button"
                    role="tab"
                    aria-selected={mode === "login"}
                    disabled={isBusy}
                    onClick={() => handleSwitchMode("login")}
                    className={`min-h-[38px] rounded-xl text-xs font-semibold transition-all flex items-center justify-center gap-1.5 ${
                      mode === "login"
                        ? "bg-white text-[#4F46E5] shadow-sm ring-1 ring-[#E6E0D6]"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <LogIn className="h-3.5 w-3.5" />
                    <span>Đăng nhập</span>
                  </button>
                  <button
                    type="button"
                    role="tab"
                    aria-selected={mode === "register"}
                    disabled={isBusy}
                    onClick={() => handleSwitchMode("register")}
                    className={`min-h-[38px] rounded-xl text-xs font-semibold transition-all flex items-center justify-center gap-1.5 ${
                      mode === "register"
                        ? "bg-white text-[#4F46E5] shadow-sm ring-1 ring-[#E6E0D6]"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <UserPlus className="h-3.5 w-3.5" />
                    <span>Đăng ký</span>
                  </button>
                </div>

                {errorMessage && (
                  <div
                    role="alert"
                    className="rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-3 flex items-start gap-2.5 text-xs text-rose-800"
                  >
                    <AlertCircle className="h-4 w-4 text-rose-600 shrink-0 mt-0.5" />
                    <span className="leading-relaxed">{errorMessage}</span>
                  </div>
                )}

                {successMessage && (
                  <div
                    role="status"
                    className="rounded-2xl border border-emerald-200 bg-emerald-50 px-3.5 py-3 space-y-2 text-xs text-emerald-900"
                  >
                    <div className="flex items-start gap-2.5">
                      <CheckCircle2 className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5" />
                      <span className="leading-relaxed">{successMessage}</span>
                    </div>
                  </div>
                )}

                {/* Nút Tiếp tục với Google (cả tab Đăng nhập & Đăng ký) */}
                <button
                  type="button"
                  onClick={() => void handleGoogleAuth()}
                  disabled={isBusy}
                  className="w-full min-h-[44px] rounded-xl bg-white hover:bg-[#FAF8F5] text-slate-800 border border-[#DDD6CA] shadow-sm text-sm font-semibold flex items-center justify-center gap-2.5 transition-colors disabled:opacity-60"
                >
                  {isGoogleLoading ? (
                    <>
                      <Loader2 className="h-4 w-4 animate-spin text-[#4F46E5]" />
                      <span>Đang kết nối Google...</span>
                    </>
                  ) : (
                    <>
                      <svg viewBox="0 0 24 24" className="h-4 w-4 shrink-0" aria-hidden="true">
                        <path fill="#4285F4" d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.665-5.17 3.665-9.17z" />
                        <path fill="#34A853" d="M12 24c3.3 0 6.08-1.09 8.105-2.96l-3.88-3.05c-1.1.74-2.51 1.18-4.225 1.18-3.24 0-5.99-2.19-6.97-5.13H1.02v3.14C3.035 21.18 7.21 24 12 24z" />
                        <path fill="#FBBC05" d="M5.03 14.04c-.25-.74-.39-1.53-.39-2.34s.14-1.6.39-2.34V6.22H1.02C.37 7.51 0 8.97 0 10.5s.37 2.99 1.02 4.28l4.01-3.14z" />
                        <path fill="#EA4335" d="M12 4.75c1.8 0 3.41.62 4.68 1.84l3.51-3.51C18.07 1.19 15.3 0 12 0 7.21 0 3.035 2.82 1.02 6.86l4.01 3.14c.98-2.94 3.73-5.25 6.97-5.25z" />
                      </svg>
                      <span>Tiếp tục với Google</span>
                    </>
                  )}
                </button>

                {/* Khu vực xin cấp quyền truy cập từ xa cho người dùng mới */}
                <div className="rounded-xl border border-indigo-100 bg-indigo-50/50 p-2.5">
                  <button
                    type="button"
                    onClick={() => setShowRequestBox((v) => !v)}
                    className="w-full flex items-center justify-between text-xs font-semibold text-indigo-700 hover:text-indigo-900"
                  >
                    <span className="inline-flex items-center gap-1.5">
                      <KeyRound className="h-3.5 w-3.5" />
                      <span>Tài khoản mới? Gửi Gmail xin cấp quyền dùng App</span>
                    </span>
                    <span className="text-[11px] underline">
                      {showRequestBox ? "Thu gọn" : "Mở"}
                    </span>
                  </button>

                  {showRequestBox && (
                    <div className="mt-2.5 pt-2.5 border-t border-indigo-100 space-y-2">
                      <p className="text-[11px] text-slate-600 leading-relaxed">
                        Nhập địa chỉ Gmail của bạn để gửi yêu cầu. Chủ sở hữu sẽ duyệt trực tiếp từ xa — bạn <strong>không cần cài lại app</strong>.
                      </p>
                      <input
                        type="email"
                        value={requestEmailInput}
                        onChange={(e) => setRequestEmailInput(e.target.value)}
                        placeholder="nhap-gmail-cua-ban@gmail.com"
                        className="w-full h-9 rounded-lg border border-indigo-200 bg-white px-3 text-xs text-slate-900 outline-none focus:border-indigo-500"
                      />
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={requestNoteInput}
                          onChange={(e) => setRequestNoteInput(e.target.value)}
                          placeholder="Tên của bạn (không bắt buộc)..."
                          className="flex-1 h-9 rounded-lg border border-indigo-200 bg-white px-3 text-xs text-slate-900 outline-none focus:border-indigo-500"
                        />
                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => void handleRequestAccessSubmit()}
                          className="inline-flex items-center gap-1 px-3 h-9 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shrink-0 disabled:opacity-60"
                        >
                          {isRequestingAccess ? (
                            <Loader2 className="h-3.5 w-3.5 animate-spin" />
                          ) : (
                            <Send className="h-3.5 w-3.5" />
                          )}
                          <span>Gửi duyệt</span>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="relative flex items-center justify-center">
                  <div aria-hidden="true" className="absolute inset-x-0 top-1/2 border-t border-[#EAE5DC]" />
                  <span className="relative bg-white px-3 text-[11px] font-medium uppercase tracking-wider text-slate-400">
                    hoặc dùng mật khẩu
                  </span>
                </div>

                <form onSubmit={handleSubmit} className="space-y-3.5" noValidate>
                  <div className="space-y-1">
                    <label htmlFor="auth-email-input" className="block text-xs font-semibold text-slate-700">
                      Email
                    </label>
                    <div className="relative">
                      <Mail className="h-4 w-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                      <input
                        id="auth-email-input"
                        type="email"
                        autoComplete="email"
                        value={email}
                        disabled={isBusy}
                        onChange={(event) => {
                          setEmail(event.target.value);
                          if (errorMessage) setErrorMessage(null);
                        }}
                        placeholder="ten.ban@gmail.com"
                        className="w-full h-10 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5]/60 pl-10 pr-3.5 text-sm text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-[#4F46E5] focus:outline-none"
                      />
                    </div>
                  </div>

                  <div className="space-y-1">
                    <label htmlFor="auth-password-input" className="block text-xs font-semibold text-slate-700">
                      Mật khẩu
                    </label>
                    <div className="relative">
                      <Lock className="h-4 w-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                      <input
                        id="auth-password-input"
                        type={showPassword ? "text" : "password"}
                        autoComplete={mode === "login" ? "current-password" : "new-password"}
                        value={password}
                        disabled={isBusy}
                        onChange={(event) => {
                          setPassword(event.target.value);
                          if (errorMessage) setErrorMessage(null);
                        }}
                        placeholder={mode === "register" ? "Tối thiểu 6 ký tự..." : "Nhập mật khẩu..."}
                        className="w-full h-10 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5]/60 pl-10 pr-11 text-sm text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-[#4F46E5] focus:outline-none"
                      />
                      <button
                        type="button"
                        onClick={() => setShowPassword((prev) => !prev)}
                        disabled={isBusy}
                        aria-label={showPassword ? "Ẩn mật khẩu" : "Hiện mật khẩu"}
                        className="absolute right-1.5 top-1/2 -translate-y-1/2 h-8 w-8 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 flex items-center justify-center"
                      >
                        {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                      </button>
                    </div>
                  </div>

                  <button
                    type="submit"
                    disabled={isBusy}
                    className="w-full min-h-[42px] rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-sm font-semibold shadow-sm shadow-indigo-600/20 flex items-center justify-center gap-2 transition-colors disabled:opacity-60"
                  >
                    {isSubmitting ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin" />
                        <span>{mode === "login" ? "Đang đăng nhập..." : "Đang tạo tài khoản..."}</span>
                      </>
                    ) : mode === "login" ? (
                      <>
                        <LogIn className="h-4 w-4" />
                        <span>Đăng nhập</span>
                      </>
                    ) : (
                      <>
                        <UserPlus className="h-4 w-4" />
                        <span>Đăng ký tài khoản</span>
                      </>
                    )}
                  </button>
                </form>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
