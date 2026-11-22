import { useEffect, useState } from 'react';
import type { FormEvent } from 'react';
import {
  AlertCircle,
  CheckCircle2,
  Database,
  Eye,
  EyeOff,
  Info,
  Loader2,
  Lock,
  LogIn,
  Mail,
  Terminal,
  UserPlus,
  X,
} from 'lucide-react';
import { isSupabaseConfigured, supabase } from '../../lib/supabase';
import { mapAuthErrorToVietnamese } from '../../services/cloudSync';

export interface AuthModalProps {
  isOpen: boolean;
  initialMode?: 'login' | 'register';
  onClose: () => void;
  onAuthSuccess?: () => void;
}

type AuthMode = 'login' | 'register';

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function AuthModal({
  isOpen,
  initialMode = 'login',
  onClose,
  onAuthSuccess,
}: AuthModalProps) {
  const [mode, setMode] = useState<AuthMode>(initialMode);
  const [email, setEmail] = useState<string>('');
  const [password, setPassword] = useState<string>('');
  const [showPassword, setShowPassword] = useState<boolean>(false);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [isGoogleLoading, setIsGoogleLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const isBusy = isSubmitting || isGoogleLoading;

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    setMode(initialMode);
    setEmail('');
    setPassword('');
    setShowPassword(false);
    setIsSubmitting(false);
    setIsGoogleLoading(false);
    setErrorMessage(null);
    setSuccessMessage(null);
  }, [isOpen, initialMode]);

  useEffect(() => {
    if (!isOpen) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !isBusy) {
        onClose();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, isBusy, onClose]);

  if (!isOpen) {
    return null;
  }

  const isCloudReady = isSupabaseConfigured && supabase !== null;

  const handleSwitchMode = (nextMode: AuthMode) => {
    if (isBusy || nextMode === mode) {
      return;
    }
    setMode(nextMode);
    setErrorMessage(null);
    setSuccessMessage(null);
    setShowPassword(false);
  };

  const handleGoogleAuth = async () => {
    if (isBusy) {
      return;
    }

    if (!isSupabaseConfigured || !supabase) {
      setErrorMessage(
        'Ứng dụng đang ở Chế độ cục bộ (Local Mode), chưa cấu hình Supabase nên chưa thể đăng nhập bằng Google.'
      );
      return;
    }

    setIsGoogleLoading(true);
    setErrorMessage(null);
    setSuccessMessage(null);

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
        setErrorMessage(mapAuthErrorToVietnamese(error.message));
        setIsGoogleLoading(false);
      }
    } catch (error) {
      const rawMessage =
        error instanceof Error ? error.message : 'Network request failed';
      setErrorMessage(mapAuthErrorToVietnamese(rawMessage));
      setIsGoogleLoading(false);
    }
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (isBusy) {
      return;
    }

    if (!isSupabaseConfigured || !supabase) {
      setErrorMessage(
        'Chưa cấu hình biến môi trường Supabase nên không thể đăng nhập.'
      );
      return;
    }

    const trimmedEmail = email.trim();
    if (!trimmedEmail || !EMAIL_PATTERN.test(trimmedEmail)) {
      setErrorMessage('Vui lòng nhập địa chỉ email hợp lệ.');
      return;
    }

    if (!password) {
      setErrorMessage('Vui lòng nhập mật khẩu.');
      return;
    }

    if (mode === 'register' && password.length < 6) {
      setErrorMessage('Mật khẩu đăng ký phải có ít nhất 6 ký tự.');
      return;
    }

    setIsSubmitting(true);
    setErrorMessage(null);
    setSuccessMessage(null);

    try {
      if (mode === 'login') {
        const { data, error } = await supabase.auth.signInWithPassword({
          email: trimmedEmail,
          password,
        });

        if (error) {
          setErrorMessage(mapAuthErrorToVietnamese(error.message));
          return;
        }

        if (data.session || data.user) {
          if (onAuthSuccess) {
            onAuthSuccess();
          } else {
            onClose();
          }
          return;
        }

        setErrorMessage(
          'Không thể khởi tạo phiên đăng nhập. Vui lòng thử lại.'
        );
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
          if (onAuthSuccess) {
            onAuthSuccess();
          } else {
            onClose();
          }
          return;
        }

        setPassword('');
        setSuccessMessage(
          `Đăng ký thành công! Vui lòng kiểm tra hộp thư ${trimmedEmail} (kể cả mục Thư rác / Spam) và bấm vào liên kết xác nhận để hoàn tất kích hoạt tài khoản.`
        );
      }
    } catch (error) {
      const rawMessage =
        error instanceof Error ? error.message : 'Network request failed';
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
          if (event.target === event.currentTarget && !isBusy) {
            onClose();
          }
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        <div className="w-full max-w-md max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
          {/* Header */}
          <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <span className="h-9 w-9 shrink-0 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-[#4F46E5]">
                {!isCloudReady ? (
                  <Database className="h-4 w-4" />
                ) : mode === 'login' ? (
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
                    ? 'Kết nối Supabase (Chế độ Local)'
                    : mode === 'login'
                      ? 'Đăng nhập tài khoản LITE'
                      : 'Tạo tài khoản đồng bộ LITE'}
                </h2>
                <p className="text-xs text-slate-500 truncate mt-0.5">
                  {!isCloudReady
                    ? 'Ứng dụng đang chạy cục bộ trên trình duyệt này'
                    : 'Đồng bộ tệp và chỉnh sửa giữa máy tính & điện thoại'}
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              disabled={isBusy}
              aria-label="Đóng hộp thoại tài khoản"
              className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-white hover:bg-slate-100 border border-[#E6E0D6] text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Body */}
          <div className="flex-1 min-h-0 overflow-y-auto">
            {!isCloudReady ? (
              <div className="p-4 sm:p-6 space-y-4">
                <div className="rounded-2xl border border-amber-200 bg-amber-50/75 p-4 flex items-start gap-3">
                  <Info className="h-5 w-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="space-y-1 text-xs leading-relaxed text-amber-950">
                    <p className="font-semibold">
                      Đang hoạt động ở Chế độ Cục bộ (Local Mode)
                    </p>
                    <p className="text-amber-800">
                      Chưa tìm thấy cấu hình{' '}
                      <code className="font-mono font-semibold bg-amber-100/90 px-1 py-0.5 rounded">
                        VITE_SUPABASE_URL
                      </code>{' '}
                      hoặc{' '}
                      <code className="font-mono font-semibold bg-amber-100/90 px-1 py-0.5 rounded">
                        VITE_SUPABASE_ANON_KEY
                      </code>
                      . Tài liệu của bạn vẫn được lưu an toàn trong trình duyệt.
                    </p>
                  </div>
                </div>

                <div className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 space-y-3">
                  <div className="flex items-center gap-2 text-xs font-semibold text-slate-800">
                    <Terminal className="h-4 w-4 text-[#4F46E5] shrink-0" />
                    <span>Hướng dẫn bật đồng bộ đám mây Supabase</span>
                  </div>

                  <ol className="space-y-2.5 text-xs text-slate-700 leading-relaxed">
                    <li className="flex items-start gap-2.5">
                      <span className="h-5 w-5 rounded-full bg-indigo-100 text-[#4F46E5] font-semibold text-[11px] flex items-center justify-center shrink-0 mt-0.5">
                        1
                      </span>
                      <span>
                        Tạo tệp{' '}
                        <code className="font-mono font-semibold text-slate-900 bg-white border border-[#E6E0D6] px-1.5 py-0.5 rounded">
                          .env
                        </code>{' '}
                        tại thư mục gốc dự án với URL và Anon Key từ Supabase
                        Project Settings:
                      </span>
                    </li>
                  </ol>

                  <pre className="rounded-xl bg-slate-900 text-slate-100 p-3 text-[11px] font-mono overflow-x-auto leading-relaxed">
                    {`VITE_SUPABASE_URL=https://your-project.supabase.co\nVITE_SUPABASE_ANON_KEY=your-anon-public-key`}
                  </pre>

                  <ol
                    start={2}
                    className="space-y-2.5 text-xs text-slate-700 leading-relaxed"
                  >
                    <li className="flex items-start gap-2.5">
                      <span className="h-5 w-5 rounded-full bg-indigo-100 text-[#4F46E5] font-semibold text-[11px] flex items-center justify-center shrink-0 mt-0.5">
                        2
                      </span>
                      <span>
                        Mở <strong>SQL Editor</strong> trong Supabase Dashboard
                        và chạy toàn bộ tệp{' '}
                        <code className="font-mono font-semibold text-slate-900 bg-white border border-[#E6E0D6] px-1.5 py-0.5 rounded">
                          supabase/schema.sql
                        </code>{' '}
                        để tạo bảng <code className="font-mono">documents</code>{' '}
                        và bucket{' '}
                        <code className="font-mono">user-files</code>.
                      </span>
                    </li>
                    <li className="flex items-start gap-2.5">
                      <span className="h-5 w-5 rounded-full bg-indigo-100 text-[#4F46E5] font-semibold text-[11px] flex items-center justify-center shrink-0 mt-0.5">
                        3
                      </span>
                      <span>
                        Khởi động lại ứng dụng để mở khóa Đăng nhập &amp; Đồng bộ
                        đa thiết bị.
                      </span>
                    </li>
                  </ol>
                </div>

                <button
                  type="button"
                  onClick={onClose}
                  className="w-full min-h-[44px] rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-sm font-semibold shadow-sm shadow-indigo-600/20 transition-colors"
                >
                  Đóng &amp; Tiếp tục dùng Local Mode
                </button>
              </div>
            ) : (
              <div className="p-4 sm:p-6 space-y-5">
                {/* Tab Switcher */}
                <div
                  role="tablist"
                  aria-label="Chế độ xác thực"
                  className="grid grid-cols-2 gap-1.5 p-1 rounded-2xl bg-[#FAF8F5] border border-[#EAE5DC]"
                >
                  <button
                    type="button"
                    role="tab"
                    aria-selected={mode === 'login'}
                    disabled={isBusy}
                    onClick={() => handleSwitchMode('login')}
                    className={`min-h-[40px] rounded-xl text-xs font-semibold transition-all flex items-center justify-center gap-1.5 ${
                      mode === 'login'
                        ? 'bg-white text-[#4F46E5] shadow-sm ring-1 ring-[#E6E0D6]'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    <LogIn className="h-3.5 w-3.5" />
                    <span>Đăng nhập</span>
                  </button>
                  <button
                    type="button"
                    role="tab"
                    aria-selected={mode === 'register'}
                    disabled={isBusy}
                    onClick={() => handleSwitchMode('register')}
                    className={`min-h-[40px] rounded-xl text-xs font-semibold transition-all flex items-center justify-center gap-1.5 ${
                      mode === 'register'
                        ? 'bg-white text-[#4F46E5] shadow-sm ring-1 ring-[#E6E0D6]'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    <UserPlus className="h-3.5 w-3.5" />
                    <span>Đăng ký</span>
                  </button>
                </div>

                {/* Alerts */}
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
                    className="rounded-2xl border border-emerald-200 bg-emerald-50 px-3.5 py-3 space-y-2.5 text-xs text-emerald-900"
                  >
                    <div className="flex items-start gap-2.5">
                      <CheckCircle2 className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5" />
                      <span className="leading-relaxed">{successMessage}</span>
                    </div>
                    <button
                      type="button"
                      onClick={() => handleSwitchMode('login')}
                      className="inline-flex items-center gap-1.5 text-xs font-semibold text-emerald-700 hover:text-emerald-900 underline underline-offset-2"
                    >
                      Chuyển sang tab Đăng nhập
                    </button>
                  </div>
                )}

                {/* Google OAuth Button (available in both Login & Register tabs) */}
                <button
                  type="button"
                  onClick={() => void handleGoogleAuth()}
                  disabled={isBusy}
                  className="w-full min-h-[44px] rounded-xl bg-white hover:bg-[#FAF8F5] text-slate-800 border border-[#DDD6CA] shadow-sm text-sm font-semibold flex items-center justify-center gap-2.5 transition-colors disabled:opacity-60 disabled:cursor-not-allowed"
                >
                  {isGoogleLoading ? (
                    <>
                      <Loader2 className="h-4 w-4 animate-spin text-[#4F46E5]" />
                      <span>Đang kết nối Google...</span>
                    </>
                  ) : (
                    <>
                      <svg
                        viewBox="0 0 24 24"
                        className="h-4 w-4 shrink-0"
                        aria-hidden="true"
                      >
                        <path
                          fill="#4285F4"
                          d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.52-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.665-5.17 3.665-9.17z"
                        />
                        <path
                          fill="#34A853"
                          d="M12 24c3.3 0 6.08-1.09 8.105-2.96l-3.88-3.05c-1.1.74-2.51 1.18-4.225 1.18-3.24 0-5.99-2.19-6.97-5.13H1.02v3.14C3.035 21.18 7.21 24 12 24z"
                        />
                        <path
                          fill="#FBBC05"
                          d="M5.03 14.04c-.25-.74-.39-1.53-.39-2.34s.14-1.6.39-2.34V6.22H1.02C.37 7.51 0 8.97 0 10.5s.37 2.99 1.02 4.28l4.01-3.14z"
                        />
                        <path
                          fill="#EA4335"
                          d="M12 4.75c1.8 0 3.41.62 4.68 1.84l3.51-3.51C18.07 1.19 15.3 0 12 0 7.21 0 3.035 2.82 1.02 6.86l4.01 3.14c.98-2.94 3.73-5.25 6.97-5.25z"
                        />
                      </svg>
                      <span>Tiếp tục với Google</span>
                    </>
                  )}
                </button>

                <div className="relative flex items-center justify-center">
                  <div
                    aria-hidden="true"
                    className="absolute inset-x-0 top-1/2 border-t border-[#EAE5DC]"
                  />
                  <span className="relative bg-white px-3 text-[11px] font-medium uppercase tracking-wider text-slate-400">
                    hoặc
                  </span>
                </div>

                {/* Form */}
                <form onSubmit={handleSubmit} className="space-y-4" noValidate>
                  <div className="space-y-1.5">
                    <label
                      htmlFor="auth-email-input"
                      className="block text-xs font-semibold text-slate-700"
                    >
                      Email
                    </label>
                    <div className="relative">
                      <Mail
                        className="h-4 w-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none"
                        aria-hidden="true"
                      />
                      <input
                        id="auth-email-input"
                        type="email"
                        autoComplete="email"
                        value={email}
                        disabled={isBusy}
                        onChange={(event) => {
                          setEmail(event.target.value);
                          if (errorMessage) {
                            setErrorMessage(null);
                          }
                        }}
                        placeholder="ten.ban@example.com"
                        className="w-full h-11 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5]/60 pl-10 pr-3.5 text-sm text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-[#4F46E5] focus:outline-none focus:ring-2 focus:ring-indigo-500/15 disabled:opacity-60"
                      />
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <div className="flex items-center justify-between gap-2">
                      <label
                        htmlFor="auth-password-input"
                        className="block text-xs font-semibold text-slate-700"
                      >
                        Mật khẩu
                      </label>
                      {mode === 'register' && (
                        <span className="text-[11px] text-slate-500">
                          Tối thiểu 6 ký tự
                        </span>
                      )}
                    </div>
                    <div className="relative">
                      <Lock
                        className="h-4 w-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none"
                        aria-hidden="true"
                      />
                      <input
                        id="auth-password-input"
                        type={showPassword ? 'text' : 'password'}
                        autoComplete={
                          mode === 'login' ? 'current-password' : 'new-password'
                        }
                        value={password}
                        disabled={isBusy}
                        onChange={(event) => {
                          setPassword(event.target.value);
                          if (errorMessage) {
                            setErrorMessage(null);
                          }
                        }}
                        placeholder={
                          mode === 'register'
                            ? 'Nhập ít nhất 6 ký tự...'
                            : 'Nhập mật khẩu của bạn...'
                        }
                        className="w-full h-11 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5]/60 pl-10 pr-11 text-sm text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-[#4F46E5] focus:outline-none focus:ring-2 focus:ring-indigo-500/15 disabled:opacity-60"
                      />
                      <button
                        type="button"
                        onClick={() => setShowPassword((prev) => !prev)}
                        disabled={isBusy}
                        aria-label={
                          showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'
                        }
                        className="absolute right-1.5 top-1/2 -translate-y-1/2 h-8 w-8 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 flex items-center justify-center transition-colors"
                      >
                        {showPassword ? (
                          <EyeOff className="h-4 w-4" />
                        ) : (
                          <Eye className="h-4 w-4" />
                        )}
                      </button>
                    </div>
                  </div>

                  <button
                    type="submit"
                    disabled={isBusy}
                    className="w-full min-h-[44px] rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-sm font-semibold shadow-sm shadow-indigo-600/20 flex items-center justify-center gap-2 transition-colors disabled:opacity-60 disabled:cursor-not-allowed"
                  >
                    {isSubmitting ? (
                      <>
                        <Loader2 className="h-4 w-4 animate-spin" />
                        <span>
                          {mode === 'login'
                            ? 'Đang đăng nhập...'
                            : 'Đang tạo tài khoản...'}
                        </span>
                      </>
                    ) : mode === 'login' ? (
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

                <p className="text-[11.5px] text-center text-slate-500 leading-relaxed pt-1">
                  Tài liệu cục bộ trên máy vẫn được giữ nguyên; sau khi đăng
                  nhập bạn có thể chọn tải lên để đồng bộ đa thiết bị.
                </p>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
