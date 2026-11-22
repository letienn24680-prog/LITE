import { isSupabaseConfigured, supabase } from '../lib/supabase';
import {
  setActiveGoogleDriveUser,
  setStoredGoogleDriveToken,
} from './googleDrive';

const GIS_SCRIPT_SRC = 'https://accounts.google.com/gsi/client';

export const GOOGLE_DRIVE_DIRECT_SCOPES =
  'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly';

interface GoogleIdCredentialResponse {
  credential?: string;
  select_by?: string;
  clientId?: string;
}

interface GooglePromptMomentNotification {
  isNotDisplayed: () => boolean;
  getNotDisplayedReason?: () => string;
  isSkippedMoment: () => boolean;
  getSkippedReason?: () => string;
  isDismissedMoment?: () => boolean;
  getDismissedReason?: () => string;
}

interface GoogleTokenResponse {
  access_token?: string;
  error?: string;
  error_description?: string;
  scope?: string;
}

interface GoogleTokenClient {
  requestAccessToken: (overrideConfig?: {
    prompt?: '' | 'none' | 'consent' | 'select_account';
    hint?: string;
    scope?: string;
  }) => void;
}

interface GoogleAccountsNamespace {
  id: {
    initialize: (config: {
      client_id: string;
      callback: (response: GoogleIdCredentialResponse) => void;
      nonce?: string;
      auto_select?: boolean;
      cancel_on_tap_outside?: boolean;
      use_fedcm_for_prompt?: boolean;
      itp_support?: boolean;
    }) => void;
    prompt: (
      momentListener?: (notification: GooglePromptMomentNotification) => void
    ) => void;
    renderButton: (
      parent: HTMLElement,
      options: {
        type?: 'standard' | 'icon';
        theme?: 'outline' | 'filled_blue' | 'filled_black';
        size?: 'large' | 'medium' | 'small';
        text?: 'signin_with' | 'signup_with' | 'continue_with' | 'signin';
        shape?: 'rectangular' | 'pill' | 'circle' | 'square';
        logo_alignment?: 'left' | 'center';
        width?: number;
        locale?: string;
      }
    ) => void;
    cancel: () => void;
  };
  oauth2: {
    initTokenClient: (config: {
      client_id: string;
      scope: string;
      hint?: string;
      callback: (response: GoogleTokenResponse) => void;
      error_callback?: (error: { type?: string; message?: string }) => void;
    }) => GoogleTokenClient;
  };
}

interface GoogleGlobal {
  accounts?: GoogleAccountsNamespace;
}

declare global {
  interface Window {
    google?: GoogleGlobal;
  }
}

export interface GoogleDirectSignInResult {
  status: 'success' | 'prompt_suppressed';
  userId?: string;
  email?: string | null;
  driveConnected?: boolean;
  driveError?: string;
  suppressedReason?: string;
}

let scriptLoadPromise: Promise<GoogleAccountsNamespace> | null = null;

export function getGoogleClientId(): string {
  return (import.meta.env.VITE_GOOGLE_CLIENT_ID ?? '').trim();
}

export function isGoogleDirectAuthConfigured(): boolean {
  const clientId = getGoogleClientId();
  return (
    clientId.length > 0 &&
    !clientId.toLowerCase().includes('your-google-client-id') &&
    !/^<.*>$/.test(clientId)
  );
}

/**
 * Loads `https://accounts.google.com/gsi/client` dynamically if `window.google?.accounts` is absent.
 */
export function loadGoogleIdentityScript(): Promise<GoogleAccountsNamespace> {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return Promise.reject(
      new Error('Google Identity Services chỉ hoạt động trên trình duyệt.')
    );
  }

  if (window.google?.accounts?.id && window.google?.accounts?.oauth2) {
    return Promise.resolve(window.google.accounts);
  }

  if (scriptLoadPromise) {
    return scriptLoadPromise;
  }

  scriptLoadPromise = new Promise<GoogleAccountsNamespace>((resolve, reject) => {
    const resolveIfReady = () => {
      if (window.google?.accounts?.id && window.google?.accounts?.oauth2) {
        resolve(window.google.accounts);
        return true;
      }
      return false;
    };

    const existingScript = document.querySelector<HTMLScriptElement>(
      `script[src="${GIS_SCRIPT_SRC}"]`
    );
    if (existingScript) {
      if (resolveIfReady()) {
        return;
      }
      existingScript.addEventListener(
        'load',
        () => {
          if (!resolveIfReady()) {
            scriptLoadPromise = null;
            reject(new Error('Không thể khởi tạo Google Identity Services.'));
          }
        },
        { once: true }
      );
      existingScript.addEventListener(
        'error',
        () => {
          scriptLoadPromise = null;
          reject(
            new Error(
              'Không thể tải script đăng nhập Google (https://accounts.google.com/gsi/client).'
            )
          );
        },
        { once: true }
      );
      return;
    }

    const script = document.createElement('script');
    script.src = GIS_SCRIPT_SRC;
    script.async = true;
    script.defer = true;
    script.onload = () => {
      if (!resolveIfReady()) {
        scriptLoadPromise = null;
        reject(new Error('Không thể khởi tạo Google Identity Services.'));
      }
    };
    script.onerror = () => {
      scriptLoadPromise = null;
      reject(
        new Error(
          'Không thể tải script Google Identity Services. Vui lòng kiểm tra kết nối mạng.'
        )
      );
    };
    document.head.appendChild(script);
  });

  return scriptLoadPromise;
}

async function createNoncePair(): Promise<{
  rawNonce?: string;
  hashedNonce?: string;
}> {
  try {
    if (
      typeof window === 'undefined' ||
      !window.crypto?.getRandomValues ||
      !window.crypto?.subtle
    ) {
      return {};
    }
    const randomBytes = new Uint8Array(16);
    window.crypto.getRandomValues(randomBytes);
    const rawNonce = Array.from(randomBytes, (b) =>
      b.toString(16).padStart(2, '0')
    ).join('');
    const encoded = new TextEncoder().encode(rawNonce);
    const digest = await window.crypto.subtle.digest('SHA-256', encoded);
    const hashedNonce = Array.from(new Uint8Array(digest), (b) =>
      b.toString(16).padStart(2, '0')
    ).join('');
    return { rawNonce, hashedNonce };
  } catch {
    return {};
  }
}

/**
 * Requests a Google Drive OAuth2 access token directly from Google via GIS TokenClient
 * (no intermediate `supabase.co` redirect) and saves it scoped to `userId`.
 */
export async function requestGoogleDriveAccessTokenDirect(options?: {
  userId?: string | null;
  loginHint?: string | null;
  scopes?: string;
  prompt?: '' | 'none' | 'consent' | 'select_account';
}): Promise<string> {
  const clientId = getGoogleClientId();
  if (!isGoogleDirectAuthConfigured()) {
    throw new Error(
      'Chưa cấu hình VITE_GOOGLE_CLIENT_ID trong tệp .env để kết nối Google Drive trực tiếp.'
    );
  }

  const accounts = await loadGoogleIdentityScript();
  const targetScopes = options?.scopes?.trim() || GOOGLE_DRIVE_DIRECT_SCOPES;
  const loginHint = options?.loginHint?.trim() || undefined;

  const requestOnce = (
    promptMode: '' | 'none' | 'consent' | 'select_account'
  ): Promise<string> =>
    new Promise<string>((resolve, reject) => {
      let settled = false;
      const client = accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: targetScopes,
        ...(loginHint ? { hint: loginHint } : {}),
        callback: (tokenResp) => {
          if (settled) return;
          settled = true;
          if (tokenResp.error || !tokenResp.access_token) {
            reject(
              new Error(
                tokenResp.error_description ||
                  tokenResp.error ||
                  'Không nhận được quyền truy cập Google Drive từ Google.'
              )
            );
            return;
          }
          resolve(tokenResp.access_token);
        },
        error_callback: (err) => {
          if (settled) return;
          settled = true;
          reject(
            new Error(
              err.message ||
                'Trình duyệt đã chặn cửa sổ cấp quyền Google Drive hoặc bạn đã đóng hộp thoại.'
            )
          );
        },
      });

      client.requestAccessToken({
        prompt: promptMode,
        ...(loginHint ? { hint: loginHint } : {}),
      });
    });

  let accessToken: string;
  try {
    accessToken = await requestOnce(options?.prompt ?? '');
  } catch (firstError) {
    if (options?.prompt === undefined) {
      accessToken = await requestOnce('consent');
    } else {
      throw firstError;
    }
  }

  if (options?.userId) {
    setStoredGoogleDriveToken(accessToken, options.userId);
  } else {
    setStoredGoogleDriveToken(accessToken);
  }

  return accessToken;
}

async function exchangeIdTokenAndLinkDrive(
  credential: string,
  rawNonce?: string,
  scopes: string = GOOGLE_DRIVE_DIRECT_SCOPES
): Promise<GoogleDirectSignInResult> {
  if (!isSupabaseConfigured || !supabase) {
    throw new Error('Ứng dụng đang ở Chế độ cục bộ, chưa cấu hình Supabase.');
  }

  const { data, error } = await supabase.auth.signInWithIdToken({
    provider: 'google',
    token: credential,
    ...(rawNonce ? { nonce: rawNonce } : {}),
  });

  if (error) {
    throw new Error(error.message);
  }

  const session = data.session;
  const user = data.user ?? session?.user;
  if (!user?.id) {
    throw new Error('Không thể khởi tạo phiên đăng nhập Google trên Supabase.');
  }

  setActiveGoogleDriveUser(user.id);

  let driveConnected = false;
  let driveError: string | undefined;

  try {
    await requestGoogleDriveAccessTokenDirect({
      userId: user.id,
      loginHint: user.email ?? undefined,
      scopes,
      prompt: 'consent',
    });
    driveConnected = true;
  } catch (err) {
    driveConnected = false;
    driveError =
      err instanceof Error
        ? err.message
        : 'Chưa thể tự động liên kết Google Drive trong bước này.';
  }

  return {
    status: 'success',
    userId: user.id,
    email: user.email ?? null,
    driveConnected,
    driveError,
  };
}

/**
 * Direct Web sign-in using Google Identity Services (`google.accounts.id` + Supabase `signInWithIdToken`)
 * followed by requesting Drive OAuth2 scopes directly from Google for the authenticated `userId`.
 */
export async function signInWithGoogleDirect(
  scopes: string = GOOGLE_DRIVE_DIRECT_SCOPES
): Promise<GoogleDirectSignInResult> {
  const clientId = getGoogleClientId();
  if (!isGoogleDirectAuthConfigured()) {
    throw new Error(
      'Chưa cấu hình biến VITE_GOOGLE_CLIENT_ID trong .env. Vui lòng điền Google OAuth Client ID (Web application) từ Google Cloud Console.'
    );
  }
  if (!isSupabaseConfigured || !supabase) {
    throw new Error('Ứng dụng đang ở Chế độ cục bộ, chưa cấu hình Supabase.');
  }

  const accounts = await loadGoogleIdentityScript();
  const { rawNonce, hashedNonce } = await createNoncePair();

  return new Promise<GoogleDirectSignInResult>((resolve, reject) => {
    let settled = false;

    const timeoutId = window.setTimeout(() => {
      if (settled) return;
      settled = true;
      resolve({
        status: 'prompt_suppressed',
        suppressedReason: 'Hết thời gian chờ hộp thoại Google One Tap.',
      });
    }, 15000);

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(timeoutId);
      fn();
    };

    accounts.id.initialize({
      client_id: clientId,
      auto_select: false,
      cancel_on_tap_outside: true,
      use_fedcm_for_prompt: true,
      itp_support: true,
      ...(hashedNonce ? { nonce: hashedNonce } : {}),
      callback: (response) => {
        if (!response.credential) {
          finish(() =>
            reject(new Error('Không nhận được mã xác thực id_token từ Google.'))
          );
          return;
        }
        const credential = response.credential;
        finish(() => {
          void exchangeIdTokenAndLinkDrive(credential, rawNonce, scopes)
            .then(resolve)
            .catch(reject);
        });
      },
    });

    accounts.id.prompt((notification) => {
      if (settled) return;
      if (notification.isNotDisplayed() || notification.isSkippedMoment()) {
        const reason =
          notification.getNotDisplayedReason?.() ||
          notification.getSkippedReason?.() ||
          'One Tap không khả dụng trên trình duyệt hiện tại';
        finish(() =>
          resolve({
            status: 'prompt_suppressed',
            suppressedReason: reason,
          })
        );
      } else if (notification.isDismissedMoment?.()) {
        const reason = notification.getDismissedReason?.() || '';
        if (reason !== 'credential_returned') {
          finish(() =>
            resolve({
              status: 'prompt_suppressed',
              suppressedReason: reason || 'Hộp thoại chọn tài khoản Google đã đóng.',
            })
          );
        }
      }
    });
  });
}

/**
 * Renders the official Google Identity Services Sign-In button inside `container` as a reliable
 * direct fallback on Web when Google One Tap `prompt()` is suppressed/cooled down.
 */
export async function renderGoogleDirectSignInButton(
  container: HTMLElement,
  callbacks: {
    onStart?: () => void;
    onSuccess: (result: GoogleDirectSignInResult) => void;
    onError: (error: Error) => void;
    scopes?: string;
  }
): Promise<void> {
  const clientId = getGoogleClientId();
  if (!isGoogleDirectAuthConfigured()) {
    throw new Error(
      'Chưa cấu hình biến VITE_GOOGLE_CLIENT_ID trong .env.'
    );
  }

  const accounts = await loadGoogleIdentityScript();
  const { rawNonce, hashedNonce } = await createNoncePair();

  accounts.id.initialize({
    client_id: clientId,
    auto_select: false,
    cancel_on_tap_outside: true,
    use_fedcm_for_prompt: true,
    itp_support: true,
    ...(hashedNonce ? { nonce: hashedNonce } : {}),
    callback: (response) => {
      if (!response.credential) {
        callbacks.onError(
          new Error('Không nhận được mã xác thực id_token từ Google.')
        );
        return;
      }
      callbacks.onStart?.();
      void exchangeIdTokenAndLinkDrive(
        response.credential,
        rawNonce,
        callbacks.scopes ?? GOOGLE_DRIVE_DIRECT_SCOPES
      )
        .then(callbacks.onSuccess)
        .catch((err) => {
          callbacks.onError(
            err instanceof Error ? err : new Error('Đăng nhập Google thất bại.')
          );
        });
    },
  });

  container.innerHTML = '';
  const targetWidth = Math.max(
    240,
    Math.min(360, container.clientWidth || 320)
  );
  accounts.id.renderButton(container, {
    type: 'standard',
    theme: 'outline',
    size: 'large',
    text: 'continue_with',
    shape: 'pill',
    logo_alignment: 'left',
    width: targetWidth,
    locale: 'vi',
  });
}
