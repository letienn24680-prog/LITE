import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import { Capacitor } from '@capacitor/core';
import { Browser } from '@capacitor/browser';

const rawSupabaseUrl = (import.meta.env.VITE_SUPABASE_URL ?? '').trim();
const rawSupabaseAnonKey = (import.meta.env.VITE_SUPABASE_ANON_KEY ?? '').trim();

const PLACEHOLDER_TOKENS = [
  'your-project-id',
  'your-anon-public-key',
  'your-supabase-url',
  'your-anon-key',
];

function containsPlaceholder(value: string): boolean {
  const lower = value.toLowerCase();
  return PLACEHOLDER_TOKENS.some((token) => lower.includes(token));
}

function isValidHttpUrl(value: string): boolean {
  if (!/^https?:\/\/\S+$/i.test(value) || containsPlaceholder(value)) {
    return false;
  }
  try {
    const parsed = new URL(value);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:';
  } catch {
    return false;
  }
}

function isValidAnonKey(value: string): boolean {
  if (value.length === 0 || containsPlaceholder(value)) {
    return false;
  }
  return !/^<.*>$/.test(value);
}

let resolvedConfigured = isValidHttpUrl(rawSupabaseUrl) && isValidAnonKey(rawSupabaseAnonKey);
let resolvedClient: SupabaseClient | null = null;
let resolvedUrlPreview: string | null = null;

if (resolvedConfigured) {
  try {
    resolvedClient = createClient(rawSupabaseUrl, rawSupabaseAnonKey, {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
      },
    });
    resolvedUrlPreview = new URL(rawSupabaseUrl).origin;
  } catch {
    resolvedClient = null;
    resolvedConfigured = false;
    resolvedUrlPreview = null;
  }
}

/** True when valid VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY are configured. */
export const isSupabaseConfigured: boolean = resolvedConfigured && resolvedClient !== null;

/** Singleton Supabase browser client, or null when running in Local Mode. */
export const supabase: SupabaseClient | null = resolvedClient;

/**
 * Returns non-sensitive Supabase configuration state and URL preview for UI diagnostics.
 */
export function getSupabaseConfigStatus(): {
  configured: boolean;
  urlPreview: string | null;
} {
  return {
    configured: isSupabaseConfigured,
    urlPreview: resolvedUrlPreview,
  };
}

export const GOOGLE_DRIVE_DEFAULT_SCOPES =
  'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly';

export interface GoogleOAuthFlowResult {
  isNativeApp: boolean;
  openedExternally: boolean;
  error: Error | null;
}

interface LiteDesktopBridge {
  isDesktop?: boolean;
  openExternal?: (url: string) => Promise<void>;
}

/**
 * Checks whether LITE is running inside a packaged native shell (Capacitor Android or Electron).
 */
export function isNativeAppEnvironment(): {
  isNativeApp: boolean;
  isDesktopApp: boolean;
  isMobileNative: boolean;
  desktopBridge?: LiteDesktopBridge;
} {
  const desktopBridge =
    typeof window !== 'undefined'
      ? (window as unknown as { liteDesktop?: LiteDesktopBridge }).liteDesktop
      : undefined;
  const isDesktopApp = Boolean(desktopBridge?.isDesktop);
  const isMobileNative = Capacitor.isNativePlatform();
  return {
    isNativeApp: isDesktopApp || isMobileNative,
    isDesktopApp,
    isMobileNative,
    desktopBridge,
  };
}

/**
 * Starts the unified Google OAuth flow across Web, Electron Desktop, and Capacitor Android.
 * Packaged native shells redirect back to `liteapp://auth/callback` via external/in-app browser.
 */
export async function startGoogleOAuthFlow(
  scopes: string = GOOGLE_DRIVE_DEFAULT_SCOPES
): Promise<GoogleOAuthFlowResult> {
  if (!isSupabaseConfigured || !supabase) {
    return {
      isNativeApp: false,
      openedExternally: false,
      error: new Error('Ứng dụng đang ở Chế độ cục bộ, chưa cấu hình Supabase.'),
    };
  }

  const { isNativeApp, isDesktopApp, isMobileNative, desktopBridge } =
    isNativeAppEnvironment();
  const redirectTo = isNativeApp
    ? 'liteapp://auth/callback'
    : window.location.origin;

  const { data, error } = await supabase.auth.signInWithOAuth({
    provider: 'google',
    options: {
      redirectTo,
      skipBrowserRedirect: isNativeApp,
      scopes,
      queryParams: {
        access_type: 'offline',
        prompt: 'consent',
      },
    },
  });

  if (error) {
    return {
      isNativeApp,
      openedExternally: false,
      error: new Error(error.message),
    };
  }

  if (isNativeApp && data?.url) {
    if (isDesktopApp && desktopBridge?.openExternal) {
      await desktopBridge.openExternal(data.url);
      return { isNativeApp: true, openedExternally: true, error: null };
    }
    if (isMobileNative) {
      await Browser.open({ url: data.url });
      return { isNativeApp: true, openedExternally: true, error: null };
    }
  }

  return {
    isNativeApp,
    openedExternally: false,
    error: null,
  };
}

/**
 * Completes a native OAuth redirect (`liteapp://auth/callback`) by restoring the Supabase
 * session and returning any Google `provider_token` included in the callback or exchange result.
 */
export async function completeNativeOAuthCallback(
  callbackUrl: string
): Promise<{ providerToken: string | null; userId: string | null }> {
  if (!supabase) {
    return { providerToken: null, userId: null };
  }

  const parsed = new URL(callbackUrl);
  const hashParams = new URLSearchParams(
    parsed.hash.startsWith('#') ? parsed.hash.slice(1) : parsed.hash
  );
  const accessToken = hashParams.get('access_token');
  const refreshToken = hashParams.get('refresh_token');
  const hashProviderToken = hashParams.get('provider_token');
  const authCode = parsed.searchParams.get('code');

  if (accessToken && refreshToken) {
    const { data } = await supabase.auth.setSession({
      access_token: accessToken,
      refresh_token: refreshToken,
    });
    return {
      providerToken: data.session?.provider_token ?? hashProviderToken ?? null,
      userId: data.session?.user?.id ?? null,
    };
  }

  if (authCode) {
    const { data } = await supabase.auth.exchangeCodeForSession(authCode);
    return {
      providerToken: data.session?.provider_token ?? hashProviderToken ?? null,
      userId: data.session?.user?.id ?? null,
    };
  }

  return { providerToken: hashProviderToken ?? null, userId: null };
}
