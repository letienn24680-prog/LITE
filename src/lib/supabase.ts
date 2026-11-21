import { createClient, type SupabaseClient } from '@supabase/supabase-js';

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
