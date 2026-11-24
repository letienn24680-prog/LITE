import { App as CapacitorApp } from '@capacitor/app';
import { Browser } from '@capacitor/browser';
import { Capacitor } from '@capacitor/core';

export interface AppVersionManifest {
  version: string;
  buildNumber: number;
  downloadUrl: string;
  changelog: string;
}

export interface AppUpdateCheckResult {
  hasUpdate: boolean;
  currentVersion: string;
  currentBuildNumber: number;
  latest: AppVersionManifest | null;
}

const FALLBACK_CURRENT_VERSION = '1.0';
const FALLBACK_CURRENT_BUILD_NUMBER = 1;
const REMOTE_VERSION_MANIFEST_URL =
  (import.meta.env.VITE_APP_UPDATE_URL ?? '').trim() ||
  'https://raw.githubusercontent.com/letienn24680-prog/LITE/main/public/version.json';

function parseBuildNumber(raw: unknown): number {
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    return Math.floor(raw);
  }
  if (typeof raw === 'string' && raw.trim()) {
    const num = Number.parseInt(raw.trim(), 10);
    if (Number.isFinite(num)) {
      return num;
    }
  }
  return FALLBACK_CURRENT_BUILD_NUMBER;
}

function isValidManifest(value: unknown): value is AppVersionManifest {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const obj = value as Record<string, unknown>;
  return (
    typeof obj.version === 'string' &&
    obj.version.trim().length > 0 &&
    typeof obj.buildNumber === 'number' &&
    Number.isFinite(obj.buildNumber) &&
    typeof obj.downloadUrl === 'string' &&
    obj.downloadUrl.trim().length > 0 &&
    typeof obj.changelog === 'string'
  );
}

async function fetchManifestFromUrl(
  url: string,
  timeoutMs = 5000
): Promise<AppVersionManifest | null> {
  const controller =
    typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timeoutId =
    controller !== null
      ? window.setTimeout(() => controller.abort(), timeoutMs)
      : null;

  try {
    const separator = url.includes('?') ? '&' : '?';
    const res = await fetch(`${url}${separator}t=${Date.now()}`, {
      cache: 'no-store',
      signal: controller?.signal,
    });
    if (!res.ok) {
      return null;
    }
    const json: unknown = await res.json();
    return isValidManifest(json) ? json : null;
  } catch {
    return null;
  } finally {
    if (timeoutId !== null) {
      window.clearTimeout(timeoutId);
    }
  }
}

export async function getCurrentNativeAppInfo(): Promise<{
  version: string;
  buildNumber: number;
}> {
  if (!Capacitor.isNativePlatform()) {
    return {
      version: FALLBACK_CURRENT_VERSION,
      buildNumber: FALLBACK_CURRENT_BUILD_NUMBER,
    };
  }
  try {
    const info = await CapacitorApp.getInfo();
    return {
      version: info.version?.trim() || FALLBACK_CURRENT_VERSION,
      buildNumber: parseBuildNumber(info.build),
    };
  } catch {
    return {
      version: FALLBACK_CURRENT_VERSION,
      buildNumber: FALLBACK_CURRENT_BUILD_NUMBER,
    };
  }
}

export async function checkForAndroidAppUpdate(options?: {
  force?: boolean;
}): Promise<AppUpdateCheckResult> {
  const isNative = Capacitor.isNativePlatform();
  const { version: currentVersion, buildNumber: currentBuildNumber } =
    await getCurrentNativeAppInfo();

  if (!isNative && !options?.force) {
    return {
      hasUpdate: false,
      currentVersion,
      currentBuildNumber,
      latest: null,
    };
  }

  let manifest = await fetchManifestFromUrl(REMOTE_VERSION_MANIFEST_URL, 4500);
  if (!manifest) {
    manifest = await fetchManifestFromUrl('./version.json', 3000);
  }

  if (!manifest) {
    return {
      hasUpdate: false,
      currentVersion,
      currentBuildNumber,
      latest: null,
    };
  }

  return {
    hasUpdate: manifest.buildNumber > currentBuildNumber,
    currentVersion,
    currentBuildNumber,
    latest: manifest,
  };
}

export async function launchApkUpdateDownload(downloadUrl: string): Promise<void> {
  const cleanUrl = downloadUrl.trim();
  if (!cleanUrl) {
    return;
  }

  if (Capacitor.isNativePlatform()) {
    try {
      await Browser.open({ url: cleanUrl });
      return;
    } catch {
      // Fallback to direct navigation below
    }
  }

  window.open(cleanUrl, '_blank', 'noopener,noreferrer');
}
