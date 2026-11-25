import { App as CapacitorApp } from '@capacitor/app';
import { Browser } from '@capacitor/browser';
import { Capacitor } from '@capacitor/core';
import {
  APP_BUILD_NUMBER,
  APP_VERSION,
  PUBLIC_VERSION_MANIFEST_URL,
} from '../config/appVersion';

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

export const CURRENT_APP_VERSION = APP_VERSION;
export const CURRENT_BUILD_NUMBER = APP_BUILD_NUMBER;

const ENV_UPDATE_URL = (import.meta.env.VITE_APP_UPDATE_URL ?? '').trim();
const REMOTE_VERSION_MANIFEST_URL = ENV_UPDATE_URL.startsWith('https://')
  ? ENV_UPDATE_URL
  : PUBLIC_VERSION_MANIFEST_URL;

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
  return APP_BUILD_NUMBER;
}

function isValidManifest(value: unknown): value is AppVersionManifest {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const obj = value as Record<string, unknown>;
  const normalizedBuild = Number(obj.buildNumber);
  return (
    typeof obj.version === 'string' &&
    obj.version.trim().length > 0 &&
    Number.isFinite(normalizedBuild) &&
    normalizedBuild > 0 &&
    typeof obj.downloadUrl === 'string' &&
    obj.downloadUrl.trim().length > 0 &&
    typeof obj.changelog === 'string'
  );
}

async function fetchManifestFromUrl(
  url: string,
  timeoutMs = 8000
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
    if (!isValidManifest(json)) {
      return null;
    }
    return {
      version: json.version.trim(),
      buildNumber: Number(json.buildNumber),
      downloadUrl: json.downloadUrl.trim(),
      changelog: json.changelog,
    };
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
      version: APP_VERSION,
      buildNumber: APP_BUILD_NUMBER,
    };
  }
  try {
    const info = await CapacitorApp.getInfo();
    return {
      version: info.version?.trim() || APP_VERSION,
      buildNumber: parseBuildNumber(info.build),
    };
  } catch {
    return {
      version: APP_VERSION,
      buildNumber: APP_BUILD_NUMBER,
    };
  }
}

export async function checkForAndroidAppUpdate(options?: {
  force?: boolean;
}): Promise<AppUpdateCheckResult> {
  const isNative = Capacitor.isNativePlatform();

  if (!isNative && !options?.force) {
    return {
      hasUpdate: false,
      currentVersion: APP_VERSION,
      currentBuildNumber: APP_BUILD_NUMBER,
      latest: null,
    };
  }

  const manifest = await fetchManifestFromUrl(REMOTE_VERSION_MANIFEST_URL, 8000);
  if (!manifest) {
    return {
      hasUpdate: false,
      currentVersion: APP_VERSION,
      currentBuildNumber: APP_BUILD_NUMBER,
      latest: null,
    };
  }

  return {
    hasUpdate: Number(manifest.buildNumber) > APP_BUILD_NUMBER,
    currentVersion: APP_VERSION,
    currentBuildNumber: APP_BUILD_NUMBER,
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
