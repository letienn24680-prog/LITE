import { Capacitor } from '@capacitor/core';
import { Directory, Filesystem } from '@capacitor/filesystem';
import { Share } from '@capacitor/share';

export interface SaveFileOptions {
  mimeType?: string;
  shareTitle?: string;
  /** When running inside Electron, automatically opens the containing folder after saveFile succeeds */
  openFolderAfterSave?: boolean;
  /** Optional callback triggered when the saved file folder-open action is executed */
  onShowInFolder?: (savedPath: string) => void | Promise<void>;
}

export interface SaveFileAction {
  label: string;
  run: () => Promise<void>;
}

export interface SaveFileResult {
  method: 'android' | 'electron' | 'web';
  savedPath?: string;
  message: string;
  /** Supported on Electron when window.liteDesktop?.showItemInFolder is available */
  canShowInFolder?: boolean;
  /** Opens the folder containing the saved file on Electron */
  showInFolder?: () => Promise<void>;
  /** Action descriptor for UI notifications (e.g., "Mở thư mục") */
  action?: SaveFileAction;
}

interface ElectronSaveBridge {
  isDesktop?: boolean;
  saveFile?: (payload: {
    fileName: string;
    base64: string;
    mimeType?: string;
  }) => Promise<{ canceled?: boolean; filePath?: string }>;
  showItemInFolder?: (filePath: string) => Promise<void>;
}

function sanitizeFileName(rawName: string): string {
  const trimmed = rawName.trim().replace(/[/\\:*?"<>|\x00-\x1F]+/g, '_');
  return trimmed.length > 0 ? trimmed : 'document';
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  const parts: string[] = [];
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const slice = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
    let chunk = '';
    for (let j = 0; j < slice.length; j += 1) {
      chunk += String.fromCharCode(slice[j]);
    }
    parts.push(chunk);
  }
  return btoa(parts.join(''));
}

async function blobToBase64(blob: Blob): Promise<string> {
  if (typeof FileReader !== 'undefined') {
    try {
      return await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          const result = reader.result;
          if (typeof result === 'string') {
            const commaIdx = result.indexOf(',');
            resolve(commaIdx >= 0 ? result.slice(commaIdx + 1) : result);
          } else {
            reject(new Error('Không thể chuyển đổi dữ liệu tệp.'));
          }
        };
        reader.onerror = () => {
          reject(reader.error ?? new Error('Không thể đọc dữ liệu tệp.'));
        };
        reader.readAsDataURL(blob);
      });
    } catch {
      // Fallback to ArrayBuffer conversion
    }
  }
  const buffer = await blob.arrayBuffer();
  return arrayBufferToBase64(buffer);
}

function getElectronBridge(): ElectronSaveBridge | undefined {
  if (typeof window === 'undefined') {
    return undefined;
  }
  return (window as unknown as { liteDesktop?: ElectronSaveBridge }).liteDesktop;
}

/**
 * Opens the folder containing a saved file on Electron Desktop if available.
 */
export async function openSavedFileFolder(filePath: string): Promise<boolean> {
  const trimmed = filePath.trim();
  const bridge = getElectronBridge();
  if (!trimmed || !bridge?.showItemInFolder) {
    return false;
  }
  await bridge.showItemInFolder(trimmed);
  return true;
}

export async function saveFileUniversal(
  blob: Blob,
  fileName: string,
  options?: SaveFileOptions
): Promise<SaveFileResult> {
  const safeName = sanitizeFileName(fileName);
  const mimeType =
    options?.mimeType || blob.type || 'application/octet-stream';

  // 1. Capacitor Android / Native Platform
  if (Capacitor.isNativePlatform()) {
    const base64 = await blobToBase64(blob);
    const relativePath = `LITE/${safeName}`;
    let savedDirectory = Directory.Documents;
    let displayLocation = `Documents/LITE/${safeName}`;

    try {
      await Filesystem.writeFile({
        path: relativePath,
        data: base64,
        directory: Directory.Documents,
        recursive: true,
      });
    } catch {
      savedDirectory = Directory.Cache;
      displayLocation = `Cache/LITE/${safeName}`;
      await Filesystem.writeFile({
        path: relativePath,
        data: base64,
        directory: Directory.Cache,
        recursive: true,
      });
    }

    let fileUri = '';
    try {
      const uriResult = await Filesystem.getUri({
        directory: savedDirectory,
        path: relativePath,
      });
      fileUri = uriResult.uri;
    } catch {
      fileUri = '';
    }

    let shareSucceeded = false;
    if (fileUri) {
      try {
        await Share.share({
          title: options?.shareTitle || safeName,
          url: fileUri,
          dialogTitle: 'Mở hoặc chia sẻ tệp',
        });
        shareSucceeded = true;
      } catch {
        // User dismissed Share sheet; file is already persisted at displayLocation
      }
    }

    const storageLabel =
      savedDirectory === Directory.Documents
        ? 'Directory.Documents'
        : 'Directory.Cache';
    const shareStatusSuffix = shareSucceeded
      ? ' • Đã mở bảng Chia sẻ (Share.share).'
      : fileUri
        ? ' • Đã sẵn sàng chia sẻ qua Share.share.'
        : '.';

    return {
      method: 'android',
      savedPath: displayLocation,
      message: `Đã lưu tệp tại ${displayLocation} (${storageLabel})${shareStatusSuffix}`,
    };
  }

  // 2. Electron Desktop IPC
  const electronBridge = getElectronBridge();
  if (electronBridge?.saveFile) {
    const base64 = await blobToBase64(blob);
    const result = await electronBridge.saveFile({
      fileName: safeName,
      base64,
      mimeType,
    });

    if (result?.canceled) {
      return {
        method: 'electron',
        message: 'Đã hủy thao tác lưu tệp.',
      };
    }

    const rawSavedFilePath = result?.filePath?.trim();
    const chosenPath = rawSavedFilePath || safeName;
    const hasShowInFolderBridge =
      Boolean(rawSavedFilePath) && Boolean(electronBridge.showItemInFolder);

    if (hasShowInFolderBridge && rawSavedFilePath && electronBridge.showItemInFolder) {
      const showInFolderFn = electronBridge.showItemInFolder;
      const runShowInFolder = async (): Promise<void> => {
        await showInFolderFn(rawSavedFilePath);
        if (options?.onShowInFolder) {
          await options.onShowInFolder(rawSavedFilePath);
        }
      };

      if (options?.openFolderAfterSave) {
        try {
          await runShowInFolder();
        } catch {
          // Opening folder failed should not invalidate a successful file save
        }
      }

      return {
        method: 'electron',
        savedPath: chosenPath,
        message: `Đã lưu tệp tại ${chosenPath}`,
        canShowInFolder: true,
        showInFolder: runShowInFolder,
        action: {
          label: 'Mở thư mục',
          run: runShowInFolder,
        },
      };
    }

    return {
      method: 'electron',
      savedPath: chosenPath,
      message: `Đã lưu tệp tại ${chosenPath}`,
    };
  }

  // 3. Web Browser (retain object URL for 60 seconds)
  const normalizedBlob =
    blob.type === mimeType ? blob : new Blob([blob], { type: mimeType });
  const objectUrl = URL.createObjectURL(normalizedBlob);

  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = safeName;
  anchor.rel = 'noopener noreferrer';
  anchor.style.display = 'none';

  const supportsDownloadAttr = typeof anchor.download !== 'undefined';
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);

  if (!supportsDownloadAttr) {
    window.open(objectUrl, '_blank', 'noopener,noreferrer');
  }

  window.setTimeout(() => {
    URL.revokeObjectURL(objectUrl);
  }, 60_000);

  return {
    method: 'web',
    savedPath: safeName,
    message: `Đã tải xuống "${safeName}".`,
  };
}
