interface LITEAndroidBridge {
  setStatusBarTheme?: (colorHex: string, isDarkBackground: boolean) => void;
}

declare global {
  interface Window {
    LITEAndroidSystem?: LITEAndroidBridge;
  }
}

export function syncAndroidStatusBar(options: {
  color: string;
  isDarkBackground: boolean;
}): void {
  if (typeof document === 'undefined') {
    return;
  }

  let meta = document.querySelector<HTMLMetaElement>('meta[name="theme-color"]');
  if (!meta) {
    meta = document.createElement('meta');
    meta.name = 'theme-color';
    document.head.appendChild(meta);
  }
  meta.content = options.color;

  try {
    window.LITEAndroidSystem?.setStatusBarTheme?.(
      options.color,
      options.isDarkBackground
    );
  } catch {
    // Ignore native bridge call outside Android WebView
  }
}
