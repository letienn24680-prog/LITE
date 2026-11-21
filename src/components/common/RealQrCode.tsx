import { useEffect, useMemo, useState } from 'react';
import QRCode from 'qrcode';

export interface RealQrCodeProps {
  value: string;
  sizeClassName?: string;
  darkHex?: string;
  lightHex?: string;
  ariaLabel?: string;
}

export function RealQrCode({
  value,
  sizeClassName = 'h-44 w-44 sm:h-48 sm:w-48',
  darkHex = '#0F172A',
  lightHex = '#FFFFFF',
  ariaLabel = 'Mã QR',
}: RealQrCodeProps) {
  const [svgMarkup, setSvgMarkup] = useState<string | null>(null);
  const [hasError, setHasError] = useState<boolean>(false);

  const trimmedValue = value.trim();

  useEffect(() => {
    if (!trimmedValue) {
      setSvgMarkup(null);
      setHasError(true);
      return;
    }

    let cancelled = false;
    setHasError(false);

    QRCode.toString(trimmedValue, {
      type: 'svg',
      margin: 1,
      errorCorrectionLevel: 'M',
      color: {
        dark: darkHex,
        light: lightHex,
      },
    })
      .then((svg) => {
        if (!cancelled) {
          setSvgMarkup(svg);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setSvgMarkup(null);
          setHasError(true);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [trimmedValue, darkHex, lightHex]);

  const svgDataUri = useMemo(() => {
    if (!svgMarkup) {
      return null;
    }
    return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgMarkup)}`;
  }, [svgMarkup]);

  if (hasError) {
    return (
      <div
        role="img"
        aria-label={ariaLabel}
        className={`${sizeClassName} rounded-xl border border-slate-200 bg-slate-100 flex items-center justify-center p-3 text-center text-[11px] font-medium text-slate-500`}
      >
        Không thể tạo mã QR
      </div>
    );
  }

  if (!svgDataUri) {
    return (
      <div
        role="img"
        aria-label={ariaLabel}
        className={`${sizeClassName} rounded-xl bg-slate-100 animate-pulse`}
      />
    );
  }

  return (
    <div
      className={`${sizeClassName} select-none overflow-hidden rounded-lg`}
      style={{ backgroundColor: lightHex }}
    >
      <img
        src={svgDataUri}
        alt={ariaLabel}
        draggable={false}
        className="w-full h-full block object-contain"
      />
    </div>
  );
}
