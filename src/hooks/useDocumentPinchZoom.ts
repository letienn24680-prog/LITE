import { useEffect, useRef } from 'react';
import type { RefObject } from 'react';

export interface UseDocumentPinchZoomOptions {
  min?: number;
  max?: number;
  disabled?: boolean;
}

function clampZoom(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, Math.round(value)));
}

function getTouchDistance(t1: Touch, t2: Touch): number {
  const dx = t1.clientX - t2.clientX;
  const dy = t1.clientY - t2.clientY;
  return Math.hypot(dx, dy);
}

export function useDocumentPinchZoom(
  viewportRef: RefObject<HTMLElement | null>,
  zoomPercent: number,
  onZoomChange: (nextZoom: number) => void,
  options?: UseDocumentPinchZoomOptions
): void {
  const minZoom = options?.min ?? 25;
  const maxZoom = options?.max ?? 200;
  const disabled = options?.disabled ?? false;

  const zoomRef = useRef<number>(zoomPercent);
  const onZoomChangeRef = useRef(onZoomChange);
  const optionsRef = useRef({ min: minZoom, max: maxZoom, disabled });

  zoomRef.current = zoomPercent;
  onZoomChangeRef.current = onZoomChange;
  optionsRef.current = { min: minZoom, max: maxZoom, disabled };

  useEffect(() => {
    const element = viewportRef.current;
    if (!element) {
      return;
    }

    let pinchStartDistance: number | null = null;
    let pinchStartZoom = zoomRef.current;

    const handleWheel = (event: WheelEvent) => {
      const { min, max, disabled: isDisabled } = optionsRef.current;
      if (isDisabled || !(event.ctrlKey || event.metaKey)) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();

      if (event.deltaY === 0) {
        return;
      }

      // Smooth step for both discrete mouse wheel and trackpad pinch gestures
      const absDelta = Math.abs(event.deltaY);
      const step = absDelta < 12 ? Math.max(1, Math.round(absDelta)) : 5;
      const direction = event.deltaY < 0 ? 1 : -1;
      const next = clampZoom(zoomRef.current + direction * step, min, max);

      if (next !== zoomRef.current) {
        zoomRef.current = next;
        onZoomChangeRef.current(next);
      }
    };

    const handleTouchStart = (event: TouchEvent) => {
      if (optionsRef.current.disabled) {
        pinchStartDistance = null;
        return;
      }
      if (event.touches.length === 2) {
        const dist = getTouchDistance(event.touches[0], event.touches[1]);
        if (dist > 0) {
          pinchStartDistance = dist;
          pinchStartZoom = zoomRef.current;
        }
      } else {
        pinchStartDistance = null;
      }
    };

    const handleTouchMove = (event: TouchEvent) => {
      const { min, max, disabled: isDisabled } = optionsRef.current;
      if (
        isDisabled ||
        pinchStartDistance === null ||
        event.touches.length !== 2
      ) {
        return;
      }

      event.preventDefault();
      const currentDist = getTouchDistance(event.touches[0], event.touches[1]);
      if (currentDist <= 0 || pinchStartDistance <= 0) {
        return;
      }

      const scale = currentDist / pinchStartDistance;
      const next = clampZoom(pinchStartZoom * scale, min, max);
      if (next !== zoomRef.current) {
        zoomRef.current = next;
        onZoomChangeRef.current(next);
      }
    };

    const handleTouchEnd = (event: TouchEvent) => {
      if (event.touches.length < 2) {
        pinchStartDistance = null;
      }
    };

    element.addEventListener('wheel', handleWheel, { passive: false });
    element.addEventListener('touchstart', handleTouchStart, { passive: true });
    element.addEventListener('touchmove', handleTouchMove, { passive: false });
    element.addEventListener('touchend', handleTouchEnd, { passive: true });
    element.addEventListener('touchcancel', handleTouchEnd, { passive: true });

    return () => {
      element.removeEventListener('wheel', handleWheel);
      element.removeEventListener('touchstart', handleTouchStart);
      element.removeEventListener('touchmove', handleTouchMove);
      element.removeEventListener('touchend', handleTouchEnd);
      element.removeEventListener('touchcancel', handleTouchEnd);
    };
  }, [viewportRef]);
}
