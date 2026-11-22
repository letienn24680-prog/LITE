import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { MouseEvent as ReactMouseEvent, RefObject } from 'react';

export interface CustomVerticalScrollbarDragInfo {
  percent: number;
  scrollTop: number;
  scrollHeight: number;
  clientHeight: number;
}

export interface CustomVerticalScrollbarProps {
  scrollContainerRef: RefObject<HTMLElement | null>;
  theme?: 'light' | 'dark';
  className?: string;
  bottomOffsetClass?: string;
  getDragLabel?: (info: CustomVerticalScrollbarDragInfo) => string;
}

interface ScrollMetrics {
  scrollTop: number;
  scrollHeight: number;
  clientHeight: number;
  trackHeight: number;
}

const MIN_THUMB_HEIGHT = 28;
const SINGLE_STEP_PX = 160;
const HOLD_DELAY_MS = 180;
const HOLD_INTERVAL_MS = 20;
const HOLD_STEP_PX = 18;

export function CustomVerticalScrollbar({
  scrollContainerRef,
  theme = 'light',
  className = '',
  bottomOffsetClass = 'bottom-0',
  getDragLabel,
}: CustomVerticalScrollbarProps): JSX.Element | null {
  const trackRef = useRef<HTMLDivElement | null>(null);
  const thumbRef = useRef<HTMLDivElement | null>(null);
  const upBtnRef = useRef<HTMLButtonElement | null>(null);
  const downBtnRef = useRef<HTMLButtonElement | null>(null);

  const [metrics, setMetrics] = useState<ScrollMetrics>({
    scrollTop: 0,
    scrollHeight: 0,
    clientHeight: 0,
    trackHeight: 0,
  });
  const [isDragging, setIsDragging] = useState(false);

  const holdDelayTimerRef = useRef<number | null>(null);
  const holdIntervalRef = useRef<number | null>(null);
  const didHoldRepeatRef = useRef(false);
  const touchHandledAtRef = useRef(0);

  const measure = useCallback(() => {
    const el = scrollContainerRef.current;
    if (!el) return;

    const nextScrollTop = el.scrollTop;
    const nextScrollHeight = el.scrollHeight;
    const nextClientHeight = el.clientHeight;
    const nextTrackHeight = trackRef.current?.clientHeight ?? 0;

    setMetrics((prev) => {
      if (
        Math.abs(prev.scrollTop - nextScrollTop) < 0.5 &&
        prev.scrollHeight === nextScrollHeight &&
        prev.clientHeight === nextClientHeight &&
        prev.trackHeight === nextTrackHeight
      ) {
        return prev;
      }
      return {
        scrollTop: nextScrollTop,
        scrollHeight: nextScrollHeight,
        clientHeight: nextClientHeight,
        trackHeight: nextTrackHeight,
      };
    });
  }, [scrollContainerRef]);

  const stopHoldScroll = useCallback(() => {
    if (holdDelayTimerRef.current !== null) {
      window.clearTimeout(holdDelayTimerRef.current);
      holdDelayTimerRef.current = null;
    }
    if (holdIntervalRef.current !== null) {
      window.clearInterval(holdIntervalRef.current);
      holdIntervalRef.current = null;
    }
  }, []);

  useEffect(() => {
    return () => {
      stopHoldScroll();
    };
  }, [stopHoldScroll]);

  // Observe scroll container, its first child, DOM mutations, and window resize
  useEffect(() => {
    const el = scrollContainerRef.current;
    if (!el) return;

    measure();

    let observedChild: Element | null = el.firstElementChild;
    const resizeObserver = new ResizeObserver(() => {
      measure();
    });
    resizeObserver.observe(el);
    if (observedChild) {
      resizeObserver.observe(observedChild);
    }

    const mutationObserver = new MutationObserver(() => {
      const latestChild = el.firstElementChild;
      if (latestChild !== observedChild) {
        if (observedChild) {
          resizeObserver.unobserve(observedChild);
        }
        observedChild = latestChild;
        if (observedChild) {
          resizeObserver.observe(observedChild);
        }
      }
      measure();
    });

    mutationObserver.observe(el, {
      childList: true,
      subtree: true,
      attributes: true,
    });

    const handleScroll = () => measure();
    const handleWindowResize = () => measure();

    el.addEventListener('scroll', handleScroll, { passive: true });
    window.addEventListener('resize', handleWindowResize);

    return () => {
      el.removeEventListener('scroll', handleScroll);
      window.removeEventListener('resize', handleWindowResize);
      mutationObserver.disconnect();
      resizeObserver.disconnect();
    };
  }, [scrollContainerRef, measure]);

  const isScrollable = metrics.scrollHeight > metrics.clientHeight + 2;

  // Measure and observe track element once rendered
  useLayoutEffect(() => {
    if (!isScrollable) return;
    const trackEl = trackRef.current;
    if (!trackEl) return;

    measure();
    const trackObserver = new ResizeObserver(() => {
      measure();
    });
    trackObserver.observe(trackEl);
    return () => {
      trackObserver.disconnect();
    };
  }, [isScrollable, measure]);

  const maxScrollTop = Math.max(0, metrics.scrollHeight - metrics.clientHeight);
  const fallbackTrackHeight = Math.max(MIN_THUMB_HEIGHT, metrics.clientHeight - 48);
  const effectiveTrackHeight = Math.max(
    MIN_THUMB_HEIGHT,
    metrics.trackHeight > 0 ? metrics.trackHeight : fallbackTrackHeight
  );
  const viewportRatio =
    metrics.scrollHeight > 0 ? metrics.clientHeight / metrics.scrollHeight : 1;
  const thumbHeight = Math.min(
    effectiveTrackHeight,
    Math.max(MIN_THUMB_HEIGHT, Math.round(effectiveTrackHeight * viewportRatio))
  );
  const maxThumbTop = Math.max(0, effectiveTrackHeight - thumbHeight);
  const progress =
    maxScrollTop > 0 ? Math.min(1, Math.max(0, metrics.scrollTop / maxScrollTop)) : 0;
  const thumbTop = maxThumbTop * progress;

  const startHoldScroll = useCallback(
    (direction: -1 | 1) => {
      const el = scrollContainerRef.current;
      if (!el) return;

      stopHoldScroll();
      didHoldRepeatRef.current = false;

      holdDelayTimerRef.current = window.setTimeout(() => {
        didHoldRepeatRef.current = true;
        el.scrollTop += direction * HOLD_STEP_PX;
        holdIntervalRef.current = window.setInterval(() => {
          const target = scrollContainerRef.current;
          if (!target) {
            stopHoldScroll();
            return;
          }
          target.scrollTop += direction * HOLD_STEP_PX;
        }, HOLD_INTERVAL_MS);
      }, HOLD_DELAY_MS);
    },
    [scrollContainerRef, stopHoldScroll]
  );

  const performSingleStep = useCallback(
    (direction: -1 | 1) => {
      const el = scrollContainerRef.current;
      if (!el) return;
      el.scrollBy({
        top: direction * SINGLE_STEP_PX,
        behavior: 'smooth',
      });
    },
    [scrollContainerRef]
  );

  // Non-passive touch handlers for arrow buttons (so hold doesn't trigger page scroll or ghost click)
  useEffect(() => {
    if (!isScrollable) return;

    const attachArrowTouch = (
      btn: HTMLButtonElement | null,
      direction: -1 | 1
    ): (() => void) | undefined => {
      if (!btn) return undefined;

      const handleTouchStart = (event: TouchEvent) => {
        event.preventDefault();
        event.stopPropagation();
        startHoldScroll(direction);
      };

      const handleTouchEnd = (event: TouchEvent) => {
        event.preventDefault();
        event.stopPropagation();
        const wasRepeating = didHoldRepeatRef.current;
        stopHoldScroll();
        didHoldRepeatRef.current = false;
        touchHandledAtRef.current = Date.now();
        if (!wasRepeating) {
          performSingleStep(direction);
        }
      };

      const handleTouchCancel = () => {
        stopHoldScroll();
        didHoldRepeatRef.current = false;
      };

      btn.addEventListener('touchstart', handleTouchStart, { passive: false });
      btn.addEventListener('touchend', handleTouchEnd, { passive: false });
      btn.addEventListener('touchcancel', handleTouchCancel);

      return () => {
        btn.removeEventListener('touchstart', handleTouchStart);
        btn.removeEventListener('touchend', handleTouchEnd);
        btn.removeEventListener('touchcancel', handleTouchCancel);
      };
    };

    const cleanupUp = attachArrowTouch(upBtnRef.current, -1);
    const cleanupDown = attachArrowTouch(downBtnRef.current, 1);

    return () => {
      cleanupUp?.();
      cleanupDown?.();
    };
  }, [isScrollable, startHoldScroll, stopHoldScroll, performSingleStep]);

  // Non-passive touchstart on thumb for smooth finger dragging
  const beginDragAtClientY = useCallback(
    (startClientY: number) => {
      const el = scrollContainerRef.current;
      const trackEl = trackRef.current;
      if (!el || !trackEl) return;

      const startScrollTop = el.scrollTop;
      const currentMaxScroll = Math.max(0, el.scrollHeight - el.clientHeight);
      const currentTrackH = Math.max(MIN_THUMB_HEIGHT, trackEl.clientHeight);
      const currentRatio = el.scrollHeight > 0 ? el.clientHeight / el.scrollHeight : 1;
      const currentThumbH = Math.min(
        currentTrackH,
        Math.max(MIN_THUMB_HEIGHT, Math.round(currentTrackH * currentRatio))
      );
      const currentMaxThumbTop = Math.max(1, currentTrackH - currentThumbH);

      setIsDragging(true);

      const updateScrollForClientY = (clientY: number) => {
        const deltaY = clientY - startClientY;
        const deltaScroll = (deltaY / currentMaxThumbTop) * currentMaxScroll;
        el.scrollTop = Math.max(0, Math.min(currentMaxScroll, startScrollTop + deltaScroll));
      };

      const handleMouseMove = (moveEvent: MouseEvent) => {
        moveEvent.preventDefault();
        updateScrollForClientY(moveEvent.clientY);
      };

      const handleTouchMove = (moveEvent: TouchEvent) => {
        if (moveEvent.touches.length === 0) return;
        moveEvent.preventDefault();
        updateScrollForClientY(moveEvent.touches[0].clientY);
      };

      const handleEnd = () => {
        setIsDragging(false);
        window.removeEventListener('mousemove', handleMouseMove);
        window.removeEventListener('mouseup', handleEnd);
        window.removeEventListener('touchmove', handleTouchMove);
        window.removeEventListener('touchend', handleEnd);
        window.removeEventListener('touchcancel', handleEnd);
      };

      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleEnd);
      window.addEventListener('touchmove', handleTouchMove, { passive: false });
      window.addEventListener('touchend', handleEnd);
      window.addEventListener('touchcancel', handleEnd);
    },
    [scrollContainerRef]
  );

  useEffect(() => {
    if (!isScrollable) return;
    const thumbEl = thumbRef.current;
    if (!thumbEl) return;

    const handleThumbTouchStart = (event: TouchEvent) => {
      if (event.touches.length === 0) return;
      event.preventDefault();
      event.stopPropagation();
      beginDragAtClientY(event.touches[0].clientY);
    };

    thumbEl.addEventListener('touchstart', handleThumbTouchStart, { passive: false });
    return () => {
      thumbEl.removeEventListener('touchstart', handleThumbTouchStart);
    };
  }, [isScrollable, beginDragAtClientY]);

  if (!isScrollable) {
    return null;
  }

  const handleArrowMouseDown = (
    event: ReactMouseEvent<HTMLButtonElement>,
    direction: -1 | 1
  ) => {
    if (event.button !== 0) return;
    event.preventDefault();
    startHoldScroll(direction);
  };

  const handleArrowMouseUpOrLeave = () => {
    stopHoldScroll();
  };

  const handleArrowClick = (
    event: ReactMouseEvent<HTMLButtonElement>,
    direction: -1 | 1
  ) => {
    event.preventDefault();
    if (Date.now() - touchHandledAtRef.current < 400) {
      return;
    }
    if (didHoldRepeatRef.current) {
      didHoldRepeatRef.current = false;
      return;
    }
    performSingleStep(direction);
  };

  const handleThumbMouseDown = (event: ReactMouseEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    beginDragAtClientY(event.clientY);
  };

  const handleTrackMouseDown = (event: ReactMouseEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    event.preventDefault();
  };

  const handleTrackClick = (event: ReactMouseEvent<HTMLDivElement>) => {
    if (event.target === thumbRef.current) return;
    const el = scrollContainerRef.current;
    const trackEl = trackRef.current;
    if (!el || !trackEl) return;

    const trackRect = trackEl.getBoundingClientRect();
    const clickY = event.clientY - trackRect.top;
    const targetThumbTop = Math.max(
      0,
      Math.min(maxThumbTop, clickY - thumbHeight / 2)
    );
    const targetScrollTop =
      maxThumbTop > 0 ? (targetThumbTop / maxThumbTop) * maxScrollTop : 0;

    el.scrollTo({
      top: targetScrollTop,
      behavior: 'smooth',
    });
  };

  const dragInfo: CustomVerticalScrollbarDragInfo = {
    percent: progress * 100,
    scrollTop: metrics.scrollTop,
    scrollHeight: metrics.scrollHeight,
    clientHeight: metrics.clientHeight,
  };

  const dragLabelText = getDragLabel
    ? getDragLabel(dragInfo)
    : `${Math.round(dragInfo.percent)}%`;

  const isDark = theme === 'dark';

  return (
    <div
      aria-label="Thanh cuộn dọc tài liệu"
      onMouseDown={(event) => event.preventDefault()}
      className={`absolute right-0 top-0 ${bottomOffsetClass} z-30 flex w-4 select-none flex-col items-center justify-between border-l ${
        isDark
          ? 'border-white/10 bg-[#181D27]'
          : 'border-slate-200/80 bg-[#F1F1EF]'
      } ${className}`}
    >
      {/* Top arrow button (▲) */}
      <button
        ref={upBtnRef}
        type="button"
        aria-label="Cuộn lên"
        title="Cuộn lên"
        onMouseDown={(event) => handleArrowMouseDown(event, -1)}
        onMouseUp={handleArrowMouseUpOrLeave}
        onMouseLeave={handleArrowMouseUpOrLeave}
        onClick={(event) => handleArrowClick(event, -1)}
        className={`flex h-6 w-full shrink-0 items-center justify-center transition-colors focus:outline-none ${
          isDark
            ? 'text-slate-400 hover:bg-white/5 hover:text-slate-200 active:text-white'
            : 'text-[#9CA3AF] hover:bg-black/[0.04] hover:text-[#4B5563] active:text-[#4B5563]'
        }`}
      >
        <svg viewBox="0 0 10 6" className="h-2 w-2.5 fill-current" aria-hidden="true">
          <path d="M5 0.5L9.25 5.25H0.75L5 0.5Z" />
        </svg>
      </button>

      {/* Center track & draggable pill thumb */}
      <div
        ref={trackRef}
        onMouseDown={handleTrackMouseDown}
        onClick={handleTrackClick}
        className="relative flex w-full flex-1 cursor-pointer justify-center"
      >
        <div
          ref={thumbRef}
          role="scrollbar"
          aria-orientation="vertical"
          aria-valuemin={0}
          aria-valuemax={100}
          aria-valuenow={Math.round(dragInfo.percent)}
          onMouseDown={handleThumbMouseDown}
          style={{
            height: `${thumbHeight}px`,
            top: `${thumbTop}px`,
          }}
          className={`absolute w-2 touch-none rounded-full transition-colors ${
            isDark
              ? isDragging
                ? 'bg-slate-300'
                : 'bg-slate-500/85 hover:bg-slate-400 active:bg-slate-300'
              : isDragging
                ? 'bg-[#6B7280]'
                : 'bg-[#C4C9D1] hover:bg-[#9CA3AF] active:bg-[#6B7280]'
          }`}
        />

        {isDragging && (
          <div
            style={{ top: `${thumbTop + thumbHeight / 2}px` }}
            className="pointer-events-none absolute right-5 -translate-y-1/2 whitespace-nowrap rounded-md bg-slate-900 px-2 py-0.5 text-[11px] font-semibold text-white shadow-md"
          >
            {dragLabelText}
          </div>
        )}
      </div>

      {/* Bottom arrow button (▼) */}
      <button
        ref={downBtnRef}
        type="button"
        aria-label="Cuộn xuống"
        title="Cuộn xuống"
        onMouseDown={(event) => handleArrowMouseDown(event, 1)}
        onMouseUp={handleArrowMouseUpOrLeave}
        onMouseLeave={handleArrowMouseUpOrLeave}
        onClick={(event) => handleArrowClick(event, 1)}
        className={`flex h-6 w-full shrink-0 items-center justify-center transition-colors focus:outline-none ${
          isDark
            ? 'text-slate-400 hover:bg-white/5 hover:text-slate-200 active:text-white'
            : 'text-[#9CA3AF] hover:bg-black/[0.04] hover:text-[#4B5563] active:text-[#4B5563]'
        }`}
      >
        <svg viewBox="0 0 10 6" className="h-2 w-2.5 fill-current" aria-hidden="true">
          <path d="M5 5.5L0.75 0.75H9.25L5 5.5Z" />
        </svg>
      </button>
    </div>
  );
}
