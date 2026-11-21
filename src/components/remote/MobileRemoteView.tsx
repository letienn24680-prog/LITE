import { useEffect, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import {
  ChevronLeft,
  ChevronRight,
  Clock,
  Crosshair,
  Eraser,
  Flashlight,
  Home,
  MessageSquareText,
  Pause,
  Play,
  Radio,
  RefreshCw,
  RotateCcw,
  Sparkles,
  Wifi,
  WifiOff,
} from 'lucide-react';
import {
  createControllerRemoteSession,
  type RemoteCommandPayload,
  type RemoteTransportMode,
} from '../../services/realtimeRemote';
import type { NormalizedPoint, RemoteChannelState } from '../../types/document';

export interface MobileRemoteViewProps {
  roomCode: string;
  onExitToHome: () => void;
}

type TouchpadTab = 'laser' | 'spotlight' | 'swipe';

const MOVE_THROTTLE_MS = 40;

function formatTimer(totalSeconds: number): string {
  const safe = Math.max(0, Math.floor(totalSeconds));
  const hours = Math.floor(safe / 3600);
  const minutes = Math.floor((safe % 3600) / 60);
  const seconds = safe % 60;
  return [hours, minutes, seconds]
    .map((unit) => String(unit).padStart(2, '0'))
    .join(':');
}

function clampUnit(value: number): number {
  if (value < 0) {
    return 0;
  }
  if (value > 1) {
    return 1;
  }
  return value;
}

export function MobileRemoteView({
  roomCode,
  onExitToHome,
}: MobileRemoteViewProps) {
  const normalizedRoomCode = roomCode.trim().toUpperCase() || 'LITE-0000';

  const [transportMode, setTransportMode] =
    useState<RemoteTransportMode>('local-broadcast');
  const [transportConnected, setTransportConnected] = useState<boolean>(false);
  const [remoteState, setRemoteState] = useState<RemoteChannelState | null>(
    null
  );
  const [displaySeconds, setDisplaySeconds] = useState<number>(0);

  const [activeTab, setActiveTab] = useState<TouchpadTab>('laser');
  const [isPointerActive, setIsPointerActive] = useState<boolean>(false);
  const [pointerPreview, setPointerPreview] = useState<NormalizedPoint>({
    x: 0.5,
    y: 0.5,
  });
  const [pinSpotlight, setPinSpotlight] = useState<boolean>(false);
  const [swipeHint, setSwipeHint] = useState<string | null>(null);

  const sessionRef = useRef<ReturnType<
    typeof createControllerRemoteSession
  > | null>(null);
  const touchpadRef = useRef<HTMLDivElement | null>(null);
  const activePointerIdRef = useRef<number | null>(null);
  const swipeStartXRef = useRef<number | null>(null);
  const lastMoveSentAtRef = useRef<number>(0);
  const trailingTimerRef = useRef<number | null>(null);
  const pendingPointRef = useRef<NormalizedPoint | null>(null);
  const swipeFeedbackTimerRef = useRef<number | null>(null);

  const clearTrailingMove = () => {
    if (trailingTimerRef.current !== null) {
      window.clearTimeout(trailingTimerRef.current);
      trailingTimerRef.current = null;
    }
    pendingPointRef.current = null;
  };

  useEffect(() => {
    const session = createControllerRemoteSession(normalizedRoomCode, {
      onStateSync: (state) => {
        setRemoteState(state);
        setDisplaySeconds(state.elapsedSeconds);
      },
      onStatusChange: (status) => {
        setTransportMode(status.mode);
        setTransportConnected(status.connected);
      },
    });
    sessionRef.current = session;

    return () => {
      clearTrailingMove();
      if (swipeFeedbackTimerRef.current !== null) {
        window.clearTimeout(swipeFeedbackTimerRef.current);
        swipeFeedbackTimerRef.current = null;
      }
      session.destroy();
      sessionRef.current = null;
    };
  }, [normalizedRoomCode]);

  useEffect(() => {
    if (!remoteState?.isTimerRunning) {
      return;
    }
    const intervalId = window.setInterval(() => {
      setDisplaySeconds((prev) => prev + 1);
    }, 1000);
    return () => {
      window.clearInterval(intervalId);
    };
  }, [remoteState?.isTimerRunning]);

  const sendCommand = (
    cmd: Omit<RemoteCommandPayload, 'sentAt' | 'clientId'>
  ) => {
    sessionRef.current?.sendCommand(cmd);
  };

  const getNormalizedPoint = (
    event: ReactPointerEvent<HTMLDivElement>
  ): NormalizedPoint | null => {
    const rect = touchpadRef.current?.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) {
      return null;
    }
    return {
      x: clampUnit((event.clientX - rect.left) / rect.width),
      y: clampUnit((event.clientY - rect.top) / rect.height),
    };
  };

  const emitPointerMove = (point: NormalizedPoint, mode: TouchpadTab) => {
    if (mode === 'laser') {
      sendCommand({ action: 'laserMove', point });
    } else if (mode === 'spotlight') {
      sendCommand({ action: 'spotlightMove', point });
    }
  };

  const scheduleThrottledMove = (
    point: NormalizedPoint,
    mode: TouchpadTab,
    forceImmediate = false
  ) => {
    const now = Date.now();
    const elapsed = now - lastMoveSentAtRef.current;

    if (forceImmediate || elapsed >= MOVE_THROTTLE_MS) {
      clearTrailingMove();
      lastMoveSentAtRef.current = now;
      emitPointerMove(point, mode);
      return;
    }

    pendingPointRef.current = point;
    if (trailingTimerRef.current === null) {
      trailingTimerRef.current = window.setTimeout(() => {
        trailingTimerRef.current = null;
        const nextPoint = pendingPointRef.current;
        pendingPointRef.current = null;
        if (nextPoint) {
          lastMoveSentAtRef.current = Date.now();
          emitPointerMove(nextPoint, mode);
        }
      }, MOVE_THROTTLE_MS - elapsed);
    }
  };

  const handleTabChange = (nextTab: TouchpadTab) => {
    if (nextTab === activeTab) {
      return;
    }
    clearTrailingMove();
    if (activeTab === 'laser') {
      sendCommand({ action: 'laserOff' });
    } else if (activeTab === 'spotlight') {
      sendCommand({ action: 'spotlightOff' });
      setPinSpotlight(false);
    }
    setIsPointerActive(false);
    activePointerIdRef.current = null;
    setActiveTab(nextTab);
  };

  const showSwipeFeedback = (label: string) => {
    setSwipeHint(label);
    if (swipeFeedbackTimerRef.current !== null) {
      window.clearTimeout(swipeFeedbackTimerRef.current);
    }
    swipeFeedbackTimerRef.current = window.setTimeout(() => {
      setSwipeHint(null);
      swipeFeedbackTimerRef.current = null;
    }, 900);
  };

  const handleTouchpadPointerDown = (
    event: ReactPointerEvent<HTMLDivElement>
  ) => {
    if (activePointerIdRef.current !== null) {
      return;
    }
    activePointerIdRef.current = event.pointerId;
    try {
      event.currentTarget.setPointerCapture(event.pointerId);
    } catch {
      // Ignore pointer capture errors
    }

    if (activeTab === 'swipe') {
      swipeStartXRef.current = event.clientX;
      setIsPointerActive(true);
      return;
    }

    const point = getNormalizedPoint(event);
    if (!point) {
      return;
    }
    setPointerPreview(point);
    setIsPointerActive(true);
    scheduleThrottledMove(point, activeTab, true);
  };

  const handleTouchpadPointerMove = (
    event: ReactPointerEvent<HTMLDivElement>
  ) => {
    if (activePointerIdRef.current !== event.pointerId) {
      return;
    }
    if (activeTab === 'swipe') {
      return;
    }
    const point = getNormalizedPoint(event);
    if (!point) {
      return;
    }
    setPointerPreview(point);
    scheduleThrottledMove(point, activeTab, false);
  };

  const finishTouchpadPointer = (
    event: ReactPointerEvent<HTMLDivElement>,
    isCancel = false
  ) => {
    if (activePointerIdRef.current !== event.pointerId) {
      return;
    }
    activePointerIdRef.current = null;
    try {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
    } catch {
      // Ignore release errors
    }

    if (activeTab === 'swipe') {
      const startX = swipeStartXRef.current;
      swipeStartXRef.current = null;
      setIsPointerActive(false);
      if (isCancel || startX === null) {
        return;
      }
      const deltaX = event.clientX - startX;
      const rect = touchpadRef.current?.getBoundingClientRect();
      if (Math.abs(deltaX) >= 36) {
        if (deltaX < 0) {
          sendCommand({ action: 'next' });
          showSwipeFeedback('Slide tiếp theo →');
        } else {
          sendCommand({ action: 'prev' });
          showSwipeFeedback('← Slide trước');
        }
      } else if (rect && rect.width > 0) {
        const tapRatio = (event.clientX - rect.left) / rect.width;
        if (tapRatio < 0.4) {
          sendCommand({ action: 'prev' });
          showSwipeFeedback('← Slide trước');
        } else {
          sendCommand({ action: 'next' });
          showSwipeFeedback('Slide tiếp theo →');
        }
      }
      return;
    }

    clearTrailingMove();
    setIsPointerActive(false);
    if (activeTab === 'laser') {
      sendCommand({ action: 'laserOff' });
    } else if (activeTab === 'spotlight' && !pinSpotlight) {
      sendCommand({ action: 'spotlightOff' });
    }
  };

  const isSyncedWithProjector = remoteState !== null;
  const totalSlides = remoteState?.totalSlides ?? 0;
  const currentSlideHuman =
    totalSlides > 0 ? Math.min(totalSlides, remoteState!.slideIndex + 1) : 0;
  const isAtFirst = !remoteState || remoteState.slideIndex <= 0;
  const isAtLast =
    !remoteState ||
    remoteState.totalSlides === 0 ||
    remoteState.slideIndex >= remoteState.totalSlides - 1;

  const channelLabel =
    transportMode === 'supabase-realtime'
      ? 'Supabase Realtime'
      : 'Kết nối cục bộ';

  return (
    <div className="h-[100dvh] w-full flex flex-col bg-[#0F141F] text-slate-100 pb-safe overflow-hidden select-none">
      {/* Top Header */}
      <header className="shrink-0 px-3.5 py-2.5 border-b border-white/10 bg-[#151B28] flex items-center justify-between gap-2">
        <button
          type="button"
          onClick={onExitToHome}
          className="inline-flex items-center gap-1.5 min-h-[40px] px-3 rounded-xl bg-white/[0.07] hover:bg-white/[0.14] border border-white/10 text-xs font-medium text-slate-100 transition-colors shrink-0"
        >
          <Home className="h-3.5 w-3.5 text-indigo-300" />
          <span>Trang chủ</span>
        </button>

        <div className="min-w-0 flex items-center gap-2">
          <span className="px-2.5 py-1 rounded-lg bg-indigo-500/15 border border-indigo-400/30 font-mono text-xs font-bold tracking-wider text-indigo-200 truncate">
            {normalizedRoomCode}
          </span>
        </div>

        <div className="flex items-center gap-1.5 shrink-0">
          <div
            title={
              isSyncedWithProjector
                ? 'Đã đồng bộ với màn hình trình chiếu'
                : 'Đang chờ tín hiệu từ màn hình trình chiếu'
            }
            className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-xl border text-[11px] font-medium ${
              isSyncedWithProjector
                ? 'bg-emerald-500/15 border-emerald-400/35 text-emerald-200'
                : transportConnected
                  ? 'bg-amber-500/15 border-amber-400/30 text-amber-200'
                  : 'bg-slate-800 border-white/10 text-slate-400'
            }`}
          >
            <span
              className={`h-2 w-2 rounded-full shrink-0 ${
                isSyncedWithProjector
                  ? 'bg-emerald-400 shadow-[0_0_8px_rgba(52,211,153,0.9)]'
                  : transportConnected
                    ? 'bg-amber-400 animate-pulse'
                    : 'bg-slate-500'
              }`}
            />
            {transportConnected ? (
              <Wifi className="h-3 w-3 shrink-0 hidden min-[390px]:inline" />
            ) : (
              <WifiOff className="h-3 w-3 shrink-0 hidden min-[390px]:inline" />
            )}
            <span className="truncate max-w-[108px]">{channelLabel}</span>
          </div>

          <button
            type="button"
            onClick={() => sendCommand({ action: 'requestState' })}
            aria-label="Đồng bộ lại trạng thái trình chiếu"
            title="Đồng bộ lại"
            className="min-h-[38px] min-w-[38px] rounded-xl bg-white/[0.06] hover:bg-white/[0.14] border border-white/10 flex items-center justify-center text-slate-300 hover:text-white transition-colors"
          >
            <RefreshCw className="h-3.5 w-3.5" />
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 min-h-0 flex flex-col gap-3 p-3.5 overflow-hidden">
        {/* Current Slide Card + Timer + Speaker Notes */}
        <section className="shrink-0 rounded-2xl bg-[#161E2E] border border-white/10 p-3.5 space-y-2.5 shadow-lg">
          <div className="flex items-start justify-between gap-2.5">
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-2 text-[11px] text-slate-400">
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-indigo-500/20 border border-indigo-400/30 text-indigo-200 font-semibold tabular-nums">
                  {totalSlides > 0
                    ? `Slide ${currentSlideHuman} / ${totalSlides}`
                    : 'Chưa kết nối slide'}
                </span>
                <span className="truncate">
                  {remoteState?.documentTitle ?? 'Đang chờ máy chiếu…'}
                </span>
              </div>
              <h1 className="mt-1.5 text-sm font-semibold text-white leading-snug line-clamp-2">
                {remoteState?.slideTitle ||
                  (isSyncedWithProjector
                    ? `Slide ${currentSlideHuman}`
                    : `Mở trình chiếu với mã ${normalizedRoomCode}`)}
              </h1>
            </div>

            {/* Synced Timer Pill */}
            <div className="shrink-0 flex flex-col items-end gap-1">
              <div
                className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-xl border font-mono text-xs font-semibold tabular-nums ${
                  remoteState?.isTimerRunning
                    ? 'bg-emerald-500/15 border-emerald-400/35 text-emerald-200'
                    : 'bg-black/30 border-white/10 text-slate-200'
                }`}
              >
                <Clock className="h-3.5 w-3.5 shrink-0" />
                <span>{formatTimer(displaySeconds)}</span>
              </div>
              {remoteState && (
                <button
                  type="button"
                  onClick={() => sendCommand({ action: 'resetTimer' })}
                  className="inline-flex items-center gap-1 text-[10px] text-slate-400 hover:text-slate-200 px-1.5 py-0.5 rounded"
                >
                  <RotateCcw className="h-2.5 w-2.5" />
                  <span>Đặt lại giờ</span>
                </button>
              )}
            </div>
          </div>

          {/* Scrollable Speaker Notes */}
          <div className="rounded-xl bg-[#0F1522] border border-white/10 px-3 py-2 max-h-24 overflow-y-auto scrollbar-dark">
            <div className="flex items-center gap-1.5 text-[10px] font-semibold uppercase tracking-wider text-indigo-300 mb-1">
              <MessageSquareText className="h-3 w-3 shrink-0" />
              <span>Ghi chú diễn giả</span>
            </div>
            {remoteState ? (
              <p className="text-xs text-slate-200/90 leading-relaxed whitespace-pre-wrap select-text">
                {remoteState.speakerNote.trim()
                  ? remoteState.speakerNote
                  : 'Slide này chưa có ghi chú diễn giả.'}
              </p>
            ) : (
              <div className="flex items-center justify-between gap-2">
                <p className="text-xs text-slate-400 leading-relaxed">
                  Hãy bấm <strong className="text-slate-200">Trình chiếu</strong>{' '}
                  trên máy tính để nhận slide và ghi chú tức thì.
                </p>
                <button
                  type="button"
                  onClick={() => sendCommand({ action: 'requestState' })}
                  className="shrink-0 px-2.5 py-1 rounded-lg bg-indigo-600/80 hover:bg-indigo-500 text-[11px] font-medium text-white"
                >
                  Kết nối
                </button>
              </div>
            )}
          </div>
        </section>

        {/* Interactive Laser / Spotlight / Tap Touchpad */}
        <section className="flex-1 min-h-0 flex flex-col rounded-2xl bg-[#161E2E] border border-white/10 p-3 gap-2.5 shadow-lg">
          <div className="shrink-0 flex items-center justify-between gap-2">
            <div className="grid grid-cols-3 gap-1.5 flex-1 bg-[#0E131E] p-1 rounded-xl border border-white/10">
              {(
                [
                  { id: 'laser', label: 'Tia Laser', icon: Crosshair },
                  { id: 'spotlight', label: 'Spotlight', icon: Flashlight },
                  { id: 'swipe', label: 'Chạm chuyển', icon: Sparkles },
                ] as const
              ).map((tab) => {
                const TabIcon = tab.icon;
                const active = activeTab === tab.id;
                return (
                  <button
                    key={tab.id}
                    type="button"
                    onClick={() => handleTabChange(tab.id)}
                    className={`min-h-[36px] px-2 rounded-lg text-[11px] font-semibold inline-flex items-center justify-center gap-1 transition-all whitespace-nowrap ${
                      active
                        ? 'bg-indigo-600 text-white shadow-sm'
                        : 'text-slate-300 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <TabIcon className="h-3.5 w-3.5 shrink-0" />
                    <span className="truncate">{tab.label}</span>
                  </button>
                );
              })}
            </div>

            {activeTab === 'spotlight' && (
              <div className="shrink-0 flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => {
                    setPinSpotlight((prev) => {
                      const next = !prev;
                      if (!next) {
                        sendCommand({ action: 'spotlightOff' });
                      }
                      return next;
                    });
                  }}
                  aria-pressed={pinSpotlight}
                  className={`min-h-[36px] px-2.5 rounded-xl border text-[11px] font-medium whitespace-nowrap transition-colors ${
                    pinSpotlight
                      ? 'bg-amber-500/20 border-amber-400/50 text-amber-200'
                      : 'bg-white/[0.06] border-white/10 text-slate-300'
                  }`}
                >
                  {pinSpotlight ? 'Đang giữ sáng' : 'Giữ vùng sáng'}
                </button>
              </div>
            )}
          </div>

          {/* Touchpad Surface */}
          <div
            ref={touchpadRef}
            onPointerDown={handleTouchpadPointerDown}
            onPointerMove={handleTouchpadPointerMove}
            onPointerUp={(e) => finishTouchpadPointer(e, false)}
            onPointerCancel={(e) => finishTouchpadPointer(e, true)}
            className="relative flex-1 min-h-[140px] rounded-2xl border border-indigo-400/25 bg-[#0C111C] overflow-hidden touch-none flex items-center justify-center cursor-crosshair"
            style={{
              backgroundImage:
                'radial-gradient(rgba(148,163,184,0.16) 1px, transparent 1px)',
              backgroundSize: '20px 20px',
            }}
          >
            {activeTab === 'laser' && (
              <>
                <div className="pointer-events-none text-center px-4 space-y-1">
                  <Radio className="h-6 w-6 mx-auto text-rose-400/75" />
                  <p className="text-xs font-medium text-slate-300">
                    Chạm &amp; rê ngón tay để chiếu Tia Laser lên màn hình
                  </p>
                  <p className="text-[11px] text-slate-500">
                    Tọa độ chuẩn 0..1 đồng bộ thời gian thực
                  </p>
                </div>
                {isPointerActive && (
                  <div
                    className="pointer-events-none absolute -translate-x-1/2 -translate-y-1/2 h-5 w-5 rounded-full bg-rose-500/90 ring-4 ring-rose-500/30 shadow-[0_0_18px_rgba(244,63,94,0.95)]"
                    style={{
                      left: `${pointerPreview.x * 100}%`,
                      top: `${pointerPreview.y * 100}%`,
                    }}
                  />
                )}
              </>
            )}

            {activeTab === 'spotlight' && (
              <>
                <div className="pointer-events-none text-center px-4 space-y-1">
                  <Flashlight className="h-6 w-6 mx-auto text-amber-300/80" />
                  <p className="text-xs font-medium text-slate-300">
                    Chạm &amp; rê để rọi đèn Spotlight làm nổi bật vùng nội dung
                  </p>
                  <p className="text-[11px] text-slate-500">
                    {pinSpotlight
                      ? 'Chế độ giữ sáng đang bật — bấm "Đang giữ sáng" để tắt'
                      : 'Nhấc tay ra sẽ tự động tắt đèn rọi'}
                  </p>
                </div>
                {(isPointerActive || pinSpotlight) && (
                  <div
                    className="pointer-events-none absolute -translate-x-1/2 -translate-y-1/2 h-20 w-20 rounded-full border-2 border-amber-300/70 bg-amber-300/15 shadow-[0_0_32px_rgba(251,191,36,0.35)]"
                    style={{
                      left: `${pointerPreview.x * 100}%`,
                      top: `${pointerPreview.y * 100}%`,
                    }}
                  />
                )}
              </>
            )}

            {activeTab === 'swipe' && (
              <div className="pointer-events-none w-full h-full flex items-center justify-between px-5 text-slate-400">
                <div className="flex flex-col items-center gap-1 opacity-75">
                  <ChevronLeft className="h-6 w-6 text-slate-300" />
                  <span className="text-[11px] font-medium">Chạm trái: Lùi</span>
                </div>
                <div className="text-center px-2">
                  <p className="text-xs font-semibold text-slate-200">
                    {swipeHint ?? 'Vuốt ngang hoặc chạm nửa màn hình'}
                  </p>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    Chuyển slide nhanh không cần nhìn nút
                  </p>
                </div>
                <div className="flex flex-col items-center gap-1 opacity-75">
                  <ChevronRight className="h-6 w-6 text-indigo-300" />
                  <span className="text-[11px] font-medium">Chạm phải: Tiếp</span>
                </div>
              </div>
            )}
          </div>
        </section>
      </main>

      {/* Large Bottom Control Dock */}
      <footer className="shrink-0 px-3.5 pb-3 pt-1.5 bg-[#151B28] border-t border-white/10 space-y-2">
        <div className="grid grid-cols-2 gap-2">
          <button
            type="button"
            onClick={() => sendCommand({ action: 'clearAnnotations' })}
            className="min-h-[56px] px-3 rounded-2xl bg-white/[0.06] hover:bg-white/[0.12] active:scale-[0.98] border border-white/10 flex items-center justify-center gap-2 text-xs font-semibold text-slate-200 transition-all"
          >
            <Eraser className="h-4 w-4 text-rose-300 shrink-0" />
            <span>Xóa nét vẽ</span>
            {remoteState && remoteState.annotationsCount > 0 && (
              <span className="px-1.5 py-0.5 rounded-full bg-rose-500/20 text-rose-200 text-[10px] font-bold tabular-nums">
                {remoteState.annotationsCount}
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={() => sendCommand({ action: 'toggleTimer' })}
            className={`min-h-[56px] px-3 rounded-2xl border active:scale-[0.98] flex items-center justify-center gap-2 text-xs font-semibold transition-all ${
              remoteState?.isTimerRunning
                ? 'bg-emerald-600/25 hover:bg-emerald-600/35 border-emerald-400/40 text-emerald-100'
                : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
            }`}
          >
            {remoteState?.isTimerRunning ? (
              <Pause className="h-4 w-4 text-emerald-300 shrink-0" />
            ) : (
              <Play className="h-4 w-4 text-indigo-300 shrink-0" />
            )}
            <span>
              {remoteState?.isTimerRunning ? 'Tạm dừng giờ' : 'Bấm giờ'}
            </span>
          </button>
        </div>

        <div className="grid grid-cols-2 gap-2.5">
          <button
            type="button"
            onClick={() => sendCommand({ action: 'prev' })}
            disabled={isSyncedWithProjector && isAtFirst}
            className="min-h-[56px] px-4 rounded-2xl bg-white/10 hover:bg-white/15 active:scale-[0.98] border border-white/15 flex items-center justify-center gap-2 text-sm font-semibold text-white transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <ChevronLeft className="h-5 w-5 shrink-0" />
            <span>Slide Trước</span>
          </button>

          <button
            type="button"
            onClick={() => sendCommand({ action: 'next' })}
            disabled={isSyncedWithProjector && isAtLast}
            className="min-h-[56px] px-4 rounded-2xl bg-indigo-600 hover:bg-indigo-500 active:scale-[0.98] shadow-lg shadow-indigo-600/25 flex items-center justify-center gap-2 text-sm font-semibold text-white transition-all disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <span>Slide Tiếp</span>
            <ChevronRight className="h-5 w-5 shrink-0" />
          </button>
        </div>
      </footer>
    </div>
  );
}
