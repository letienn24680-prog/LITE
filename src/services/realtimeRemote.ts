import { isSupabaseConfigured, supabase } from '../lib/supabase';
import type {
  AnnotationStroke,
  NormalizedPoint,
  RemoteChannelState,
  RemoteCommandAction,
} from '../types/document';

export type RemoteTransportMode = 'supabase-realtime' | 'local-broadcast';

export interface RemoteCommandPayload {
  action: RemoteCommandAction;
  slideIndex?: number;
  point?: NormalizedPoint;
  clientId?: string;
  sentAt: number;
}

interface PresenterStatus {
  mode: RemoteTransportMode;
  connected: boolean;
  peerCount: number;
}

interface ControllerStatus {
  mode: RemoteTransportMode;
  connected: boolean;
}

type BroadcastEnvelope =
  | {
      kind: 'remote_cmd';
      msgId: string;
      payload: RemoteCommandPayload;
    }
  | {
      kind: 'presenter_state';
      msgId: string;
      payload: RemoteChannelState;
    }
  | {
      kind: 'presenter_annotations';
      msgId: string;
      payload: {
        slideIndex: number;
        strokes: AnnotationStroke[];
      };
    }
  | {
      kind: 'peer_hello' | 'peer_leave';
      role: 'presenter' | 'controller';
      clientId: string;
      sentAt: number;
    };

const VALID_ACTIONS: ReadonlySet<RemoteCommandAction> = new Set<RemoteCommandAction>([
  'next',
  'prev',
  'goto',
  'laserMove',
  'laserOff',
  'spotlightMove',
  'spotlightOff',
  'clearAnnotations',
  'toggleTimer',
  'resetTimer',
  'requestState',
]);

function normalizeRoomCode(roomCode: string): string {
  const cleaned = roomCode.trim().toUpperCase();
  return cleaned || 'LITE-0000';
}

function createUniqueId(prefix: string): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${prefix}-${crypto.randomUUID()}`;
  }
  const randomPart = Math.random().toString(36).slice(2, 10);
  return `${prefix}-${Date.now().toString(36)}-${randomPart}`;
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

function parseNormalizedPoint(raw: unknown): NormalizedPoint | undefined {
  if (!raw || typeof raw !== 'object') {
    return undefined;
  }
  const candidate = raw as { x?: unknown; y?: unknown };
  if (
    typeof candidate.x === 'number' &&
    Number.isFinite(candidate.x) &&
    typeof candidate.y === 'number' &&
    Number.isFinite(candidate.y)
  ) {
    return {
      x: clampUnit(candidate.x),
      y: clampUnit(candidate.y),
    };
  }
  return undefined;
}

function parseRemoteCommandPayload(raw: unknown): RemoteCommandPayload | null {
  if (!raw || typeof raw !== 'object') {
    return null;
  }
  const candidate = raw as Record<string, unknown>;
  if (
    typeof candidate.action !== 'string' ||
    !VALID_ACTIONS.has(candidate.action as RemoteCommandAction)
  ) {
    return null;
  }
  const parsed: RemoteCommandPayload = {
    action: candidate.action as RemoteCommandAction,
    sentAt:
      typeof candidate.sentAt === 'number' && Number.isFinite(candidate.sentAt)
        ? candidate.sentAt
        : Date.now(),
  };
  if (
    typeof candidate.slideIndex === 'number' &&
    Number.isFinite(candidate.slideIndex)
  ) {
    parsed.slideIndex = Math.max(0, Math.floor(candidate.slideIndex));
  }
  const point = parseNormalizedPoint(candidate.point);
  if (point) {
    parsed.point = point;
  }
  if (typeof candidate.clientId === 'string' && candidate.clientId.length > 0) {
    parsed.clientId = candidate.clientId;
  }
  return parsed;
}

function parseRemoteChannelState(raw: unknown): RemoteChannelState | null {
  if (!raw || typeof raw !== 'object') {
    return null;
  }
  const candidate = raw as Record<string, unknown>;
  if (
    typeof candidate.roomCode !== 'string' ||
    typeof candidate.documentId !== 'string' ||
    typeof candidate.documentTitle !== 'string' ||
    typeof candidate.slideIndex !== 'number' ||
    typeof candidate.totalSlides !== 'number'
  ) {
    return null;
  }
  return {
    roomCode: candidate.roomCode,
    documentId: candidate.documentId,
    documentTitle: candidate.documentTitle,
    slideIndex: Math.max(0, Math.floor(candidate.slideIndex)),
    totalSlides: Math.max(0, Math.floor(candidate.totalSlides)),
    slideTitle:
      typeof candidate.slideTitle === 'string' ? candidate.slideTitle : '',
    speakerNote:
      typeof candidate.speakerNote === 'string' ? candidate.speakerNote : '',
    elapsedSeconds:
      typeof candidate.elapsedSeconds === 'number' &&
      Number.isFinite(candidate.elapsedSeconds)
        ? Math.max(0, Math.floor(candidate.elapsedSeconds))
        : 0,
    isTimerRunning: Boolean(candidate.isTimerRunning),
    annotationsCount:
      typeof candidate.annotationsCount === 'number' &&
      Number.isFinite(candidate.annotationsCount)
        ? Math.max(0, Math.floor(candidate.annotationsCount))
        : 0,
  };
}

function createMessageDeduplicator(maxEntries = 240): (msgId: string) => boolean {
  const seenSet = new Set<string>();
  const queue: string[] = [];

  return (msgId: string): boolean => {
    if (!msgId) {
      return false;
    }
    if (seenSet.has(msgId)) {
      return true;
    }
    seenSet.add(msgId);
    queue.push(msgId);
    if (queue.length > maxEntries) {
      const oldest = queue.shift();
      if (oldest) {
        seenSet.delete(oldest);
      }
    }
    return false;
  };
}

function openBroadcastChannelSafe(name: string): BroadcastChannel | null {
  try {
    if (typeof BroadcastChannel !== 'undefined') {
      return new BroadcastChannel(name);
    }
  } catch {
    // Ignore when BroadcastChannel is unavailable in restricted contexts
  }
  return null;
}

function postToBroadcastChannelSafe(
  channel: BroadcastChannel | null,
  message: BroadcastEnvelope
): void {
  if (!channel) {
    return;
  }
  try {
    channel.postMessage(message);
  } catch {
    // Ignore closed/blocked BroadcastChannel errors
  }
}

export function createPresenterRemoteSession(
  roomCode: string,
  handlers: {
    onCommand: (cmd: RemoteCommandPayload) => void;
    onStatusChange?: (status: PresenterStatus) => void;
  }
): {
  broadcastState: (state: RemoteChannelState) => void;
  broadcastAnnotations: (slideIndex: number, strokes: AnnotationStroke[]) => void;
  destroy: () => void;
} {
  const normalizedRoom = normalizeRoomCode(roomCode);
  const clientId = createUniqueId('presenter');
  const isDuplicateMessage = createMessageDeduplicator();

  let destroyed = false;
  let supabaseSubscribed = false;
  let supabaseErrored = false;
  let lastBroadcastedState: RemoteChannelState | null = null;
  let lastStatusKey = '';

  const localControllerPeers = new Set<string>();
  const supabaseControllerPeers = new Set<string>();

  const bc = openBroadcastChannelSafe(`lite-remote-${normalizedRoom}`);

  const emitStatus = () => {
    if (destroyed || !handlers.onStatusChange) {
      return;
    }
    const peerIds = new Set<string>([
      ...localControllerPeers,
      ...supabaseControllerPeers,
    ]);
    const useSupabase =
      isSupabaseConfigured && supabase !== null && !supabaseErrored;
    const mode: RemoteTransportMode = useSupabase
      ? 'supabase-realtime'
      : 'local-broadcast';
    const connected = useSupabase ? supabaseSubscribed : bc !== null;
    const peerCount = peerIds.size;
    const key = `${mode}|${connected}|${peerCount}`;
    if (key === lastStatusKey) {
      return;
    }
    lastStatusKey = key;
    handlers.onStatusChange({ mode, connected, peerCount });
  };

  const handleIncomingCommand = (
    rawCmd: unknown,
    msgId?: string,
    fromLocalBc = false
  ) => {
    if (destroyed) {
      return;
    }
    if (msgId && isDuplicateMessage(msgId)) {
      return;
    }
    const parsed = parseRemoteCommandPayload(rawCmd);
    if (!parsed || parsed.clientId === clientId) {
      return;
    }
    if (fromLocalBc && parsed.clientId) {
      const prevSize = localControllerPeers.size;
      localControllerPeers.add(parsed.clientId);
      if (localControllerPeers.size !== prevSize) {
        emitStatus();
      }
    }
    handlers.onCommand(parsed);
  };

  if (bc) {
    bc.onmessage = (event: MessageEvent<unknown>) => {
      if (destroyed) {
        return;
      }
      const data = event.data;
      if (!data || typeof data !== 'object') {
        return;
      }
      const envelope = data as Partial<BroadcastEnvelope> & Record<string, unknown>;
      if (envelope.kind === 'remote_cmd') {
        handleIncomingCommand(
          envelope.payload,
          typeof envelope.msgId === 'string' ? envelope.msgId : undefined,
          true
        );
      } else if (
        envelope.kind === 'peer_hello' &&
        envelope.role === 'controller' &&
        typeof envelope.clientId === 'string'
      ) {
        localControllerPeers.add(envelope.clientId);
        emitStatus();
        if (lastBroadcastedState) {
          postToBroadcastChannelSafe(bc, {
            kind: 'presenter_state',
            msgId: createUniqueId('msg'),
            payload: lastBroadcastedState,
          });
        }
      } else if (
        envelope.kind === 'peer_leave' &&
        envelope.role === 'controller' &&
        typeof envelope.clientId === 'string'
      ) {
        localControllerPeers.delete(envelope.clientId);
        emitStatus();
      } else if (typeof envelope.action === 'string') {
        handleIncomingCommand(envelope, undefined, true);
      }
    };

    postToBroadcastChannelSafe(bc, {
      kind: 'peer_hello',
      role: 'presenter',
      clientId,
      sentAt: Date.now(),
    });
  }

  const supabaseClient = isSupabaseConfigured ? supabase : null;
  const supabaseChannel = supabaseClient
    ? supabaseClient.channel(`lite-remote:${normalizedRoom}`, {
        config: {
          broadcast: { self: false },
          presence: { key: clientId },
        },
      })
    : null;

  if (supabaseChannel) {
    supabaseChannel
      .on<Record<string, unknown>>(
        'broadcast',
        { event: 'remote_cmd' },
        ({ payload }) => {
          const msgId =
            typeof payload._msgId === 'string' ? payload._msgId : undefined;
          handleIncomingCommand(payload, msgId, false);
        }
      )
      .on('presence', { event: 'sync' }, () => {
        if (destroyed) {
          return;
        }
        supabaseControllerPeers.clear();
        const presenceMap = supabaseChannel.presenceState<{
          role?: string;
          clientId?: string;
        }>();
        for (const [key, entries] of Object.entries(presenceMap)) {
          if (key === clientId) {
            continue;
          }
          const isController = entries.some(
            (entry) => entry.role === 'controller'
          );
          if (isController) {
            supabaseControllerPeers.add(key);
          }
        }
        emitStatus();
      })
      .subscribe((status) => {
        if (destroyed) {
          return;
        }
        if (status === 'SUBSCRIBED') {
          supabaseSubscribed = true;
          supabaseErrored = false;
          void supabaseChannel
            .track({ role: 'presenter', clientId, joinedAt: Date.now() })
            .catch(() => {});
          if (lastBroadcastedState) {
            void supabaseChannel
              .send({
                type: 'broadcast',
                event: 'presenter_state',
                payload: {
                  ...lastBroadcastedState,
                  _msgId: createUniqueId('msg'),
                },
              })
              .catch(() => {});
          }
          emitStatus();
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
          supabaseSubscribed = false;
          supabaseErrored = true;
          emitStatus();
        } else if (status === 'CLOSED') {
          supabaseSubscribed = false;
          emitStatus();
        }
      });
  }

  emitStatus();

  return {
    broadcastState(state: RemoteChannelState) {
      if (destroyed) {
        return;
      }
      lastBroadcastedState = state;
      const msgId = createUniqueId('msg');

      postToBroadcastChannelSafe(bc, {
        kind: 'presenter_state',
        msgId,
        payload: state,
      });

      if (supabaseChannel && supabaseSubscribed) {
        void supabaseChannel
          .send({
            type: 'broadcast',
            event: 'presenter_state',
            payload: {
              ...state,
              _msgId: msgId,
            },
          })
          .catch(() => {});
      }
    },

    broadcastAnnotations(slideIndex: number, strokes: AnnotationStroke[]) {
      if (destroyed) {
        return;
      }
      const msgId = createUniqueId('msg');
      const payload = { slideIndex, strokes };

      postToBroadcastChannelSafe(bc, {
        kind: 'presenter_annotations',
        msgId,
        payload,
      });

      if (supabaseChannel && supabaseSubscribed) {
        void supabaseChannel
          .send({
            type: 'broadcast',
            event: 'presenter_annotations',
            payload: {
              ...payload,
              _msgId: msgId,
            },
          })
          .catch(() => {});
      }
    },

    destroy() {
      if (destroyed) {
        return;
      }
      destroyed = true;
      if (bc) {
        postToBroadcastChannelSafe(bc, {
          kind: 'peer_leave',
          role: 'presenter',
          clientId,
          sentAt: Date.now(),
        });
        try {
          bc.close();
        } catch {
          // Ignore close errors
        }
      }
      if (supabaseChannel && supabaseClient) {
        void supabaseChannel.untrack().catch(() => {});
        void supabaseClient.removeChannel(supabaseChannel).catch(() => {});
      }
    },
  };
}

export function createControllerRemoteSession(
  roomCode: string,
  handlers: {
    onStateSync: (state: RemoteChannelState) => void;
    onStatusChange?: (status: ControllerStatus) => void;
  }
): {
  sendCommand: (
    cmd: Omit<RemoteCommandPayload, 'sentAt' | 'clientId'>
  ) => void;
  destroy: () => void;
} {
  const normalizedRoom = normalizeRoomCode(roomCode);
  const clientId = createUniqueId('controller');
  const isDuplicateMessage = createMessageDeduplicator();

  let destroyed = false;
  let supabaseSubscribed = false;
  let supabaseErrored = false;
  let lastStatusKey = '';

  const bc = openBroadcastChannelSafe(`lite-remote-${normalizedRoom}`);

  const emitStatus = () => {
    if (destroyed || !handlers.onStatusChange) {
      return;
    }
    const useSupabase =
      isSupabaseConfigured && supabase !== null && !supabaseErrored;
    const mode: RemoteTransportMode = useSupabase
      ? 'supabase-realtime'
      : 'local-broadcast';
    const connected = useSupabase ? supabaseSubscribed : bc !== null;
    const key = `${mode}|${connected}`;
    if (key === lastStatusKey) {
      return;
    }
    lastStatusKey = key;
    handlers.onStatusChange({ mode, connected });
  };

  const handleIncomingState = (rawState: unknown, msgId?: string) => {
    if (destroyed) {
      return;
    }
    if (msgId && isDuplicateMessage(msgId)) {
      return;
    }
    const parsed = parseRemoteChannelState(rawState);
    if (!parsed) {
      return;
    }
    handlers.onStateSync(parsed);
  };

  const supabaseClient = isSupabaseConfigured ? supabase : null;
  const supabaseChannel = supabaseClient
    ? supabaseClient.channel(`lite-remote:${normalizedRoom}`, {
        config: {
          broadcast: { self: false },
          presence: { key: clientId },
        },
      })
    : null;

  const sendCommand = (
    cmd: Omit<RemoteCommandPayload, 'sentAt' | 'clientId'>
  ) => {
    if (destroyed) {
      return;
    }
    const msgId = createUniqueId('msg');
    const payload: RemoteCommandPayload = {
      ...cmd,
      clientId,
      sentAt: Date.now(),
    };

    postToBroadcastChannelSafe(bc, {
      kind: 'remote_cmd',
      msgId,
      payload,
    });

    if (supabaseChannel && supabaseSubscribed) {
      void supabaseChannel
        .send({
          type: 'broadcast',
          event: 'remote_cmd',
          payload: {
            ...payload,
            _msgId: msgId,
          },
        })
        .catch(() => {});
    }
  };

  if (bc) {
    bc.onmessage = (event: MessageEvent<unknown>) => {
      if (destroyed) {
        return;
      }
      const data = event.data;
      if (!data || typeof data !== 'object') {
        return;
      }
      const envelope = data as Partial<BroadcastEnvelope> & Record<string, unknown>;
      if (envelope.kind === 'presenter_state') {
        handleIncomingState(
          envelope.payload,
          typeof envelope.msgId === 'string' ? envelope.msgId : undefined
        );
      } else if (
        envelope.kind === 'peer_hello' &&
        envelope.role === 'presenter'
      ) {
        sendCommand({ action: 'requestState' });
      }
    };

    postToBroadcastChannelSafe(bc, {
      kind: 'peer_hello',
      role: 'controller',
      clientId,
      sentAt: Date.now(),
    });
    sendCommand({ action: 'requestState' });
  }

  if (supabaseChannel) {
    supabaseChannel
      .on<Record<string, unknown>>(
        'broadcast',
        { event: 'presenter_state' },
        ({ payload }) => {
          const msgId =
            typeof payload._msgId === 'string' ? payload._msgId : undefined;
          handleIncomingState(payload, msgId);
        }
      )
      .subscribe((status) => {
        if (destroyed) {
          return;
        }
        if (status === 'SUBSCRIBED') {
          supabaseSubscribed = true;
          supabaseErrored = false;
          void supabaseChannel
            .track({ role: 'controller', clientId, joinedAt: Date.now() })
            .catch(() => {});
          emitStatus();
          sendCommand({ action: 'requestState' });
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
          supabaseSubscribed = false;
          supabaseErrored = true;
          emitStatus();
        } else if (status === 'CLOSED') {
          supabaseSubscribed = false;
          emitStatus();
        }
      });
  }

  emitStatus();

  return {
    sendCommand,
    destroy() {
      if (destroyed) {
        return;
      }
      destroyed = true;
      if (bc) {
        postToBroadcastChannelSafe(bc, {
          kind: 'peer_leave',
          role: 'controller',
          clientId,
          sentAt: Date.now(),
        });
        try {
          bc.close();
        } catch {
          // Ignore close errors
        }
      }
      if (supabaseChannel && supabaseClient) {
        void supabaseChannel.untrack().catch(() => {});
        void supabaseClient.removeChannel(supabaseChannel).catch(() => {});
      }
    },
  };
}
