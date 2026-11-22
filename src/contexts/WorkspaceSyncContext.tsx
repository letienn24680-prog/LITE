import { createContext, useContext } from 'react';
import type { ReactNode } from 'react';

export type DriveSyncBadgeStatus =
  | 'idle'
  | 'saving'
  | 'saved'
  | 'offline'
  | 'error';

export interface WorkspaceSyncContextValue {
  driveSyncBadge: DriveSyncBadgeStatus;
  driveSyncLabel: string;
  pendingOfflineCount: number;
  onOpenVersionHistory?: () => void;
  onOpenQuickConverter?: () => void;
  onSaveVersionCheckpoint?: () => Promise<void> | void;
}

const DEFAULT_WORKSPACE_SYNC_VALUE: WorkspaceSyncContextValue = {
  driveSyncBadge: 'idle',
  driveSyncLabel: '',
  pendingOfflineCount: 0,
};

const WorkspaceSyncContext = createContext<WorkspaceSyncContextValue>(
  DEFAULT_WORKSPACE_SYNC_VALUE
);

export interface WorkspaceSyncProviderProps {
  value: WorkspaceSyncContextValue;
  children: ReactNode;
}

export function WorkspaceSyncProvider({
  value,
  children,
}: WorkspaceSyncProviderProps) {
  return (
    <WorkspaceSyncContext.Provider value={value}>
      {children}
    </WorkspaceSyncContext.Provider>
  );
}

export function useWorkspaceSync(): WorkspaceSyncContextValue {
  return useContext(WorkspaceSyncContext);
}
