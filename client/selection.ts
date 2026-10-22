import { useSyncExternalStore } from 'react';

const paths = new Map<string, string>();
const listeners = new Set<() => void>();
const subscribe = (listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener); }; };
export function selectPreview(hostId: string, workspaceId: string, path: string | null) {
  const key = `${hostId}:${workspaceId}`;
  if (path === null) paths.delete(key); else paths.set(key, path);
  listeners.forEach(listener => listener());
}
export function useSelectedPreview(hostId: string, workspaceId: string) {
  return useSyncExternalStore(subscribe, () => paths.get(`${hostId}:${workspaceId}`) ?? null);
}
