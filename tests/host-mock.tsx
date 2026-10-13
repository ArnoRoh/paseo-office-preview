// Browser test adapter for Paseo-owned UI. The shipped plugin uses the real host.
import { useCallback, useEffect } from 'react';
import { createPortal } from 'react-dom';
import { FlatList, ScrollView, TextInput, View, Text, Pressable } from 'react-native';
export { FlatList, ScrollView, TextInput };
export function useRpc(contract: { name: string }) {
  return useCallback(async (input: unknown) => {
    const response = await fetch(`/rpc/${contract.name}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(input) });
    const body = await response.json();
    if (!response.ok) throw new Error(body.error);
    return body;
  }, [contract.name]);
}
export function Modal({ title, open, onOpenChange, children }: any) {
  useEffect(() => {
    const key = (event: KeyboardEvent) => { if (event.key === 'Escape') onOpenChange(false); };
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  }, [onOpenChange]);
  if (!open) return null;
  return createPortal(<div role="dialog" aria-label={title} style={{ position: 'fixed', inset: 0, background: '#0008', display: 'flex', alignItems: window.innerWidth < 600 ? 'flex-end' : 'center', justifyContent: 'center', zIndex: 100 }}>
    <View style={{ width: '100%', maxWidth: 960, height: '85%', backgroundColor: '#fafafa', borderRadius: 16, overflow: 'hidden' }}>
      <View style={{ padding: 16, flexDirection: 'row', justifyContent: 'space-between', backgroundColor: '#e7e7e7' }}><Text style={{ fontSize: 18, color: '#111' }}>{title}</Text><Pressable accessibilityRole="button" accessibilityLabel="Close preview" onPress={() => onOpenChange(false)}><Text style={{ color: '#111' }}>Close</Text></Pressable></View>
      {children}
    </View>
  </div>, document.body);
}
Modal.Content = ({ children, style }: any) => <View style={[{ flex: 1, minHeight: 0 }, style]}>{children}</View>;
