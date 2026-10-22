import type { PluginHostProps } from '@getpaseo/plugin/client';
import { Pressable, Text } from 'react-native';
type Theme = PluginHostProps['theme'];
export function Button({ label, onPress, theme, disabled = false, selected = false }: {
  label: string; onPress(): void; theme: Theme; disabled?: boolean; selected?: boolean;
}) {
  return <Pressable accessibilityRole="button" accessibilityLabel={label} accessibilityState={{ disabled, selected }} disabled={disabled} onPress={onPress}
    style={{ paddingHorizontal: 12, paddingVertical: 10, minHeight: 44, borderRadius: 8, opacity: disabled ? 0.4 : 1, backgroundColor: selected ? theme.colors.accent : theme.colors.surface2 }}>
    <Text style={{ color: selected ? theme.colors.accentForeground : theme.colors.foreground, fontSize: 14 }}>{label}</Text>
  </Pressable>;
}
