import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { useRpc, type PluginWorkspacePanelProps, type PluginTimelineItemProps } from '@getpaseo/plugin/client';
import { FlatList, TextInput } from '@getpaseo/plugin/client/react-native';
import { Text, View } from 'react-native';
import { listFiles } from '../shared/contracts';
import { isOfficePath } from '../shared/links';
import { selectPreview, useSelectedPreview } from './selection';
import { Button, PreviewContent, PreviewModal } from './viewer';

export function OfficePanel(props: PluginWorkspacePanelProps) {
  const [path, setPath] = useState('');
  const selected = useSelectedPreview(props.host.id, props.workspaceId);
  const setSelected = (path: string | null) => selectPreview(props.host.id, props.workspaceId, path);
  const list = useRpc(listFiles);
  const files = useQuery({ queryKey: ['office-files', props.host.id, props.workspaceId], queryFn: () => list({ workspaceId: props.workspaceId }), retry: false, gcTime: 0 });
  const cleanPath = path.trim().replace(/^['"]|['"]$/g, '');
  if (selected) return <View style={{ flex: 1, minHeight: 0, backgroundColor: props.theme.colors.surface0 }}>
    <View style={{ flexDirection: 'row', alignItems: 'center', padding: 8, gap: 12 }}>
      <Button label="Close preview" theme={props.theme} onPress={() => setSelected(null)} />
      <Text numberOfLines={1} style={{ flex: 1, color: props.theme.colors.foreground, fontWeight: '700' }}>{selected.split(/[\\/]/).pop()}</Text>
    </View>
    <PreviewContent key={`${props.workspaceId}:${selected}`} {...props} path={selected} />
  </View>;
  return <View style={{ flex: 1, padding: props.layout.compact ? 12 : 20, gap: 12, backgroundColor: props.theme.colors.surface0 }}>
    <Text accessibilityRole="header" style={{ color: props.theme.colors.foreground, fontSize: 22, fontWeight: '700' }}>Office preview</Text>
    <Text style={{ color: props.theme.colors.foregroundMuted }}>Choose a Word, Excel or PDF file in this workspace, or paste its path.</Text>
    <TextInput accessibilityLabel="Document path" placeholder="reports/example.docx" placeholderTextColor={props.theme.colors.foregroundMuted}
      autoCapitalize="none" autoCorrect={false} value={path} onChangeText={setPath} onSubmitEditing={() => { if (isOfficePath(cleanPath)) setSelected(cleanPath); }}
      style={{ color: props.theme.colors.foreground, padding: 12, borderColor: props.theme.colors.border, borderWidth: 1, borderRadius: 8 }} />
    <View style={{ flexDirection: 'row', gap: 8 }}>
      <Button label="Preview file" disabled={!isOfficePath(cleanPath)} theme={props.theme} onPress={() => setSelected(cleanPath)} />
      <Button label="Refresh files" theme={props.theme} disabled={files.isFetching} onPress={() => void files.refetch()} />
    </View>
    {files.isPending && <Text style={{ color: props.theme.colors.foregroundMuted }}>Looking for documents…</Text>}
    {files.isError && <Text accessibilityRole="alert" style={{ color: props.theme.colors.foreground }}>{files.error.message}</Text>}
    {files.data?.truncated && <Text style={{ color: props.theme.colors.foregroundMuted }}>Showing a limited file list. Paste a path to open another file.</Text>}
    <FlatList data={files.data?.paths ?? []} keyExtractor={value => value} contentContainerStyle={{ gap: 8 }}
      ListEmptyComponent={!files.isPending && !files.isError ? <Text style={{ color: props.theme.colors.foregroundMuted }}>No .docx, .xlsx, .xls or .pdf files found.</Text> : null}
      renderItem={({ item }) => <Button label={item} theme={props.theme} onPress={() => setSelected(item)} />} />
  </View>;
}
export function OfficeCard(props: PluginTimelineItemProps<{ workspaceId: string; paths: string[] }> & { openReader?: (workspaceId: string, path: string) => void }) {
  const [selected, setSelected] = useState<string | null>(null);
  return <View style={{ padding: 12, gap: 8 }}>
    <Text style={{ color: props.theme.colors.foregroundMuted }}>Read documents</Text>
    {props.item.data.paths.map(path => <Button key={path} label={`Preview ${path.split(/[\\/]/).pop()}`} theme={props.theme} onPress={() => { if (!props.layout.compact && props.openReader) props.openReader(props.item.data.workspaceId, path); else setSelected(path); }} />)}
    <PreviewModal {...props} workspaceId={props.item.data.workspaceId} path={selected} onClose={() => setSelected(null)} />
  </View>;
}
