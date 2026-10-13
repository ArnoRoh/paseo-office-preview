import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { useRpc, type PluginWorkspacePanelProps, type PluginTimelineItemProps } from '@getpaseo/plugin/client';
import { FlatList, TextInput } from '@getpaseo/plugin/client/react-native';
import { Text, View } from 'react-native';
import { listFiles } from '../shared/contracts';
import { isOfficePath } from '../shared/links';
import { Button, PreviewModal } from './viewer';

export function OfficePanel(props: PluginWorkspacePanelProps) {
  const [path, setPath] = useState('');
  const [selected, setSelected] = useState<string | null>(null);
  const list = useRpc(listFiles);
  const files = useQuery({ queryKey: ['office-files', props.host.id, props.workspaceId], queryFn: () => list({ workspaceId: props.workspaceId }), retry: false, gcTime: 0 });
  const cleanPath = path.trim().replace(/^['"]|['"]$/g, '');
  return <View style={{ flex: 1, padding: props.layout.compact ? 12 : 20, gap: 12, backgroundColor: props.theme.colors.surface0 }}>
    <Text accessibilityRole="header" style={{ color: props.theme.colors.foreground, fontSize: 22, fontWeight: '700' }}>Office preview</Text>
    <Text style={{ color: props.theme.colors.foregroundMuted }}>Choose a Word or Excel file in this workspace, or paste its path.</Text>
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
      ListEmptyComponent={!files.isPending && !files.isError ? <Text style={{ color: props.theme.colors.foregroundMuted }}>No .docx, .xlsx or .xls files found.</Text> : null}
      renderItem={({ item }) => <Button label={item} theme={props.theme} onPress={() => setSelected(item)} />} />
    <PreviewModal {...props} path={selected} onClose={() => setSelected(null)} />
  </View>;
}
export function OfficeCard(props: PluginTimelineItemProps<{ workspaceId: string; paths: string[] }>) {
  const [selected, setSelected] = useState<string | null>(null);
  return <View style={{ padding: 12, gap: 8 }}>
    <Text style={{ color: props.theme.colors.foregroundMuted }}>Read documents</Text>
    {props.item.data.paths.map(path => <Button key={path} label={`Preview ${path.split(/[\\/]/).pop()}`} theme={props.theme} onPress={() => setSelected(path)} />)}
    <PreviewModal {...props} workspaceId={props.item.data.workspaceId} path={selected} onClose={() => setSelected(null)} />
  </View>;
}
