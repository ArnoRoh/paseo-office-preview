import { useRpc, type PluginHostProps } from '@getpaseo/plugin/client';
import { FlatList, Modal, ScrollView } from '@getpaseo/plugin/client/react-native';
import { useQuery } from '@tanstack/react-query';
import { useState } from 'react';
import { Image, Pressable, Text, View } from 'react-native';
import { columnLabel, limits, readPreview, type Block, type Preview, type Run, type Selection } from '../shared/contracts';

type Theme = PluginHostProps['theme'];
export function Button({ label, onPress, theme, disabled = false, selected = false }: {
  label: string; onPress(): void; theme: Theme; disabled?: boolean; selected?: boolean;
}) {
  return <Pressable accessibilityRole="button" accessibilityLabel={label} accessibilityState={{ disabled, selected }} disabled={disabled} onPress={onPress}
    style={{ paddingHorizontal: 12, paddingVertical: 10, minHeight: 44, borderRadius: 8, opacity: disabled ? 0.4 : 1, backgroundColor: selected ? theme.colors.accent : theme.colors.surface2 }}>
    <Text style={{ color: selected ? theme.colors.accentForeground : theme.colors.foreground, fontSize: 14 }}>{label}</Text>
  </Pressable>;
}
function Runs({ runs, theme, heading }: { runs: Run[]; theme: Theme; heading?: number }) {
  return <Text selectable accessibilityRole={heading ? 'header' : undefined} style={{ color: theme.colors.foreground, fontSize: heading ? Math.max(17, 27 - heading * 2) : 16, lineHeight: heading ? 30 : 25, fontWeight: heading ? '700' : '400' }}>
    {runs.map((run, i) => <Text key={i} style={{ fontWeight: run.bold ? '700' : undefined, fontStyle: run.italic ? 'italic' : undefined }}>{run.text}</Text>)}
  </Text>;
}
export function DocumentBlock({ block, theme }: { block: Block; theme: Theme }) {
  if (block.kind === 'paragraph') return <View style={{ paddingVertical: 8 }}><Runs runs={block.runs} theme={theme} heading={block.heading} /></View>;
  if (block.kind === 'image') return <Image accessibilityLabel={block.alt} source={{ uri: block.uri }} resizeMode="contain" style={{ width: '100%', height: 260, marginVertical: 8 }} />;
  return <ScrollView horizontal style={{ marginVertical: 8 }}>
    <View>{block.rows.map((row, r) => <View key={r} style={{ flexDirection: 'row' }}>
      {row.map((runs, c) => <View key={c} style={{ width: 200, padding: 8, borderWidth: 0.5, borderColor: theme.colors.border }}><Runs runs={runs} theme={theme} /></View>)}
    </View>)}</View>
  </ScrollView>;
}
function Warnings({ warnings, theme }: { warnings: string[]; theme: Theme }) {
  return <View style={{ padding: 12, gap: 4 }}>
    {warnings.map(warning => <Text key={warning} style={{ color: theme.colors.foregroundMuted, fontSize: 12 }}>{warning}</Text>)}
  </View>;
}
export function Workbook({ data, theme, onSelection }: { data: Extract<Preview, { kind: 'workbook' }>; theme: Theme; onSelection(selection: Selection): void }) {
  const selection = { sheet: data.sheet, rowStart: data.rowStart, columnStart: data.columnStart };
  const width = Math.min(limits.pageColumns, data.totalColumns - data.columnStart);
  return <View style={{ flex: 1, minHeight: 0 }}>
    <ScrollView horizontal style={{ flexGrow: 0, flexShrink: 0 }} contentContainerStyle={{ paddingHorizontal: 12, gap: 6 }}>
      {data.sheets.map((sheet, index) => <Button key={`${index}-${sheet}`} label={sheet} selected={index === data.sheet} theme={theme} onPress={() => onSelection({ sheet: index, rowStart: 0, columnStart: 0 })} />)}
    </ScrollView>
    <View style={{ flexDirection: 'row', flexWrap: 'wrap', alignItems: 'center', gap: 6, padding: 12 }}>
      <Button label="Previous rows" disabled={data.rowStart === 0} theme={theme} onPress={() => onSelection({ ...selection, rowStart: Math.max(0, data.rowStart - limits.pageRows) })} />
      <Button label="Next rows" disabled={data.rowStart + limits.pageRows >= data.totalRows} theme={theme} onPress={() => onSelection({ ...selection, rowStart: data.rowStart + limits.pageRows })} />
      <Button label="Previous columns" disabled={data.columnStart === 0} theme={theme} onPress={() => onSelection({ ...selection, columnStart: Math.max(0, data.columnStart - limits.pageColumns) })} />
      <Button label="Next columns" disabled={data.columnStart + limits.pageColumns >= data.totalColumns} theme={theme} onPress={() => onSelection({ ...selection, columnStart: data.columnStart + limits.pageColumns })} />
      <Text style={{ color: theme.colors.foregroundMuted, fontSize: 12 }}>{data.totalRows ? `Rows ${data.rowStart + 1}–${Math.min(data.rowStart + limits.pageRows, data.totalRows)} of ${data.totalRows} · Columns ${columnLabel(data.columnStart)}–${columnLabel(data.columnStart + width - 1)}` : 'Empty worksheet'}</Text>
    </View>
    <ScrollView horizontal style={{ flex: 1, minHeight: 80 }} contentContainerStyle={{ minHeight: '100%' }}>
      <View style={{ width: 52 + Math.max(1, width) * 160, flex: 1 }}>
        <FlatList data={data.rows} keyExtractor={(_, index) => String(data.rowStart + index)} initialNumToRender={15}
          ListHeaderComponent={<View style={{ flexDirection: 'row', backgroundColor: theme.colors.surface2 }}>
            <Text style={{ width: 52, padding: 8, color: theme.colors.foregroundMuted }}>#</Text>
            {Array.from({ length: width }, (_, c) => <Text key={c} style={{ width: 160, padding: 8, color: theme.colors.foreground, fontWeight: '700' }}>{columnLabel(data.columnStart + c)}</Text>)}
          </View>}
          stickyHeaderIndices={[0]}
          renderItem={({ item: row, index }) => <View style={{ flexDirection: 'row', backgroundColor: index % 2 ? theme.colors.surface1 : theme.colors.surface0 }}>
            <Text style={{ width: 52, padding: 8, color: theme.colors.foregroundMuted, fontSize: 12 }}>{data.rowStart + index + 1}</Text>
            {row.map((value, c) => <View key={c} style={{ width: 160, padding: 8, borderWidth: 0.5, borderColor: theme.colors.border }}><Text selectable style={{ color: theme.colors.foreground, fontSize: 14 }}>{value}</Text></View>)}
          </View>} />
      </View>
    </ScrollView>
  </View>;
}
export function PreviewContent({ workspaceId, path, host, theme }: { workspaceId: string; path: string } & PluginHostProps) {
  const read = useRpc(readPreview);
  const [selection, setSelection] = useState<Selection>({ sheet: 0, rowStart: 0, columnStart: 0 });
  const query = useQuery({ queryKey: ['office-preview', host.id, workspaceId, path, selection], queryFn: () => read({ workspaceId, path, ...selection }), retry: false, gcTime: 0, staleTime: 0, refetchOnWindowFocus: false });
  if (query.isPending) return <Text accessibilityRole="text" style={{ padding: 20, color: theme.colors.foregroundMuted }}>Opening document…</Text>;
  if (query.isError) return <View style={{ padding: 20, gap: 12 }}>
    <Text accessibilityRole="alert" style={{ color: theme.colors.foreground }}>{query.error.message}</Text>
    <Button label="Try again" theme={theme} onPress={() => void query.refetch()} />
  </View>;
  const data = query.data;
  return <View style={{ flex: 1, minHeight: 0 }}>
    <Warnings warnings={data.warnings} theme={theme} />
    {data.kind === 'document' ? <FlatList data={data.blocks} keyExtractor={(_, index) => String(index)} initialNumToRender={12} contentContainerStyle={{ paddingHorizontal: 20, paddingBottom: 24 }}
      ListEmptyComponent={<Text style={{ color: theme.colors.foregroundMuted }}>No readable content was found.</Text>}
      renderItem={({ item }) => <DocumentBlock block={item} theme={theme} />} />
      : <Workbook data={data} theme={theme} onSelection={setSelection} />}
  </View>;
}
export function PreviewModal({ path, workspaceId, onClose, ...hostProps }: { path: string | null; workspaceId: string; onClose(): void } & PluginHostProps) {
  return <Modal title={path?.split(/[\\/]/).pop() ?? 'Office preview'} open={path !== null} onOpenChange={open => { if (!open) onClose(); }}>
    <Modal.Content scrollable={false} style={{ backgroundColor: hostProps.theme.colors.surface0 }} contentContainerStyle={{ padding: 0, gap: 0 }}>
      {path && <PreviewContent key={`${workspaceId}:${path}`} path={path} workspaceId={workspaceId} {...hostProps} />}
    </Modal.Content>
  </Modal>;
}
