import { useRpc, type PluginHostProps } from '@getpaseo/plugin/client';
import { Modal, ScrollView, TextInput } from '@getpaseo/plugin/client/react-native';
import { useQuery } from '@tanstack/react-query';
import { useEffect, useRef, useState } from 'react';
import { Image, Text, View, type ScrollView as NativeScrollView } from 'react-native';
import { readPreview, type Block, type Preview, type Run, type Selection } from '../shared/contracts';
import { Button } from './controls';
import { Workbook } from './workbook';
export { Button } from './controls';
type Theme = PluginHostProps['theme'];

function Runs({ runs, theme, heading, scale = 1 }: { runs: Run[]; theme: Theme; heading?: number; scale?: number }) {
  return <Text selectable accessibilityRole={heading ? 'header' : undefined} style={{ color: theme.colors.foreground, fontSize: (heading ? Math.max(17, 29 - heading * 2) : 16) * scale, lineHeight: (heading ? 36 : 26) * scale, fontWeight: heading ? '700' : '400' }}>
    {runs.map((run, i) => <Text key={i} style={{ fontWeight: run.bold ? '700' : undefined, fontStyle: run.italic ? 'italic' : undefined }}>{run.text}</Text>)}
  </Text>;
}
export function DocumentBlock({ block, theme, scale = 1 }: { block: Block; theme: Theme; scale?: number }) {
  const [width, setWidth] = useState(760);
  if (block.kind === 'paragraph') return <View style={{ paddingVertical: block.heading ? 14 : 8 }}><Runs runs={block.runs} theme={theme} heading={block.heading} scale={scale} /></View>;
  if (block.kind === 'image') return <Image accessibilityLabel={block.alt} source={{ uri: block.uri }} resizeMode="contain" style={{ width: '100%', height: 300 * scale, marginVertical: 8 }} />;
  const columns = Math.max(0, ...block.rows.map(row => row.length));
  return <View onLayout={event => setWidth(event.nativeEvent.layout.width)}><ScrollView horizontal style={{ marginVertical: 12 }}>
    <View>{block.rows.map((row, r) => <View key={r} style={{ flexDirection: 'row', backgroundColor: r % 2 ? theme.colors.surface1 : theme.colors.surface0 }}>
      {row.map((runs, c) => <View key={c} style={{ width: Math.max(140 * scale, width / Math.max(1, columns)), padding: 12, borderWidth: 0.5, borderColor: theme.colors.border }}><Runs runs={runs} theme={theme} scale={scale} /></View>)}
    </View>)}</View>
  </ScrollView></View>;
}
function Warnings({ warnings, theme }: { warnings: string[]; theme: Theme }) {
  const [expanded, setExpanded] = useState(false);
  if (!warnings.length) return null;
  return <View style={{ paddingHorizontal: 8, paddingVertical: 4 }}>
    <Button label={expanded ? 'Hide preview notes' : `Preview notes (${warnings.length})`} onPress={() => setExpanded(!expanded)} theme={theme} />
    {expanded && warnings.map(warning => <Text key={warning} style={{ color: theme.colors.foregroundMuted, fontSize: 12, paddingVertical: 4 }}>{warning}</Text>)}
  </View>;
}
function blockText(block: Block): string {
  return block.kind === 'paragraph' ? block.runs.map(run => run.text).join('') : block.kind === 'table' ? block.rows.flat().map(runs => runs.map(run => run.text).join('')).join(' ') : block.alt;
}
function Word({ data, theme }: { data: Extract<Preview, { kind: 'document' }>; theme: Theme }) {
  const [scale, setScale] = useState(1);
  const [search, setSearch] = useState('');
  const [contents, setContents] = useState(false);
  const scroll = useRef<NativeScrollView>(null);
  const positions = useRef(new Map<number, number>());
  const visible = data.blocks.map((block, index) => ({ block, index })).filter(({ block }) => !search.trim() || blockText(block).toLowerCase().includes(search.trim().toLowerCase()));
  return <View style={{ flex: 1, minHeight: 0, backgroundColor: theme.colors.surface1 }}>
    <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: 6, padding: 8 }}>
      <Button label="Contents" selected={contents} theme={theme} onPress={() => { setSearch(''); setContents(!contents); }} />
      <Button label="Smaller text" theme={theme} disabled={scale <= 0.8} onPress={() => setScale(Math.max(0.8, scale - 0.1))} />
      <Button label="Larger text" theme={theme} disabled={scale >= 1.6} onPress={() => setScale(Math.min(1.6, scale + 0.1))} />
      <TextInput accessibilityLabel="Find in document" placeholder="Find in document" placeholderTextColor={theme.colors.foregroundMuted} value={search} onChangeText={setSearch} style={{ color: theme.colors.foreground, borderWidth: 1, borderColor: theme.colors.border, borderRadius: 8, padding: 8, minWidth: 160, flexGrow: 1 }} />
      {search && <Button label="Clear search" theme={theme} onPress={() => setSearch('')} />}
    </View>
    {search ? <Text accessibilityLiveRegion="polite" style={{ padding: 8, color: theme.colors.foregroundMuted }}>{visible.length} matching sections</Text> : null}
    {contents && <ScrollView style={{ maxHeight: 180, flexGrow: 0 }} contentContainerStyle={{ padding: 8, gap: 4 }}>
      {data.blocks.map((block, index) => block.kind === 'paragraph' && block.heading ? <Button key={index} label={blockText(block)} theme={theme} onPress={() => scroll.current?.scrollTo({ y: positions.current.get(index) ?? 0, animated: true })} /> : null)}
      {!data.blocks.some(block => block.kind === 'paragraph' && block.heading) && <Text style={{ color: theme.colors.foregroundMuted }}>No headings in this document.</Text>}
    </ScrollView>}
    <ScrollView ref={scroll} style={{ flex: 1 }} contentContainerStyle={{ padding: 16 }}>
      <View style={{ width: '100%', maxWidth: 900, alignSelf: 'center', padding: 20, backgroundColor: theme.colors.surface0, borderWidth: 1, borderColor: theme.colors.border }}>
        {visible.map(({ block, index }) => <View key={index} onLayout={event => positions.current.set(index, event.nativeEvent.layout.y)}><DocumentBlock block={block} theme={theme} scale={scale} /></View>)}
        {!visible.length && <Text style={{ color: theme.colors.foregroundMuted }}>{search ? 'No matching text.' : 'No readable content was found.'}</Text>}
      </View>
    </ScrollView>
  </View>;
}
function Pdf({ data, theme, onSelection }: { data: Extract<Preview, { kind: 'pdf' }>; theme: Theme; onSelection(selection: Selection): void }) {
  const [width, setWidth] = useState(800);
  const [zoom, setZoom] = useState(1);
  const [textView, setTextView] = useState(false);
  const [pageInput, setPageInput] = useState(String(data.page));
  const [error, setError] = useState('');
  useEffect(() => setPageInput(String(data.page)), [data.page]);
  function go(page: number) {
    if (!Number.isInteger(page) || page < 1 || page > Math.min(data.pages, 10000)) { setError(`Enter a page from 1 to ${Math.min(data.pages, 10000)}.`); return; }
    setError(''); onSelection({ page });
  }
  return <View style={{ flex: 1, minHeight: 0 }} onLayout={event => setWidth(event.nativeEvent.layout.width)}>
    <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: 6, padding: 8, alignItems: 'center' }}>
      <Button label="Previous page" disabled={data.page === 1} theme={theme} onPress={() => go(data.page - 1)} />
      <TextInput accessibilityLabel="Page number" value={pageInput} onChangeText={setPageInput} onSubmitEditing={() => go(Number(pageInput))} keyboardType="number-pad" style={{ color: theme.colors.foreground, borderWidth: 1, borderColor: theme.colors.border, padding: 8, width: 60, minHeight: 44 }} />
      <Text style={{ color: theme.colors.foregroundMuted }}>of {data.pages}</Text>
      <Button label="Go to page" theme={theme} onPress={() => go(Number(pageInput))} />
      <Button label="Next page" disabled={data.page >= Math.min(data.pages, 10000)} theme={theme} onPress={() => go(data.page + 1)} />
      <Button label="Fit width" selected={zoom === 1} theme={theme} onPress={() => setZoom(1)} />
      <Button label="Zoom out" disabled={zoom <= 0.5} theme={theme} onPress={() => setZoom(Math.max(0.5, zoom - 0.25))} />
      <Button label="Zoom in" disabled={zoom >= 2} theme={theme} onPress={() => setZoom(Math.min(2, zoom + 0.25))} />
      <Button label={textView ? 'Page view' : 'Selectable text'} theme={theme} selected={textView} onPress={() => setTextView(!textView)} />
    </View>
    {error ? <Text accessibilityRole="alert" style={{ padding: 8, color: theme.colors.foreground }}>{error}</Text> : null}
    {textView ? <ScrollView style={{ flex: 1 }} contentContainerStyle={{ padding: 24 }}><Text selectable style={{ fontSize: 16 * zoom, lineHeight: 26 * zoom, color: theme.colors.foreground }}>{data.text || 'No selectable text on this page. Use Page view to read the scanned image.'}</Text></ScrollView>
      : <ScrollView horizontal style={{ flex: 1 }} contentContainerStyle={{ minWidth: '100%', justifyContent: 'center' }}><ScrollView key={data.page} style={{ width: Math.max(1, width - 24) * zoom }}><Image accessibilityLabel={`PDF page ${data.page} of ${data.pages}`} source={{ uri: data.uri }} resizeMode="contain" style={{ width: Math.max(1, width - 24) * zoom, height: Math.max(1, width - 24) * zoom * data.height / data.width }} /></ScrollView></ScrollView>}
  </View>;
}
export function PreviewContent({ workspaceId, path, host, theme }: { workspaceId: string; path: string } & PluginHostProps) {
  const read = useRpc(readPreview);
  const [selection, setSelection] = useState<Selection>({ sheet: 0, rowStart: 0, columnStart: 0, page: 1 });
  const query = useQuery({ queryKey: ['office-preview', host.id, workspaceId, path, selection], queryFn: () => read({ workspaceId, path, ...selection }), placeholderData: previous => previous, retry: false, gcTime: 0, staleTime: 0, refetchOnWindowFocus: false });
  if (query.isPending) return <Text style={{ padding: 20, color: theme.colors.foregroundMuted }}>Opening document…</Text>;
  if (query.isError) return <View style={{ padding: 20, gap: 12 }}><Text accessibilityRole="alert" style={{ color: theme.colors.foreground }}>{query.error.message}</Text><Button label="Try again" theme={theme} onPress={() => void query.refetch()} /></View>;
  const data = query.data;
  return <View style={{ flex: 1, minHeight: 0 }}>
    {query.isFetching && <Text accessibilityLiveRegion="polite" style={{ color: theme.colors.foregroundMuted, padding: 4 }}>Loading page…</Text>}
    {data.kind === 'document' ? <Word data={data} theme={theme} /> : data.kind === 'pdf' ? <Pdf data={data} theme={theme} onSelection={value => { if (!query.isFetching) setSelection(value); }} /> : <Workbook data={data} theme={theme} onSelection={value => { if (!query.isFetching) setSelection(value); }} />}
    <Warnings warnings={data.warnings} theme={theme} />
  </View>;
}
export function PreviewModal({ path, workspaceId, onClose, ...hostProps }: { path: string | null; workspaceId: string; onClose(): void } & PluginHostProps) {
  return <Modal title={path?.split(/[\\/]/).pop() ?? 'Document preview'} open={path !== null} onOpenChange={open => { if (!open) onClose(); }}>
    <Modal.Content scrollable={false} style={{ backgroundColor: hostProps.theme.colors.surface0 }} contentContainerStyle={{ padding: 0, gap: 0 }}>
      {path && <PreviewContent key={`${workspaceId}:${path}`} path={path} workspaceId={workspaceId} {...hostProps} />}
    </Modal.Content>
  </Modal>;
}
