import { ScrollView, TextInput } from '@getpaseo/plugin/client/react-native';
import type { PluginHostProps } from '@getpaseo/plugin/client';
import { useEffect, useState } from 'react';
import { Pressable, Text, View } from 'react-native';
import { columnLabel, limits, type Preview, type Selection } from '../shared/contracts';
import { Button } from './controls';

type WorkbookData = Extract<Preview, { kind: 'workbook' }>;
export function Workbook({ data, theme, onSelection }: { data: WorkbookData; theme: PluginHostProps['theme']; onSelection(selection: Selection): void }) {
  const [availableWidth, setAvailableWidth] = useState(900);
  const [zoom, setZoom] = useState<number | null>(1);
  const [selected, setSelected] = useState<{ address: string; text: string } | null>(null);
  const [address, setAddress] = useState('');
  const [error, setError] = useState('');
  useEffect(() => setSelected(null), [data.sheet, data.rowStart, data.columnStart]);
  const count = data.columnWidths.length;
  const naturalWidth = data.columnWidths.reduce((a, b) => a + b, 0);
  const scale = zoom ?? Math.max(0.55, Math.min(1, (availableWidth - 46) / Math.max(1, naturalWidth)));
  const widths = data.columnWidths.map(width => width * scale);
  const offsets = [46];
  widths.forEach(width => offsets.push(offsets[offsets.length - 1]! + width));
  const fontSize = Math.max(11, 14 * scale);
  const lineHeight = Math.ceil(fontSize * 1.4);
  const cells: { row: number; column: number; rowSpan: number; columnSpan: number; text: string }[] = [];
  const covered = new Set<string>();
  const anchors = new Map<string, WorkbookData['merges'][number]>();
  for (const merge of data.merges) {
    anchors.set(`${merge.row}:${merge.column}`, merge);
    for (let r = merge.row; r < merge.row + merge.rowSpan; r++) for (let c = merge.column; c < merge.column + merge.columnSpan; c++) covered.add(`${r}:${c}`);
  }
  data.rows.forEach((row, r) => row.forEach((text, c) => {
    const key = `${r}:${c}`, merge = anchors.get(key);
    if (merge) cells.push(merge);
    else if (!covered.has(key)) cells.push({ row: r, column: c, rowSpan: 1, columnSpan: 1, text });
  }));
  const heights = data.rows.map(() => 36);
  for (const cell of cells) {
    const width = offsets[cell.column + cell.columnSpan]! - offsets[cell.column]! - 16;
    const lines = cell.text.split('\n').reduce((n, line) => n + Math.max(1, Math.ceil(line.length * fontSize * 0.54 / width)), 0);
    const required = Math.min(260, Math.max(36, lines * lineHeight + 16));
    const existing = heights.slice(cell.row, cell.row + cell.rowSpan).reduce((a, b) => a + b, 0);
    if (required > existing) heights[cell.row + cell.rowSpan - 1]! += required - existing;
  }
  const tops = [0]; heights.forEach(height => tops.push(tops[tops.length - 1]! + height));
  const selection = { sheet: data.sheet, rowStart: data.rowStart, columnStart: data.columnStart };
  function go() {
    const match = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(address.trim());
    const column = match ? [...match[1]!.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1 : -1;
    const row = match ? Number(match[2]) - 1 : -1;
    if (row < 0 || column < 0 || row >= data.totalRows || column >= data.totalColumns) { setError('Enter a cell within this worksheet, such as A1.'); return; }
    setError(''); setSelected(null);
    onSelection({ sheet: data.sheet, rowStart: Math.floor(row / limits.pageRows) * limits.pageRows, columnStart: Math.floor(column / limits.pageColumns) * limits.pageColumns });
  }
  return <View style={{ flex: 1, minHeight: 0 }} onLayout={event => setAvailableWidth(event.nativeEvent.layout.width)}>
    <View style={{ padding: 8, gap: 6, flexDirection: 'row', flexWrap: 'wrap', alignItems: 'center' }}>
      <Button label="Fit width" selected={zoom === null} theme={theme} onPress={() => setZoom(null)} />
      <Button label="Zoom out" theme={theme} disabled={scale <= 0.55} onPress={() => setZoom(Math.max(0.55, scale - 0.15))} />
      <Text style={{ color: theme.colors.foregroundMuted }}>{Math.round(scale * 100)}%</Text>
      <Button label="Zoom in" theme={theme} disabled={scale >= 1.75} onPress={() => setZoom(Math.min(1.75, scale + 0.15))} />
      <TextInput accessibilityLabel="Go to cell" placeholder="A1" placeholderTextColor={theme.colors.foregroundMuted} value={address} onChangeText={setAddress} onSubmitEditing={go} autoCapitalize="characters" autoCorrect={false} style={{ color: theme.colors.foreground, borderWidth: 1, borderColor: theme.colors.border, padding: 8, width: 90, minHeight: 44 }} />
      <Button label="Go to cell" theme={theme} onPress={go} />
    </View>
    {error ? <Text accessibilityRole="alert" style={{ color: theme.colors.foreground, padding: 8 }}>{error}</Text> : null}
    {selected && <View style={{ padding: 8, borderBottomWidth: 1, borderColor: theme.colors.border, maxHeight: 140 }}><Text style={{ color: theme.colors.foregroundMuted }}>{selected.address}</Text><ScrollView><Text selectable style={{ color: theme.colors.foreground }}>{selected.text || 'Empty cell'}</Text></ScrollView></View>}
    <ScrollView horizontal style={{ flex: 1, minHeight: 100 }} contentContainerStyle={{ minHeight: '100%' }}>
      <ScrollView key={`${data.sheet}:${data.rowStart}:${data.columnStart}`} stickyHeaderIndices={[0]} style={{ width: offsets[count], flex: 1 }}>
        <View style={{ flexDirection: 'row', backgroundColor: theme.colors.surface2 }}>
          <Text style={{ width: 46, padding: 8, color: theme.colors.foregroundMuted }}>#</Text>
          {widths.map((width, c) => <Text key={c} style={{ width, padding: 8, color: theme.colors.foreground, fontWeight: '700' }}>{columnLabel(data.columnStart + c)}</Text>)}
        </View>
        <View style={{ height: tops[data.rows.length], width: offsets[count] }}>
          {data.rows.map((_, r) => <Text key={r} style={{ position: 'absolute', top: tops[r], left: 0, width: 46, height: heights[r], padding: 8, fontSize: 11, color: theme.colors.foregroundMuted, backgroundColor: theme.colors.surface2 }}>{data.rowStart + r + 1}</Text>)}
          {cells.map(cell => {
            const name = `${columnLabel(data.columnStart + cell.column)}${data.rowStart + cell.row + 1}`;
            return <Pressable key={`${cell.row}:${cell.column}`} accessibilityRole="button" accessibilityLabel={`${name}: ${cell.text || 'Empty cell'}`} onPress={() => setSelected({ address: name, text: cell.text })}
              style={{ position: 'absolute', top: tops[cell.row], left: offsets[cell.column], width: offsets[cell.column + cell.columnSpan]! - offsets[cell.column]!, height: tops[cell.row + cell.rowSpan]! - tops[cell.row]!, padding: 8, overflow: 'hidden', borderWidth: 0.5, borderColor: selected?.address === name ? theme.colors.accent : theme.colors.border, backgroundColor: cell.row % 2 ? theme.colors.surface1 : theme.colors.surface0 }}>
              <Text numberOfLines={Math.max(1, Math.floor((tops[cell.row + cell.rowSpan]! - tops[cell.row]! - 16) / lineHeight))} style={{ fontSize, lineHeight, color: theme.colors.foreground, textAlign: data.numeric[cell.row]?.[cell.column] && cell.columnSpan === 1 ? 'right' : 'left' }}>{cell.text}</Text>
            </Pressable>;
          })}
        </View>
      </ScrollView>
    </ScrollView>
    <View style={{ padding: 8, gap: 6, flexDirection: 'row', flexWrap: 'wrap', alignItems: 'center', borderTopWidth: 1, borderColor: theme.colors.border }}>
      <Text style={{ color: theme.colors.foregroundMuted, fontSize: 12 }}>{data.totalRows ? `Rows ${data.rowStart + 1}–${data.rowStart + data.rows.length} of ${data.totalRows} · Columns ${columnLabel(data.columnStart)}–${columnLabel(data.columnStart + count - 1)}` : 'Empty worksheet'} · Select a cell to read and copy its value</Text>
      {data.totalRows > limits.pageRows && <><Button label="Previous rows" disabled={!data.rowStart} theme={theme} onPress={() => onSelection({ ...selection, rowStart: Math.max(0, data.rowStart - limits.pageRows) })} /><Button label="Next rows" disabled={data.rowStart + limits.pageRows >= data.totalRows} theme={theme} onPress={() => onSelection({ ...selection, rowStart: data.rowStart + limits.pageRows })} /></>}
      {data.totalColumns > limits.pageColumns && <><Button label="Previous columns" disabled={!data.columnStart} theme={theme} onPress={() => onSelection({ ...selection, columnStart: Math.max(0, data.columnStart - limits.pageColumns) })} /><Button label="Next columns" disabled={data.columnStart + limits.pageColumns >= data.totalColumns} theme={theme} onPress={() => onSelection({ ...selection, columnStart: data.columnStart + limits.pageColumns })} /></>}
    </View>
    <ScrollView horizontal style={{ flexGrow: 0, flexShrink: 0 }} contentContainerStyle={{ padding: 8, gap: 6 }}>
      {data.sheets.map((sheet, index) => <Button key={index} label={sheet} selected={index === data.sheet} theme={theme} onPress={() => { setSelected(null); onSelection({ sheet: index, rowStart: 0, columnStart: 0 }); }} />)}
    </ScrollView>
  </View>;
}
