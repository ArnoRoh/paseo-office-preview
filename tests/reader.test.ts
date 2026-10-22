import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, symlink, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { documentLinks } from '../shared/links';
import { columnLabel } from '../shared/contracts';
import { readDocument, discoverFiles } from '../server/files';
import { previewFile } from '../server/preview';
import { parseDocument } from '../server/reader';
import { checkZip } from '../server/zip';
import { wordFixture, workbookFixture, writeFixtures, pdfFixture, coverageFixture } from './fixtures';
import contribute from '../index.server';
import type { PluginServerContext, PluginLifecycleEvents } from '@getpaseo/plugin/server';
import type { PaseoApi } from '@getpaseo/client';

const first = { sheet: 0, rowStart: 0, columnStart: 0 };
test('Word reader preserves text emphasis and table data', async () => {
  const preview = await parseDocument(await wordFixture(), 'Reading.docx', first);
  assert.equal(preview.kind, 'document');
  if (preview.kind !== 'document') return;
  assert(preview.blocks.some(block => block.kind === 'paragraph' && block.runs.some(run => run.bold && run.text === 'Read only.')));
  assert(preview.blocks.some(block => block.kind === 'paragraph' && block.runs.some(run => run.italic)));
  assert(preview.blocks.some(block => block.kind === 'table' && block.rows[1]?.[1]?.[0]?.text === '24'));
});
for (const format of ['xlsx', 'biff8'] as const) test(`reads ${format} saved values, visible sheets and paged ranges`, async () => {
  const bytes = workbookFixture(format);
  const name = format === 'xlsx' ? 'test.xlsx' : 'test.xls';
  const preview = await parseDocument(bytes, name, first);
  assert.equal(preview.kind, 'workbook');
  if (preview.kind !== 'workbook') return;
  assert.deepEqual(preview.sheets, ['Inventory', 'Summary']);
  assert.equal(preview.rows.length, 100);
  assert.equal(preview.rows[1]?.[0], 'Sample item 1');
  assert.equal(preview.rows[1]?.[3], '2.50');
  const page = await parseDocument(bytes, name, { ...first, rowStart: 100 });
  assert(page.kind === 'workbook' && page.rows[0]?.[0] === 'Sample item 100');
  const columns = await parseDocument(bytes, name, { ...first, columnStart: 20 });
  assert(columns.kind === 'workbook' && columns.rows[0]?.[0] === 'Far column');
  const second = await parseDocument(bytes, name, { ...first, sheet: 1 });
  assert(second.kind === 'workbook' && second.rows[1]?.[0] === 'Saved values only');
});
test('file references handle spaces, encoded paths and line suffixes without remote links', () => {
  assert.deepEqual(documentLinks('[Word](</reports/Annual review.docx>) [Excel](report%20one.xlsx:12) `other.xls` [remote](https://example.com/leak.xlsx) [again](other.xls)'), ['/reports/Annual review.docx', 'report one.xlsx', 'other.xls']);
  assert.deepEqual(documentLinks('[remote](file://server/a.xlsx) [local](file:///tmp/a.docx) `javascript:bad.xlsx`'), ['/tmp/a.docx']);
  assert.equal(columnLabel(0), 'A'); assert.equal(columnLabel(26), 'AA'); assert.equal(columnLabel(255), 'IV');
});
test('workspace access blocks traversal, secrets, symlink escapes and oversized files', async () => {
  const base = await mkdtemp(path.join(os.tmpdir(), 'office-access-'));
  try {
    const root = path.join(base, 'workspace'); await mkdir(root);
    await writeFixtures(root);
    await writeFile(path.join(base, 'outside.docx'), await wordFixture());
    await symlink(path.join(base, 'outside.docx'), path.join(root, 'escape.docx'));
    await mkdir(path.join(root, '.secrets')); await writeFile(path.join(root, '.secrets', 'private.docx'), await wordFixture());
    await assert.rejects(readDocument(root, '../outside.docx'), /inside this workspace/);
    await assert.rejects(readDocument(root, 'escape.docx'), /outside this workspace/);
    await assert.rejects(readDocument(root, '.secrets/private.docx'), /inside this workspace/);
    await assert.rejects(readDocument(root, '/etc/passwd'), /Choose a .docx/);
    await assert.rejects(readDocument(root, 'https://example.com/a.docx'));
    await writeFile(path.join(root, 'large.xlsx'), Buffer.alloc(10 * 1024 * 1024 + 1));
    await assert.rejects(readDocument(root, 'large.xlsx'), /10 MB/);
    const list = await discoverFiles(root);
    assert(list.paths.includes('Reading sample.docx'));
    assert(!list.paths.some(value => value.includes('escape') || value.includes('private')));
  } finally { await rm(base, { recursive: true, force: true }); }
});
test('worker opens actual files and leaves source bytes unchanged', async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), 'office-worker-'));
  try {
    await writeFixtures(root);
    const original = await readFile(path.join(root, 'Reading sample.docx'));
    const result = await previewFile(root, 'Reading sample.docx', first);
    assert.equal(result.kind, 'document');
    assert.deepEqual(await readFile(path.join(root, 'Reading sample.docx')), original);
    await assert.rejects(previewFile(root, 'Broken.docx', first), /damaged|encrypted/);
  } finally { await rm(root, { recursive: true, force: true }); }
});
test('rejects expanded ZIP bombs and corrupt Office containers', async () => {
  const zip = new JSZip(); zip.file('huge.xml', Buffer.alloc(21 * 1024 * 1024));
  await assert.rejects(checkZip(await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })), /expanded preview limit/);
  await assert.rejects(parseDocument(Buffer.from('PKbroken'), 'bad.docx', first));
  await assert.rejects(parseDocument(Buffer.from('text'), 'bad.xls', first), /binary Excel/);
});
test('large dimensions are bounded and formulas without saved values stay blank', async () => {
  const workbook = XLSX.utils.book_new();
  const sheet: XLSX.WorkSheet = { A1: { t: 'n', f: '1+2' }, B1: { t: 's', v: '=CMD()' }, '!ref': 'A1:B1' };
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet');
  const zip = await JSZip.loadAsync(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer);
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', xml.replace('<dimension ref="A1:B1"/>', '<dimension ref="A1:XFD1048576"/>'));
  const data = await zip.generateAsync({ type: 'nodebuffer' });
  const preview = await parseDocument(data, 'range.xlsx', first);
  assert(preview.kind === 'workbook');
  assert.equal(preview.totalRows, 10000); assert.equal(preview.totalColumns, 256);
  assert.equal(preview.rows[0]?.[0], ''); assert.equal(preview.rows[0]?.[1], '=CMD()');
  assert(preview.warnings.some(warning => warning.includes('10,000')));
});
test('completed turns append preview cards without replacing chat and resolve nested agent paths', async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), 'office-hook-'));
  try {
    await writeFixtures(root);
    await mkdir(path.join(root, 'nested'));
    let hook: (event: PluginLifecycleEvents['agent.turn_ended'], context: { paseo: PaseoApi; signal: AbortSignal }) => void | Promise<void>;
    const appended: { data: { paths: string[] } }[] = [];
    const fakeServer = { handle() {}, on(_event: string, handler: typeof hook) { hook = handler; return () => {}; } } as unknown as PluginServerContext;
    const paseo = {
      workspaces: { ref: () => ({ refresh: async () => ({ workspaceDirectory: root }) }) },
      agents: { ref: () => ({ timeline: { append: async (item: typeof appended[number]) => { appended.push(item); } } }) },
    } as unknown as PaseoApi;
    const cleanup = contribute(fakeServer);
    const event: PluginLifecycleEvents['agent.turn_ended'] = {
      agent: { id: 'test', workspaceId: 'workspace', parentAgentId: null, provider: 'test', cwd: path.join(root, 'nested'), title: null },
      turnId: 'turn', outcome: { kind: 'completed' },
      timeline: [{ type: 'assistant_message', text: '[Read](<../Reading sample.docx>) [Missing](missing.xlsx)' }],
    };
    const before = JSON.stringify(event.timeline);
    await hook!(event, { paseo, signal: new AbortController().signal });
    assert.equal(JSON.stringify(event.timeline), before);
    assert.deepEqual(appended[0]?.data.paths, [path.join(root, 'Reading sample.docx')]);
    await hook!({ ...event, outcome: { kind: 'canceled', reason: 'test' } }, { paseo, signal: new AbortController().signal });
    assert.equal(appended.length, 1);
    await cleanup();
  } finally { await rm(root, { recursive: true, force: true }); }
});

test('PDF worker renders pages, extracts selectable text and preserves source', async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), 'pdf-worker-'));
  try {
    const bytes = pdfFixture();
    await writeFile(path.join(root, 'Reading.pdf'), bytes);
    const firstPage = await previewFile(root, 'Reading.pdf', { ...first, page: 1 });
    assert(firstPage.kind === 'pdf');
    assert.equal(firstPage.pages, 3);
    assert.match(firstPage.text, /PDF reading sample/);
    assert.match(firstPage.uri, /^data:image\/png;base64,iVBOR/);
    assert(firstPage.width <= 1600 && firstPage.height <= 2200);
    const secondPage = await previewFile(root, 'Reading.pdf', { ...first, page: 2 });
    assert(secondPage.kind === 'pdf' && secondPage.page === 2 && secondPage.text.includes('Second page'));
    const drawing = await previewFile(root, 'Reading.pdf', { ...first, page: 3 });
    assert(drawing.kind === 'pdf' && !drawing.text.trim() && drawing.uri !== firstPage.uri);
    assert.deepEqual(await readFile(path.join(root, 'Reading.pdf')), bytes);
    await assert.rejects(parseDocument(Buffer.from('invalid'), 'broken.pdf', first), /valid PDF/);
    assert.deepEqual(documentLinks('[PDF](Reading.pdf) [remote](https://example.com/a.pdf)'), ['Reading.pdf']);
  } finally { await rm(root, { recursive: true, force: true }); }
});
test('workbook keeps column widths and merged titles, including page boundaries', async () => {
  const preview = await parseDocument(coverageFixture(), 'coverage.xlsx', first);
  assert(preview.kind === 'workbook');
  assert.equal(preview.merges[0]?.columnSpan, 7);
  assert.match(preview.merges[0]!.text, /Group coverage/);
  assert(preview.columnWidths[2]! > preview.columnWidths[0]!);
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Across pages']]);
  sheet['!ref'] = 'A1:V102';
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 101, c: 21 } }];
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet');
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
  const page = await parseDocument(bytes, 'merged.xlsx', { ...first, rowStart: 100, columnStart: 20 });
  assert(page.kind === 'workbook');
  assert.deepEqual(page.merges, [{ row: 0, column: 0, rowSpan: 2, columnSpan: 2, text: 'Across pages' }]);
});
