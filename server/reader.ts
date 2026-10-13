import { parentPort, workerData } from 'node:worker_threads';
import mammoth from 'mammoth';
import { parseFragment } from 'parse5';
import * as XLSX from 'xlsx';
import type { Block, Preview, Run, Selection } from '../shared/contracts';
import { limits } from '../shared/limits';
import { checkZip } from './zip';

interface HtmlNode { nodeName: string; value?: string; attrs?: { name: string; value: string }[]; childNodes?: HtmlNode[] }
const attr = (node: HtmlNode, name: string) => node.attrs?.find(item => item.name === name)?.value ?? '';
const children = (node: HtmlNode) => node.childNodes ?? [];

async function word(buffer: Buffer, name: string): Promise<Preview> {
  let imageBytes = 0;
  let omittedImages = false;
  const converted = await mammoth.convertToHtml({ buffer }, {
    externalFileAccess: false,
    convertImage: mammoth.images.imgElement(async image => {
      if (!['image/png', 'image/jpeg', 'image/gif'].includes(image.contentType)) { omittedImages = true; return { src: '' }; }
      const data = await image.read('base64');
      imageBytes += data.length;
      if (data.length > 1400000 || imageBytes > 4000000) { omittedImages = true; return { src: '' }; }
      return { src: `data:${image.contentType};base64,${data}` };
    }),
  });
  const blocks: Block[] = [];
  let textLength = 0;
  let truncated = false;
  function runs(node: HtmlNode, style: Omit<Run, 'text'> = {}, depth = 0): Run[] {
    if (depth > 50 || textLength >= limits.text) { truncated = true; return []; }
    if (node.nodeName === '#text') {
      const text = (node.value ?? '').slice(0, limits.text - textLength);
      textLength += text.length;
      return [{ text, ...style }];
    }
    if (node.nodeName === 'br') return [{ text: '\n' }];
    return children(node).flatMap(child => runs(child, {
      ...style,
      ...(['b', 'strong'].includes(node.nodeName) ? { bold: true } : {}),
      ...(['i', 'em'].includes(node.nodeName) ? { italic: true } : {}),
    }, depth + 1));
  }
  function images(node: HtmlNode, depth = 0) {
    if (depth > 50 || blocks.length >= limits.blocks) return;
    if (node.nodeName === 'img') {
      const uri = attr(node, 'src');
      if (/^data:image\/(png|jpeg|gif);base64,/.test(uri)) blocks.push({ kind: 'image', uri, alt: attr(node, 'alt') || 'Document image' });
    } else children(node).forEach(child => images(child, depth + 1));
  }
  function visit(node: HtmlNode, depth = 0) {
    if (depth > 50 || blocks.length >= limits.blocks || textLength >= limits.text) { truncated = true; return; }
    if (/^(p|h[1-6]|li)$/.test(node.nodeName)) {
      const content = runs(node);
      if (node.nodeName === 'li') content.unshift({ text: '• ' });
      if (content.some(run => run.text.trim())) blocks.push({ kind: 'paragraph', runs: content, ...(/^h[1-6]$/.test(node.nodeName) ? { heading: Number(node.nodeName[1]) } : {}) });
      images(node);
    } else if (node.nodeName === 'table') {
      const tableRows = children(node).flatMap(section => section.nodeName === 'tr' ? [section] : children(section).filter(row => row.nodeName === 'tr'));
      const rows = tableRows.slice(0, 200).map(row => children(row).filter(cell => ['td', 'th'].includes(cell.nodeName)).slice(0, 20).map(cell => runs(cell)));
      blocks.push({ kind: 'table', rows });
      images(node);
      if (tableRows.length > 200 || tableRows.some(row => children(row).filter(cell => ['td', 'th'].includes(cell.nodeName)).length > 20)) truncated = true;
    } else if (node.nodeName === 'img') images(node);
    else children(node).forEach(child => visit(child, depth + 1));
  }
  visit(parseFragment(converted.value) as HtmlNode);
  return {
    kind: 'document', name, blocks,
    warnings: [
      'Reading view. Page layout, fonts, headers, footers and tracked changes may differ from Word.',
      ...(converted.messages.length ? ['Some document elements could not be displayed.'] : []),
      ...(omittedImages ? ['Some images are too large or use an unsupported format.'] : []),
      ...(truncated ? ['This preview is shortened because the document exceeds the reading limits.'] : []),
    ],
  };
}

function spreadsheet(buffer: Buffer, name: string, selection: Selection): Preview {
  const workbook = XLSX.read(buffer, { type: 'buffer', sheetRows: limits.rows + 1, cellHTML: false, cellStyles: false, bookVBA: false, cellFormula: true });
  const sheets = workbook.SheetNames.filter((_, index) => !workbook.Workbook?.Sheets?.[index]?.Hidden).slice(0, 100);
  if (!sheets.length) throw new Error('This workbook has no visible sheets.');
  if (selection.sheet >= sheets.length) throw new Error('This worksheet is no longer available. Reopen the workbook.');
  const sheet = workbook.Sheets[sheets[selection.sheet]!]!;
  const range = XLSX.utils.decode_range(sheet['!fullref'] ?? sheet['!ref'] ?? 'A1');
  const totalRows = sheet['!ref'] ? Math.min(range.e.r + 1, limits.rows) : 0;
  const totalColumns = sheet['!ref'] ? Math.min(range.e.c + 1, limits.columns) : 0;
  const rowStart = Math.min(selection.rowStart, Math.max(0, totalRows - 1));
  const columnStart = Math.min(selection.columnStart, Math.max(0, totalColumns - 1));
  let missingFormulaValue = false;
  let truncatedCells = false;
  const rows: string[][] = [];
  for (let r = rowStart; r < Math.min(totalRows, rowStart + limits.pageRows); r++) {
    const row: string[] = [];
    for (let c = columnStart; c < Math.min(totalColumns, columnStart + limits.pageColumns); c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
      if (cell?.f && cell.v === undefined) missingFormulaValue = true;
      const text = cell?.v === undefined ? '' : XLSX.utils.format_cell(cell);
      if (text.length > limits.cellText) truncatedCells = true;
      row.push(text.length > limits.cellText ? `${text.slice(0, limits.cellText)}…` : text);
    }
    rows.push(row);
  }
  return {
    kind: 'workbook', name, sheets, sheet: selection.sheet, rows, rowStart, columnStart, totalRows, totalColumns,
    warnings: [
      'Saved cell values. Formulas are not recalculated; charts, styling and macros are not shown.',
      ...(range.e.r >= limits.rows || range.e.c >= limits.columns ? ['Preview is limited to the first 10,000 rows and 256 columns.'] : []),
      ...(workbook.SheetNames.length > 100 ? ['Only the first 100 visible sheets are available.'] : []),
      ...(sheet['!merges']?.length ? ['Merged cells show their value in the first cell.'] : []),
      ...(missingFormulaValue ? ['Some formulas have no saved result and appear blank.'] : []),
      ...(truncatedCells ? ['Some long cell values are shortened to 2,000 characters.'] : []),
    ],
  };
}

export async function parseDocument(buffer: Buffer, name: string, selection: Selection): Promise<Preview> {
  const zipped = buffer[0] === 0x50 && buffer[1] === 0x4b;
  if (/\.(docx|xlsx)$/i.test(name) && !zipped) throw new Error('This file is damaged, encrypted or is not the expected Office format.');
  if (/\.xls$/i.test(name) && !(buffer.subarray(0, 8).equals(Buffer.from('d0cf11e0a1b11ae1', 'hex')))) throw new Error('This file is not a supported binary Excel workbook.');
  if (zipped) await checkZip(buffer);
  return /\.docx$/i.test(name) ? word(buffer, name) : spreadsheet(buffer, name, selection);
}
if (parentPort && workerData) {
  parseDocument(Buffer.from(workerData.bytes), workerData.name, workerData.selection)
    .then(preview => parentPort!.postMessage({ preview }))
    .catch(error => parentPort!.postMessage({ error: error instanceof Error ? error.message : 'The document could not be read.' }));
}
