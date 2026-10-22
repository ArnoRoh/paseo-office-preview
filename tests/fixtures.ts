import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

export async function wordFixture(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Quarterly reading sample</w:t></w:r></w:p>
    <w:p><w:r><w:t>This synthetic document checks the Office Preview reader. </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Read only.</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>All figures are test data.</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Count</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Sample documents</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>24</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>Long text remains readable on a phone. You can select and copy this text without changing the source document.</w:t></w:r></w:p>
    </w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
export function workbookFixture(type: 'xlsx' | 'biff8' = 'xlsx'): Buffer {
  const workbook = XLSX.utils.book_new();
  const rows: (string | number)[][] = [['Item', 'Quantity', 'Unit price', 'Total']];
  for (let i = 1; i <= 125; i++) rows.push([`Sample item ${i}`, i, 2.5, i * 2.5]);
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet.D2 = { t: 'n', v: 2.5, f: 'B2*C2', z: '#,##0.00' };
  sheet.E2 = { t: 'n', f: '1+2' };
  sheet.U1 = { t: 's', v: 'Far column' };
  sheet['!ref'] = 'A1:U126';
  XLSX.utils.book_append_sheet(workbook, sheet, 'Inventory');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Summary'], ['Saved values only']]), 'Summary');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Hidden test value']]), 'Hidden');
  workbook.Workbook = { Sheets: [{ name: 'Inventory', Hidden: 0 }, { name: 'Summary', Hidden: 0 }, { name: 'Hidden', Hidden: 1 }] };
  return XLSX.write(workbook, { type: 'buffer', bookType: type }) as Buffer;
}
export async function writeFixtures(dir: string) {
  await mkdir(dir, { recursive: true });
  await writeFile(path.join(dir, 'Reading sample.docx'), await wordFixture());
  await writeFile(path.join(dir, 'Workbook sample.xlsx'), workbookFixture());
  await writeFile(path.join(dir, 'Legacy workbook.xls'), workbookFixture('biff8'));
  await writeFile(path.join(dir, 'Reading sample.pdf'), pdfFixture());
  await writeFile(path.join(dir, 'Coverage sample.xlsx'), coverageFixture());
  await writeFile(path.join(dir, 'Broken.pdf'), 'not a PDF');
  await writeFile(path.join(dir, 'Broken.docx'), 'not a document');
}

export function pdfFixture(): Buffer {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>',
    ...[6, 7, 8].map(content => `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 9 0 R >> >> /Contents ${content} 0 R >>`),
    ...['BT /F1 24 Tf 72 720 Td (PDF reading sample) Tj ET', 'BT /F1 24 Tf 72 720 Td (Second page) Tj ET', '0.2 0.4 0.7 rg 72 600 200 100 re f'].map(stream => `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`),
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  let text = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((object, index) => { offsets.push(Buffer.byteLength(text)); text += `${index + 1} 0 obj\n${object}\nendobj\n`; });
  const xref = Buffer.byteLength(text);
  text += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` + offsets.slice(1).map(offset => `${String(offset).padStart(10, '0')} 00000 n \n`).join('');
  text += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(text);
}
export function coverageFixture(): Buffer {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Group coverage and timing | Synthetic preview sample'],
    ['Company', 'Contributor', 'September timing', 'October timing', 'Status', 'Follow-up', 'Source'],
    ['Sample company A', 'Contributor A', 'September 21 update; September 22 review; September 29 correction', 'October 12 review', 'Complete', 'Confirm final figures', 'Sample source'],
    ['Sample company B', 'Contributor B', 'September 2 initial changes; September 14 advances; September 16 loans', 'October 15 follow-up', 'In review', 'Check remaining inputs', 'Sample source'],
  ]);
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
  sheet['!cols'] = [26, 22, 40, 30, 18, 30, 20].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(workbook, sheet, 'Company coverage');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Source events'], ['Test event']]), 'Source events');
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}
