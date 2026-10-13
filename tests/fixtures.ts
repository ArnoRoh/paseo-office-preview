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
  await writeFile(path.join(dir, 'Broken.docx'), 'not a document');
}
