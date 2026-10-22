import { build } from 'esbuild';
import { createServer } from 'node:http';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';
import { previewFile } from '../server/preview';
import { discoverFiles } from '../server/files';
import { writeFixtures } from './fixtures';

await mkdir('output', { recursive: true });
const root = path.resolve('output/fixtures');
await writeFixtures(root);
const bundle = await build({
  entryPoints: ['tests/ui-app.tsx'], bundle: true, write: false, platform: 'browser', format: 'iife', jsx: 'automatic',
  define: { 'process.env.NODE_ENV': '"development"', __DEV__: 'true' },
  alias: { 'react-native': 'react-native-web', '@getpaseo/plugin/client': './tests/host-mock.tsx', '@getpaseo/plugin/client/react-native': './tests/host-mock.tsx' },
});
const server = createServer(async (request, response) => {
  try {
    if (request.url?.startsWith('/rpc/')) {
      let input = ''; for await (const chunk of request) input += chunk;
      const body = JSON.parse(input);
      const result = request.url.endsWith('.list') ? await discoverFiles(root) : await previewFile(root, body.path, body);
      response.setHeader('Content-Type', 'application/json'); response.end(JSON.stringify(result));
    } else if (request.url === '/app.js') { response.setHeader('Content-Type', 'text/javascript'); response.end(bundle.outputFiles[0]!.text); }
    else { response.setHeader('Content-Type', 'text/html'); response.end('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Office preview test</title><style>html,body,#root{height:100%;margin:0}#root{display:flex;flex-direction:column}</style></head><body><div id="root"></div><script src="/app.js"></script></body></html>'); }
  } catch (error) { response.statusCode = 400; response.end(JSON.stringify({ error: error instanceof Error ? error.message : 'Preview failed' })); }
});
await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
const address = server.address(); assert(address && typeof address !== 'string');
let browser;
const errors: string[] = [];
try {
  browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE });
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  page.on('pageerror', error => errors.push(error.message));
  const url = `http://127.0.0.1:${address.port}`;
  const close = () => page.getByRole('button', { name: 'Close preview', exact: true }).click();
  await page.goto(url);
  await page.getByRole('button', { name: 'Preview Reading sample.docx', exact: true }).click();
  await page.getByText('Quarterly reading sample', { exact: true }).waitFor();
  assert.equal(await page.getByRole('dialog').count(), 0, 'Desktop cards open the full-width reader');
  await page.getByRole('textbox', { name: 'Find in document' }).fill('figures');
  await page.getByText('1 matching sections', { exact: true }).waitFor();
  assert.equal(await page.getByText('Quarterly reading sample', { exact: true }).count(), 0);
  await page.getByRole('button', { name: 'Clear search', exact: true }).click();
  await page.getByRole('button', { name: 'Larger text', exact: true }).click();
  await page.screenshot({ path: 'output/desktop-word.png' });
  await close();
  await page.getByRole('button', { name: 'Coverage sample.xlsx', exact: true }).click();
  const title = page.getByRole('button', { name: 'A1: Group coverage and timing | Synthetic preview sample', exact: true });
  await title.waitFor();
  const titleBox = await title.boundingBox(); assert(titleBox && titleBox.width > 1100, 'Merged title spans the desktop grid');
  assert.equal(await page.getByRole('button', { name: 'Next rows', exact: true }).count(), 0, 'Short sheets have no paging clutter');
  await page.screenshot({ path: 'output/desktop-excel.png' });
  await title.click();
  await page.getByText('A1', { exact: true }).waitFor();
  await close();
  await page.getByRole('button', { name: 'Workbook sample.xlsx', exact: true }).click();
  await page.getByText('Sample item 1', { exact: true }).waitFor();
  await page.getByRole('button', { name: 'Next rows', exact: true }).click();
  await page.getByText('Sample item 100', { exact: true }).waitFor();
  await page.getByRole('button', { name: 'Next columns', exact: true }).click();
  await page.getByText('Columns U', { exact: false }).waitFor();
  await page.getByRole('button', { name: 'Previous rows', exact: true }).click();
  await page.getByText('Far column', { exact: true }).waitFor();
  await page.getByRole('textbox', { name: 'Go to cell', exact: true }).fill('A101');
  await page.getByRole('button', { name: 'Go to cell', exact: true }).click();
  await page.getByText('Sample item 100', { exact: true }).waitFor();
  await page.getByRole('button', { name: 'Summary', exact: true }).click();
  await page.getByText('Saved values only', { exact: true }).waitFor();
  await close();
  await page.getByRole('button', { name: 'Reading sample.pdf', exact: true }).click();
  const pdfImage = page.getByRole('img', { name: 'PDF page 1 of 3' });
  await pdfImage.waitFor();
  await page.screenshot({ path: 'output/desktop-pdf.png' });
  const normalWidth = (await pdfImage.boundingBox())!.width;
  await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
  assert((await pdfImage.boundingBox())!.width > normalWidth);
  await page.getByRole('button', { name: 'Fit width', exact: true }).click();
  await page.getByRole('button', { name: 'Next page', exact: true }).click();
  await page.getByRole('img', { name: 'PDF page 2 of 3' }).waitFor();
  await page.getByRole('button', { name: 'Selectable text', exact: true }).click();
  await page.getByText('Second page', { exact: false }).waitFor();
  await page.getByRole('textbox', { name: 'Page number' }).fill('3');
  await page.getByRole('button', { name: 'Go to page', exact: true }).click();
  await page.getByText('No selectable text on this page.', { exact: false }).waitFor();
  await page.getByRole('button', { name: 'Page view', exact: true }).click();
  await page.getByRole('img', { name: 'PDF page 3 of 3' }).waitFor();
  await close();
  for (const name of ['Broken.docx', 'Broken.pdf']) {
    await page.getByRole('button', { name, exact: true }).click();
    await page.getByRole('alert').waitFor();
    await page.getByRole('button', { name: 'Try again', exact: true }).waitFor();
    await close();
  }
  await page.getByRole('textbox', { name: 'Document path' }).fill('../outside.xlsx');
  await page.getByRole('button', { name: 'Preview file', exact: true }).click();
  await page.getByText('Choose a file inside this workspace.', { exact: true }).waitFor();
  await close();
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(`${url}/?dark=1`);
  await page.getByRole('button', { name: 'Preview Reading sample.docx', exact: true }).click();
  await page.getByText('Quarterly reading sample', { exact: true }).waitFor();
  assert.equal(await page.getByRole('dialog').count(), 1, 'Compact cards keep native modal behavior');
  await page.screenshot({ path: 'output/mobile-word-dark.png' });
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Legacy workbook.xls', exact: true }).click();
  await page.getByText('Sample item 1', { exact: true }).waitFor();
  await page.screenshot({ path: 'output/mobile-excel-dark.png' });
  await close();
  await page.getByRole('button', { name: 'Reading sample.pdf', exact: true }).click();
  await page.getByRole('img', { name: 'PDF page 1 of 3' }).waitFor();
  await page.screenshot({ path: 'output/mobile-pdf-dark.png' });
  assert.deepEqual(errors, []);
  await writeFile('output/browser-results.json', JSON.stringify({ passed: true, cases: ['desktop full-width reader', 'Word search and text size', 'merged spreadsheet title and widths', 'cell inspector', 'row and column navigation', 'cell jump', 'sheet selection', 'PDF rendering', 'PDF zoom and fit width', 'PDF page navigation and text', 'PDF drawing page', 'corrupt Office and PDF', 'path traversal', 'compact dark Word modal', 'compact dark legacy XLS', 'compact dark PDF'], note: 'Real plugin components with a Paseo host UI test adapter; native device and live Paseo acceptance remain separate.' }, null, 2));
  console.log('16 browser scenarios passed; screenshots saved in output/.');
} finally { await browser?.close(); await new Promise<void>(resolve => server.close(() => resolve())); }
