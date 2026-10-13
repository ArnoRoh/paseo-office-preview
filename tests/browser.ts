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
  await page.goto(url);
  await page.getByRole('button', { name: 'Preview Reading sample.docx', exact: true }).click();
  await page.getByText('Quarterly reading sample', { exact: true }).waitFor();
  await page.screenshot({ path: 'output/desktop-word.png' });
  await page.getByRole('button', { name: 'Close preview' }).click();
  await page.getByRole('button', { name: 'Workbook sample.xlsx', exact: true }).click();
  await page.getByText('Sample item 1', { exact: true }).waitFor();
  await page.screenshot({ path: 'output/desktop-excel.png' });
  await page.getByRole('button', { name: 'Next rows', exact: true }).click();
  await page.getByText('Sample item 100', { exact: true }).waitFor();
  await page.getByRole('button', { name: 'Next columns', exact: true }).click();
  await page.getByRole('button', { name: 'Previous rows', exact: true }).click();
  await page.getByText('Far column', { exact: true }).waitFor();
  await page.getByRole('button', { name: 'Summary', exact: true }).click();
  await page.getByText('Saved values only', { exact: true }).waitFor();
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Broken.docx', exact: true }).click();
  await page.getByRole('alert').waitFor();
  await page.getByRole('button', { name: 'Try again', exact: true }).waitFor();
  await page.keyboard.press('Escape');
  await page.getByRole('textbox', { name: 'Document path' }).fill('../outside.xlsx');
  await page.getByRole('button', { name: 'Preview file', exact: true }).click();
  await page.getByText('Choose a file inside this workspace.', { exact: true }).waitFor();
  await page.keyboard.press('Escape');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(`${url}/?dark=1`);
  await page.getByRole('button', { name: 'Preview Reading sample.docx', exact: true }).click();
  await page.getByText('Quarterly reading sample', { exact: true }).waitFor();
  await page.screenshot({ path: 'output/mobile-word-dark.png' });
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Legacy workbook.xls', exact: true }).click();
  await page.getByText('Sample item 1', { exact: true }).waitFor();
  await page.screenshot({ path: 'output/mobile-excel-dark.png' });
  assert.deepEqual(errors, []);
  await writeFile('output/browser-results.json', JSON.stringify({ passed: true, cases: ['Word popup', 'Excel values', 'row pagination', 'column pagination', 'sheet selection', 'corrupt file', 'path traversal', 'compact dark Word', 'compact dark legacy XLS'], note: 'Real plugin components with a Paseo host UI test adapter; native device and live Paseo acceptance remain separate.' }, null, 2));
  console.log('9 browser checks passed; screenshots saved in output/.');
} finally { await browser?.close(); await new Promise<void>(resolve => server.close(() => resolve())); }
