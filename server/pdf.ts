import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import path from 'node:path';
import { workerData } from 'node:worker_threads';
import type { Preview, Selection } from '../shared/contracts';
import { limits } from '../shared/limits';

export async function pdf(buffer: Buffer, name: string, selection: Selection): Promise<Preview> {
  if (!buffer.subarray(0, 1024).includes(Buffer.from('%PDF-'))) throw new Error('This file is not a valid PDF.');
  // Eval workers need the plugin's module root, not the daemon's working directory.
  const require = createRequire(workerData?.moduleRoot ?? path.join(process.cwd(), 'package.json'));
  const modulePath = require.resolve('pdfjs-dist/legacy/build/pdf.mjs');
  const pdfjs = await import(pathToFileURL(modulePath).href) as typeof import('pdfjs-dist');
  pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs')).href;
  const assets = path.resolve(path.dirname(modulePath), '../..');
  const task = pdfjs.getDocument({
    data: new Uint8Array(buffer), useSystemFonts: false, disableFontFace: true,
    standardFontDataUrl: `${assets}/standard_fonts/`, cMapUrl: `${assets}/cmaps/`, cMapPacked: true,
    wasmUrl: `${assets}/wasm/`, maxImageSize: 8000000, canvasMaxAreaInBytes: 32000000,
  });
  try {
    const document = await task.promise;
    const pageNumber = Math.min(selection.page ?? 1, document.numPages);
    const page = await document.getPage(pageNumber);
    const original = page.getViewport({ scale: 1 });
    const scale = Math.min(2, 1600 / original.width, 2200 / original.height);
    const viewport = page.getViewport({ scale });
    const { createCanvas } = require('@napi-rs/canvas') as typeof import('@napi-rs/canvas');
    const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
    await page.render({ canvas: canvas as never, canvasContext: canvas.getContext('2d') as never, viewport }).promise;
    const image = canvas.toBuffer('image/png');
    if (image.length > 5 * 1024 * 1024) throw new Error('This PDF page exceeds the preview image limit.');
    const content = await page.getTextContent();
    const text = content.items.map(item => 'str' in item ? item.str + (item.hasEOL ? '\n' : ' ') : '').join('');
    return { kind: 'pdf', name, page: pageNumber, pages: document.numPages, width: canvas.width, height: canvas.height,
      uri: `data:image/png;base64,${image.toString('base64')}`, text: text.slice(0, limits.text),
      warnings: text.length > limits.text ? ['Page text is shortened to the preview limit.'] : [],
    };
  } catch (error) {
    if (error instanceof Error && error.name === 'PasswordException') throw new Error('This PDF is password protected. Open an unlocked copy.');
    throw error;
  } finally { await task.destroy(); }
}
