import { Worker } from 'node:worker_threads';
import path from 'node:path';
import moduleRoot from './generated/module-root';
import type { Preview, Selection } from '../shared/contracts';
import { previewSchema } from '../shared/contracts';
import workerSource from './generated/reader-source';
import { readDocument } from './files';

const workers = new Set<Worker>();
export function stopReaders() { for (const worker of workers) void worker.terminate(); }
export async function previewFile(root: string, file: string, selection: Selection): Promise<Preview> {
  if (workers.size >= 2) throw new Error('Two previews are opening. Please try again in a moment.');
  const bytes = await readDocument(root, file);
  if (workers.size >= 2) throw new Error('Two previews are opening. Please try again in a moment.');
  return new Promise((resolve, reject) => {
    const worker = new Worker(workerSource, {
      eval: true, workerData: { moduleRoot: moduleRoot, bytes, name: path.basename(file), selection },
      resourceLimits: { maxOldGenerationSizeMb: 256, maxYoungGenerationSizeMb: 32, stackSizeMb: 4 },
    });
    workers.add(worker);
    const timer = setTimeout(() => { reject(new Error('This document took too long to open. Try a smaller file.')); void worker.terminate(); }, 15000);
    worker.once('message', (message: { preview?: unknown; error?: string }) => {
      clearTimeout(timer);
      if (message.error) reject(new Error(message.error));
      else {
        const result = previewSchema.safeParse(message.preview);
        if (result.success) resolve(result.data); else reject(new Error('The reader returned an invalid preview.'));
      }
      void worker.terminate();
    });
    worker.once('error', () => reject(new Error('The document exceeded the reader limits or could not be opened.')));
    worker.once('exit', () => { clearTimeout(timer); workers.delete(worker); reject(new Error('The document reader stopped. Please try again.')); });
  });
}
