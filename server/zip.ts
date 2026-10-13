import yauzl from 'yauzl';
import { limits } from '../shared/limits';

export async function checkZip(buffer: Buffer): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    yauzl.fromBuffer(buffer, { lazyEntries: true, validateEntrySizes: true }, (error, zip) => {
      if (error || !zip) { reject(new Error('This document is not a valid Office file.')); return; }
      let entries = 0;
      let expanded = 0;
      let failed = false;
      const fail = (cause: unknown) => { if (!failed) { failed = true; zip.close(); reject(cause); } };
      zip.on('error', fail);
      zip.on('end', () => resolve());
      zip.on('entry', entry => {
        expanded += entry.uncompressedSize;
        if (++entries > limits.zipEntries || expanded > limits.expandedBytes || entry.uncompressedSize > limits.entryBytes) {
          fail(new Error('This document exceeds the expanded preview limit.')); return;
        }
        if (entry.generalPurposeBitFlag & 1) { fail(new Error('Password-protected files cannot be previewed.')); return; }
        // Drain each entry to validate actual size, rather than trusting ZIP headers.
        zip.openReadStream(entry, (streamError, stream) => {
          if (streamError || !stream) { fail(new Error('This Office file is damaged.')); return; }
          let size = 0;
          stream.on('data', chunk => {
            size += chunk.length;
            if (size > limits.entryBytes) { stream.destroy(); fail(new Error('This document exceeds the expanded preview limit.')); }
          });
          stream.on('error', fail);
          stream.on('end', () => { if (!failed) zip.readEntry(); });
        });
      });
      zip.readEntry();
    });
  });
}
