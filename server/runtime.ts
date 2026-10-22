import { access, readFile } from 'node:fs/promises';
import path from 'node:path';

export async function locatePdfRuntime(builtRoot: string, directories: () => Promise<string[]>): Promise<string> {
  try { await access(builtRoot); return builtRoot; } catch { /* Managed installs move after preparation. */ }
  const marker = `export default ${JSON.stringify(builtRoot)};`;
  for (const directory of await directories()) {
    const source = await readFile(path.join(directory, 'server/generated/module-root.ts'), 'utf8').catch(() => '');
    if (!source.includes(marker)) continue;
    const installedRoot = path.join(directory, 'node_modules/pdfjs-dist/package.json');
    await access(installedRoot);
    return installedRoot;
  }
  throw new Error('The PDF runtime could not be located. Rebuild and reload this plugin.');
}
