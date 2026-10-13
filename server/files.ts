import { constants } from 'node:fs';
import { open, realpath, readdir } from 'node:fs/promises';
import path from 'node:path';
import { limits } from '../shared/limits';
import { isOfficePath } from '../shared/links';

const excluded = new Set(['.git', '.secrets', '.ssh', '.aws', 'node_modules', '.next', '.venv', 'venv']);
function blocked(parts: string[]) {
  return parts.some(part => excluded.has(part.toLowerCase()) || /^\.env(?:\.|$)/i.test(part));
}
function contained(root: string, target: string) {
  const relative = path.relative(root, target);
  return relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative) && !blocked(relative.split(path.sep));
}
export async function resolveDocument(root: string, supplied: string): Promise<string> {
  if (!supplied || supplied.includes('\0') || !isOfficePath(supplied)) throw new Error('Choose a .docx, .xlsx or .xls file.');
  const base = await realpath(root);
  const requested = path.resolve(base, supplied);
  if (!contained(base, requested)) throw new Error('Choose a file inside this workspace.');
  const resolved = await realpath(requested);
  if (!contained(base, resolved) || !isOfficePath(resolved)) throw new Error('The file points outside this workspace or to an unsupported file.');
  return resolved;
}
export async function readDocument(root: string, supplied: string): Promise<Buffer> {
  const resolved = await resolveDocument(root, supplied);
  const handle = await open(resolved, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0));
  try {
    const stat = await handle.stat();
    if (!stat.isFile()) throw new Error('Choose a regular file.');
    if (stat.size > limits.fileBytes) throw new Error('This file exceeds the 10 MB preview limit.');
    // Verify the opened file on Linux, including races in parent symlinks.
    if (process.platform === 'linux') {
      const opened = await realpath(`/proc/self/fd/${handle.fd}`);
      if (!contained(await realpath(root), opened)) throw new Error('The file moved outside this workspace.');
    }
    const bytes = Buffer.alloc(Math.min(stat.size + 1, limits.fileBytes + 1));
    let length = 0;
    while (length < bytes.length) {
      const result = await handle.read(bytes, length, bytes.length - length, null);
      if (!result.bytesRead) break;
      length += result.bytesRead;
    }
    if (length > stat.size) throw new Error('The file changed while opening. Try again.');
    return bytes.subarray(0, length);
  } finally { await handle.close(); }
}
export async function discoverFiles(root: string) {
  const paths: string[] = [];
  const queue = [''];
  let visited = 0;
  while (queue.length && paths.length < 200 && visited < 5000) {
    const dir = queue.shift()!;
    const entries = await readdir(path.join(root, dir), { withFileTypes: true }).catch(() => []);
    for (const entry of entries) {
      if (++visited > 5000 || paths.length === 200) break;
      if (blocked([entry.name]) || entry.isSymbolicLink()) continue;
      const relative = path.join(dir, entry.name);
      if (entry.isDirectory() && relative.split(path.sep).length < 6) queue.push(relative);
      else if (entry.isFile() && isOfficePath(entry.name)) paths.push(relative);
    }
  }
  return { paths: paths.sort(), truncated: paths.length === 200 || visited >= 5000 || queue.length > 0 };
}
