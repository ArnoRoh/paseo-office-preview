import type { PluginServerContext, PluginHandlerContext } from '@getpaseo/plugin/server';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { readPreview, listFiles } from './shared/contracts';
import { documentLinks } from './shared/links';
import { discoverFiles, resolveDocument } from './server/files';
import moduleRoot from './server/generated/module-root';
import { locatePdfRuntime } from './server/runtime';
import { previewFile, stopReaders } from './server/preview';

async function workspaceRoot(paseo: PluginHandlerContext['paseo'], id: string) {
  const workspace = await paseo.workspaces.ref(id).refresh();
  if (!workspace?.workspaceDirectory) throw new Error('This workspace is unavailable.');
  return workspace.workspaceDirectory;
}
export default function contribute(server: PluginServerContext) {
  server.handle(readPreview, async (input, { paseo }) => {
    const root = await workspaceRoot(paseo, input.workspaceId);
    const runtimeRoot = /\.pdf$/i.test(input.path) ? await locatePdfRuntime(moduleRoot, async () => {
      const { config } = await paseo.config.get();
      return Object.values(config.plugins ?? {}).map(plugin => plugin.path);
    }) : moduleRoot;
    return previewFile(root, input.path, input, runtimeRoot);
  });
  server.handle(listFiles, async ({ workspaceId }, { paseo }) => discoverFiles(await workspaceRoot(paseo, workspaceId)));
  const remove = server.on('agent.turn_ended', async ({ agent, timeline, turnId, outcome }, { paseo }) => {
    if (!agent.workspaceId || outcome.kind !== 'completed') return;
    // Read only the final assistant reply. Keep all original timeline entries intact.
    const final = [...timeline].reverse().find(item => item.type === 'assistant_message');
    if (final?.type !== 'assistant_message') return;
    const candidates = documentLinks(final.text);
    if (!candidates.length) return;
    const root = await workspaceRoot(paseo, agent.workspaceId);
    const paths: string[] = [];
    for (const candidate of candidates) {
      try {
        const resolved = await resolveDocument(root, path.resolve(agent.cwd, candidate));
        if (!paths.includes(resolved) && Buffer.byteLength(JSON.stringify([...paths, resolved])) < 50000) paths.push(resolved);
      } catch { /* Missing or outside-workspace references remain ordinary links. */ }
    }
    if (!paths.length) return;
    const id = createHash('sha256').update(turnId ?? final.messageId ?? final.text).digest('hex').slice(0, 20);
    await paseo.agents.ref(agent.id).timeline.append({
      type: 'plugin', id: `preview-${id}`, kind: 'office-files', version: 1,
      data: { workspaceId: agent.workspaceId, paths },
    });
  });
  return () => { remove(); stopReaders(); };
}
