import type { PluginClientContext } from '@getpaseo/plugin/client';
import { OfficePanel, OfficeCard } from './client/panel';
import { selectPreview } from './client/selection';
import { cardSchema } from './shared/contracts';

export default function contribute(client: PluginClientContext) {
  client.addWorkspacePanel({ id: 'office-preview', title: 'Office preview', icon: 'FileText', context: 'workspace', Component: OfficePanel });
  client.addCommandCenterItem({ id: 'office-preview', title: 'Open Office preview', icon: 'FileText', context: 'workspace', onSelect: ({ openPanel }) => openPanel('office-preview') });
  client.addSlashCommand({ name: 'office-preview', description: 'Read Word, Excel and PDF files in this workspace', argumentHint: '', context: 'workspace', onSubmit: ({ openPanel }) => openPanel('office-preview') });
  client.addTimelineRenderer({ kind: 'office-files', version: 1, schema: cardSchema, Component: props => <OfficeCard {...props} openReader={(workspaceId, path) => {
    selectPreview(props.host.id, workspaceId, path);
    client.openPanel('office-preview', { workspaceId });
  }} /> });
  return () => {};
}
