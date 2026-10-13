import { createRoot } from 'react-dom/client';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { OfficePanel, OfficeCard } from '../client/panel';

const dark = new URLSearchParams(location.search).has('dark');
const theme: any = { colors: {
  foreground: dark ? '#f3f4f6' : '#17212c', foregroundMuted: dark ? '#a9b3c1' : '#536174',
  surface0: dark ? '#131920' : '#ffffff', surface1: dark ? '#1b2530' : '#f4f6f8',
  surface2: dark ? '#263545' : '#e9edf2', border: dark ? '#455468' : '#cad2dc', accent: '#245bc4', accentForeground: '#ffffff',
} };
const props: any = { theme, layout: { compact: window.innerWidth < 600, platform: 'web' }, host: { id: 'test-host', label: 'Test host' }, workspaceId: 'test-workspace', context: 'workspace' };
createRoot(document.getElementById('root')!).render(<QueryClientProvider client={new QueryClient()}>
  <OfficePanel {...props} />
  <OfficeCard {...props} agentId="test-agent" item={{ type: 'plugin', kind: 'office-files', version: 1, data: { workspaceId: 'test-workspace', paths: ['Reading sample.docx'] } }} timestamp={new Date()} />
</QueryClientProvider>);
