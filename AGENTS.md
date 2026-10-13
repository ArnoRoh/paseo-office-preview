# Office Preview

Build a shareable, read-only Paseo plugin for Word and Excel files on the daemon host.
Use native Paseo modals and React Native components on desktop and mobile.
Preserve the original chat messages. Add preview cards after agent turns.

Read only files within the selected workspace. Resolve symlinks before access.
Do not fetch remote document resources, execute macros, or evaluate formulas.
Keep file parsing in a bounded worker. Validate RPC input and output.
Use synthetic fixtures for tests. Never package private documents or host paths.

Checks: npm run build, npm run typecheck, npm test, npm run test:ui.
Package: npm pack. Keep README.md and OVERVIEW.md as standard plugin entrypoints.
Record internal development notes in dated files under docs/.
Public publication and enabling the live daemon's plugins require explicit approval.
