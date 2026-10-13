# Office Preview development record

## Accepted scope

Build a shareable Paseo plugin for read-only Word and Excel previews. Open the
reader in Paseo. Use desktop dialogs and compact sheets. The user does not want
editing or an external document portal.

## Implementation

- Native React Native reading components, using Paseo's Modal and scrolling controls.
- Workspace panel, Command Center action and `/office-preview` command.
- Preview cards appended after completed replies with supported local file references.
- Preserve source messages and files. No remote image fetches or formula calculation.
- Semantic Word rendering and saved-value Excel rendering on a bounded worker.
- Workspace path checks, symlink checks, file-size limits and ZIP expansion limits.
- MIT license, registry overview, npm package, pinned dependency shrinkwrap and CI.

## Evidence

The installed Paseo 0.10.2 compiler accepts both client and server entrypoints.
It also accepts the extracted release with production dependencies only. The
archive passes `npm ci --omit=dev` and the worker build. TypeScript checks pass.
Nine synthetic parser/integration tests cover Word content, both Excel
formats, worksheet paging, access checks, source-file preservation, ZIP limits,
large spreadsheet dimensions and automatic preview cards.

The browser harness uses real plugin components and a test adapter for host-owned
modal and RPC APIs. It covers desktop and compact layouts, dark mode, sheet and
page controls, corrupt files and rejected paths. Nine browser checks pass. This does not prove native
device gestures or live Paseo integration.

The live daemon is reachable and reports version 0.10.2. Its plugin catalog is
empty and its `pluginsEnabled` setting is false. Installation and enablement are
pending user approval. No live service was restarted. Public publication is also
pending approval of the destination and visibility.

Task Mgmt has no available authenticated tool in this session. The connected tool
catalog is empty. No new credential was created and no task row was written.

## Dependency review

Yauzl was updated to 3.4.0 to resolve the ZIP-reader advisory found during the
audit. Mammoth is pinned to 1.13.0. npm still reports a moderate advisory through
Mammoth's CLI dependency `argparse` and `sprintf-js` 1.1.3. There is no newer
sprintf-js release in the registry. The reader does not import the Mammoth CLI;
the worker bundle does not include argparse or sprintf-js. Bundled package
licenses are collected under `server/generated/`.

## Acceptance still needed

Install on an enabled daemon, open a document from an actual agent reply, and
check desktop and phone behavior. Confirm the compact sheet scrolls correctly.
Then publish the reviewed source or archive to the selected sharing destination.
