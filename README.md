# Office Preview for Paseo

Read `.docx`, `.xlsx` and `.xls` files from the daemon host in a Paseo popup.
Uses native Paseo UI on desktop, web, iOS and Android. The source files remain unchanged.

## Install

Requires Paseo **0.10.2 or 0.11.x** on the host and client, Node.js 22+, and npm.
Enable plugins in **Settings → Plugins**. Paseo plugins run as trusted code.

From a source checkout or an extracted release archive, run:

```sh
npm ci --omit=dev
npm run build
paseo plugin install /absolute/path/to/paseo-office-preview
```

Install directly from GitHub:

```sh
paseo plugin install github:ArnoRoh/paseo-office-preview
```

Run this against the Paseo daemon that owns your files. For a remote daemon, add
`--host <daemon-address>`. Local source paths must exist on that daemon host.
Git installation runs the dependency install and worker build automatically.

## Read a document

1. Ask an agent to return a Markdown link to a file in its workspace.
2. When the turn completes, select the added **Preview filename** button.
3. Read the file in the popup. Select a worksheet or page through a workbook.

The plugin preserves the original reply. It adds its own card below it.
Existing replies are not rewritten. To open an older file, use the **Office
preview** workspace panel, the **Open Office preview** command, or `/office-preview`.
Choose a listed file or paste an absolute or workspace-relative path.

Links with spaces work, for example `[Read report](<reports/Reading sample.docx>)`.
Backtick file references are also detected. External web links are not fetched.

## Supported content

| File | Reading view |
| --- | --- |
| `.docx` | Paragraphs, headings, bold and italic text, simple tables, PNG/JPEG/GIF images |
| `.xlsx` | Visible worksheets, formatted saved values, row and column paging |
| `.xls` | Binary Excel workbooks, using the same reading view |

Word uses a semantic reading view. It does not reproduce page geometry, fonts,
headers, footers, comments or tracked changes. Lists are simplified. Nested
tables, merged cells and complex layouts can lose their structure. Large or
unsupported images are omitted. Old `.doc` files require conversion to `.docx`
outside this plugin. Password-protected documents are not supported.

Excel formulas use saved results, which can be stale or absent. This plugin does
not calculate formulas or run macros. It does not show charts, workbook styling
or hidden sheets. Merged cells show their value in the first cell. Hidden rows
and columns within a visible sheet remain readable.

## Access and limits

- Reads files within the selected workspace only, after symlink resolution.
- Skips secret directories, Git metadata, dependency folders and environment files.
- Does not upload documents to a third-party service or fetch document resources.
- Transfers preview data through the existing Paseo connection.
- Does not save converted copies, modify files or add an HTTP listener.
- Opens at most two readers at once, with a 15-second timeout and bounded worker memory.
- Accepts files up to 10 MB and Office ZIP containers up to 50 MB expanded.
- Shows up to 2,000 Word blocks / 500,000 text characters and bounded embedded images.
- Shows up to 100 visible worksheets, 10,000 rows and 256 columns per sheet.
- Uses pages of 100 rows and 20 columns; long cell values are shortened with a notice.
- Lists up to 200 files, scanning up to 5,000 entries and five directory levels.

Paseo plugins are trusted, unsandboxed code. These checks constrain this reader;
they are not a security boundary against other installed plugins or the daemon user.

## Develop and verify

```sh
npm ci
npm run build
npm run typecheck
npm test
npx playwright install chromium
npm run test:ui
npm pack
```

`npm run build` bundles the document parser into a worker and collects its license
notices. `npm pack` includes the generated worker. The shrinkwrap file pins the
dependencies for both Git and npm distribution.

The parser tests use synthetic documents. The browser harness mounts the real
plugin components with a test adapter for Paseo's host-owned modal and RPC hooks.
It checks Word and Excel previews, sheet and page controls, errors, compact
layout and dark theme. It stores screenshots in ignored `output/`.

Set `PLAYWRIGHT_CHROMIUM_EXECUTABLE` to use an existing Chromium executable.
Native device gestures and actual Paseo client acceptance require a live check.

## Dependencies

Mammoth extracts Word content. Parse5 converts its generated HTML into a limited
native reading model. SheetJS CE reads Excel data. Yauzl validates Office ZIP
entries before parsing. Esbuild creates the bounded worker. These libraries are
used on the daemon; client code uses only Paseo's supported host modules.

The plugin is MIT licensed. Bundled dependency licenses are included in
`server/generated/THIRD_PARTY_NOTICES.txt`.

## Share

The source is suitable for a standalone Git repository. The npm archive can be
extracted and installed as a local plugin directory. `OVERVIEW.md` supplies the
registry listing text. Public GitHub/npm publication and registry submission are
separate from building the package.
