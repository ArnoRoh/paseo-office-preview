# Office Preview for Paseo

Read `.docx`, `.xlsx`, `.xls` and `.pdf` files from the daemon host in a full-width Paseo workspace reader or compact preview sheet.
Uses native Paseo UI on desktop, web, iOS and Android. The source files remain unchanged.

## Install

Requires Paseo **0.10.2 or 0.11.x** on the host and client, Node.js 22.13+, and npm.
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
3. On desktop, read the file in the full-width workspace panel. Compact preview
   cards use a native sheet. Select **Close preview** to return to the file list.
4. For spreadsheets, use worksheet tabs, zoom, **Fit width**, or **Go to cell**.
   Select a cell to read and copy its value. Row and column controls appear only
   when the sheet needs them.
5. For Word, use **Contents**, **Find in document**, and the text-size controls.
6. For PDF, use page navigation, **Fit width**, zoom, and **Selectable text**.

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
| `.xlsx` | Visible worksheets, saved values, source column widths, merged cells and range navigation |
| `.xls` | Binary Excel workbooks, using the same reading view |
| `.pdf` | Rendered pages, page navigation, zoom and selectable page text |

Word uses a semantic reading view. It does not reproduce page geometry, fonts,
headers, footers, comments or tracked changes. Lists are simplified. Nested
tables, merged cells and complex layouts can lose their structure. Large or
unsupported images are omitted. Old `.doc` files require conversion to `.docx`
outside this plugin. Password-protected documents are not supported.

Excel formulas use saved results, which can be stale or absent. This plugin does
not calculate formulas or run macros. It does not show charts, workbook styling
or hidden sheets. Merged cells span their visible rows and columns. Hidden rows
and columns within a visible sheet remain readable. Column widths are bounded
for readability. The grid estimates row heights and caps them at 260 pixels;
select a cell to read its complete preview value. Fit width keeps a minimum
readable scale, so very wide sheets still scroll horizontally.

PDF pages render locally with PDF.js and @napi-rs/canvas. Page images retain
layout and embedded content. A separate text view supports selection and copy.
Image-only pages have no selectable text; the plugin does not perform OCR.
Images are limited to 1,600 × 2,200 pixels and 5 MB per page. PDF source images
above 8 million pixels may be omitted by the renderer. Forms, links, scripts
and attachments are not interactive. Hosts must support @napi-rs/canvas's native
binary. No Poppler, LibreOffice or system browser installation is required.

## Access and limits

- Reads files within the selected workspace only, after symlink resolution.
- Skips secret directories, Git metadata, dependency folders and environment files.
- Does not upload documents to a third-party service or fetch document resources.
- Transfers preview data through the existing Paseo connection.
- Does not save converted copies, modify files or add an HTTP listener.
- Opens at most two readers at once, with a 15-second timeout and bounded worker memory.
- Accepts files up to 10 MB and Office ZIP containers up to 50 MB expanded.
- Shows up to 2,000 Word blocks / 500,000 text characters and bounded embedded images.
- Shows PDF pages one at a time, up to page 10,000.
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
notices. `npm pack` includes the generated worker. The local module locator is
excluded from the archive so the package does not contain build-host paths.
Always run `npm run build` after extracting or moving a local installation.
The Git install recipe does this automatically. The shrinkwrap file pins the
dependencies for both Git and npm distribution.

The parser tests use synthetic documents. The browser harness mounts the real
plugin components with a test adapter for Paseo's host-owned modal and RPC hooks.
It checks desktop width, merged cells, Word search, PDF pages, sheet and range
controls, errors, compact layout and dark theme. It stores screenshots in ignored `output/`.

Set `PLAYWRIGHT_CHROMIUM_EXECUTABLE` to use an existing Chromium executable.
Native device gestures and actual Paseo client acceptance require a live check.

## Dependencies

Mammoth extracts Word content. Parse5 converts its generated HTML into a limited
native reading model. SheetJS CE reads Excel data. Yauzl validates Office ZIP
entries before parsing. Esbuild creates the bounded worker. These libraries are
used on the daemon; PDF.js and @napi-rs/canvas render PDFs there too. Client code
uses only Paseo's supported host modules.

The plugin is MIT licensed. Bundled dependency licenses are included in
`server/generated/THIRD_PARTY_NOTICES.txt`.

## Share

The source is suitable for a standalone Git repository. The npm archive can be
extracted and installed as a local plugin directory. `OVERVIEW.md` supplies the
registry listing text. Public GitHub/npm publication and registry submission are
separate from building the package.
