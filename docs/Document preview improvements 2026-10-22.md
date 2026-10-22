# Document preview improvements

Date: 2026-10-22
Status: Implemented and tested. Live update and user acceptance pending.

## Problem and decision

The existing desktop modal is too narrow for useful spreadsheet reading.
Its fixed-width columns discard the workbook's widths and merged title cells.
The screenshot supplied by the user demonstrates this problem. PDF files have
no preview route in the plugin.

Use the existing workspace panel for desktop previews. Keep the native sheet
for compact preview cards. Do not upgrade the live daemon to gain newer APIs.
The observed daemon version is Paseo 0.11.4. Its newer DocumentFrame and file
panel APIs require 0.13.0, so they cannot solve this task on the current host.

## Reference review

- [Handsontable merged cells](https://handsontable.com/docs/16.1/javascript-data-grid/merge-cells/)
  demonstrates merged ranges and spreadsheet headers. Adopt these reading
  patterns in native components; do not add the browser-only grid dependency.
- [SheetJS column properties](https://docs.sheetjs.com/docs/csf/features/colprops/)
  describes source widths in `!cols`. Retain these widths with display bounds.
- [Mozilla PDF.js viewer options](https://github.com/mozilla/pdf.js/wiki/Viewer-options)
  provides page and zoom navigation. Use these controls in the native reader.
- [Mozilla PDF.js Node renderer](https://github.com/mozilla/pdf.js/blob/master/examples/node/pdf2png/pdf2png.mjs)
  supports local PDF page rendering. Add PDF.js and @napi-rs/canvas as pinned
  dependencies. Their output travels over the existing plugin RPC.
- [DOCX.js](https://github.com/VolodymyrBaydalka/docxjs) renders document layout
  into browser HTML. This does not fit the plugin's native client contract.
  Retain Mammoth's reading model and improve layout, navigation and text search.
- [Paseo plugin reference](https://paseo.sh/docs/plugins/reference) documents
  modal constraints, workspace panels and the newer file-viewing API.

## Result

Desktop preview cards now open the workspace reader. Excel keeps column widths
and merged ranges, including ranges that cross a paging boundary. Numeric cells
align right. Users can zoom, inspect and copy a cell, or navigate by address.
Short worksheets do not show unnecessary paging controls. Worksheet tabs sit
below the grid. Preview notes are collapsible.

Word has a centered reading page, a heading list, section search and adjustable
text size. Tables use the available reading width. PDF has local page rendering,
page navigation, zoom, fit width and a separate selectable text view. PDF links
also appear in automatic cards and file discovery.

The source files remain unchanged. Workspace checks, ZIP checks, worker limits
and the prohibition on remote resources remain in place. No HTTP service is
added. Native PDF dependencies resolve from an installation-local generated
module locator. The archive excludes that locator; installation rebuilds it.
The archive contains no private files or build-host paths.

## Verification

- Build and TypeScript checks passed.
- Eleven parser and integration tests passed.
- Sixteen browser scenarios passed at desktop and compact widths, with light
  and dark themes. These use the real components and a host UI adapter.
- Visual review confirmed the seven-column synthetic sheet and merged title,
  the Word reading page, and rendered PDF content.
- The Paseo 0.11.4 compiler accepted both entrypoints. The compiled server
  handled a PDF request with the expected page image and text.
- The 0.2.0 archive installed with production dependencies only, rebuilt and
  passed the same compiled PDF request.
- The production dependency audit reports only the existing three moderate
  Mammoth CLI dependency findings. The worker does not bundle that CLI.

Synthetic fixtures and screenshots remain in ignored output directories.
Actual native gestures and live desktop acceptance remain separate checks.
No live plugin change or daemon restart was performed during implementation.

## Limits and pending action

Word remains a reading view, without exact Word page geometry. Excel cell styles
and charts are not rendered. Formulas use saved results. PDF images have a bounded
resolution; source images above eight million pixels can be omitted. Text-only
mode does not provide OCR for scanned pages. The preview remains read-only.

Request approval to replace the running office-preview plugin with this build.
After approval, verify plugin status and RPC results, then check the intended
client. A healthy subprocess alone does not establish user acceptance.

Task Mgmt capture was unavailable: connected tool discovery returned no registered
team tools or authenticated Task Mgmt endpoint. No credential was created.
