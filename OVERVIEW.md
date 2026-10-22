# Office preview

Read Word, Excel and PDF files stored on your Paseo host. Desktop preview buttons
open a full-width workspace reader. Compact preview buttons open a native sheet.
The Office preview panel lists files and accepts a path in the current workspace.

Excel previews retain column widths and merged cells. They show saved values,
worksheet tabs, zoom controls, cell inspection and row and column navigation.
Use the cell address field to open the range that contains a cell. Select a cell
to read and copy its value. Word previews provide a reading page, a heading list,
text search and adjustable text size. PDF previews show the original page layout,
with page navigation, zoom and a selectable text view.

Requires Paseo 0.10.2 or 0.11.x on the daemon and client, Node.js 22.13 or later,
and npm on the daemon host. PDF rendering uses a local native canvas library;
the host must have a platform supported by @napi-rs/canvas. No external document
service or account is required.

The plugin reads files only within the selected workspace, after symlink checks.
It sends previews through the existing Paseo connection. It does not upload files,
fetch external document resources or change source files. All Paseo plugins run
as trusted code with the daemon user's access.

Word uses a reading view. It does not reproduce page layout, fonts, headers,
footers or tracked changes exactly. Excel charts and cell styles are not shown.
Formulas use saved results and are not recalculated. Hidden sheets are omitted.
PDF page images preserve appearance; text can be copied in the separate text
view. Scanned PDFs do not gain searchable text. Old .doc files and encrypted
files are not supported.

Files are limited to 10 MB. Excel previews show up to 10,000 rows and 256 columns,
in ranges of 100 rows and 20 columns. Long cell values are shortened with a note.
PDF previews render one page at a time, up to 10,000 pages, at a bounded resolution.
Preview notes explain limits. Automatic buttons apply to new completed replies;
use the panel to open files from older replies.
