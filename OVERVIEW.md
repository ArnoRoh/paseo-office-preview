# Office preview

Read Word and Excel files stored on your Paseo host without leaving the conversation.
The viewer opens in a dialog on wide screens and a sheet on compact screens.

When an agent completes a reply with a local `.docx`, `.xlsx` or `.xls` link,
the plugin adds a Preview button below the reply. You can also choose Office
preview from the workspace panels or Command Center, or enter `/office-preview`.
The panel lists documents and accepts a path within the current workspace.

Word previews show text, basic emphasis, tables and supported embedded images.
Excel previews show saved cell values, worksheet tabs and row and column paging.
You can select and copy text. The viewer does not change the source file.

Requires Paseo 0.10.2 or 0.11.x on the daemon and client, with Node.js 22 or later
on the daemon host. npm is needed to prepare the plugin. No external document
service or account is required.

The plugin reads documents only within the selected workspace. Relative paths
and symlinks must resolve within that workspace. Previews travel through your
existing Paseo connection. The plugin does not fetch external images or links.
All Paseo plugins run as trusted code with the daemon user's access.

This is a reading view. Word page layout, fonts, headers, footers and tracked
changes are not reproduced exactly. Excel charts, cell styling and macros are
not shown. Formulas use saved results and are not recalculated. Hidden sheets
are omitted. Old `.doc` files and encrypted documents are not supported.

Files are limited to 10 MB. Large documents have bounded previews. Excel previews
show up to 10,000 rows and 256 columns, in pages of 100 rows and 20 columns.
Limits are shown in the viewer. Automatic buttons apply to new completed turns;
use the panel for older replies or if a button is unavailable after a restart.
