export const limits = {
  fileBytes: 10 * 1024 * 1024, expandedBytes: 50 * 1024 * 1024,
  entryBytes: 20 * 1024 * 1024, zipEntries: 5000,
  rows: 10000, columns: 256, pageRows: 100, pageColumns: 20,
  blocks: 2000, text: 500000, cellText: 2000,
} as const;
