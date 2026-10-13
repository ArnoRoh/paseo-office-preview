import { defineRpc } from '@getpaseo/plugin';
import { z } from 'zod';

import { limits } from './limits';
export { limits } from './limits';

const run = z.object({ text: z.string(), bold: z.boolean().optional(), italic: z.boolean().optional() });
export const blockSchema = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('paragraph'), runs: z.array(run), heading: z.number().optional() }),
  z.object({ kind: z.literal('table'), rows: z.array(z.array(z.array(run))) }),
  z.object({ kind: z.literal('image'), uri: z.string(), alt: z.string() }),
]);
export type Block = z.infer<typeof blockSchema>;
export type Run = z.infer<typeof run>;
const common = { name: z.string(), warnings: z.array(z.string()) };
export const previewSchema = z.discriminatedUnion('kind', [
  z.object({ ...common, kind: z.literal('document'), blocks: z.array(blockSchema) }),
  z.object({
    ...common, kind: z.literal('workbook'), sheets: z.array(z.string()), sheet: z.number(),
    rows: z.array(z.array(z.string())), rowStart: z.number(), columnStart: z.number(),
    totalRows: z.number(), totalColumns: z.number(),
  }),
]);
export type Preview = z.infer<typeof previewSchema>;
export const selectionSchema = z.object({
  sheet: z.number().int().min(0).max(99).default(0),
  rowStart: z.number().int().min(0).max(limits.rows - 1).default(0),
  columnStart: z.number().int().min(0).max(limits.columns - 1).default(0),
});
export type Selection = z.infer<typeof selectionSchema>;
export const readPreview = defineRpc({
  name: 'office-preview.read',
  input: selectionSchema.extend({ workspaceId: z.string().min(1), path: z.string().min(1).max(4096) }),
  output: previewSchema,
});
export const listFiles = defineRpc({
  name: 'office-preview.list',
  input: z.object({ workspaceId: z.string().min(1) }),
  output: z.object({ paths: z.array(z.string()), truncated: z.boolean() }),
});
export const cardSchema = z.object({ workspaceId: z.string(), paths: z.array(z.string()).max(20) });

export function columnLabel(index: number): string {
  let label = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
  return label;
}
