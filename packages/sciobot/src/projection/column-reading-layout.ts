import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import type { BlockFragmentRecord } from '@docx-editor.dev/core/layout';
import type { ColumnGeometry } from '../blocks/column-geometry';

/** Read owned columns in logical order, including columns split over several pages. */
export function columnReadingLayout(
  layout: ExportSemanticLayout,
  geometries: readonly ColumnGeometry[]
): ExportSemanticLayout {
  if (!geometries.length) return layout;
  const owned = new Map(geometries.map((geometry) => [geometry.tableId, geometry]));
  const separators = new Set(geometries.flatMap((geometry) => [...geometry.separatorIds]));
  const cells = new Map<string, BlockFragmentRecord[]>();
  let visits = 0;
  const collect = (blocks: readonly BlockFragmentRecord[], depth: number) => {
    if (depth > 64) throw new Error('Column reading depth exceeded.');
    for (const block of blocks) {
      if (++visits > 100_000) throw new Error('Column reading limit exceeded.');
      if (block.kind !== 'table') continue;
      for (const row of block.rows) {
        if (row.isHeaderRepeat) continue;
        for (const cell of row.cells) {
          if (owned.has(block.tableId)) {
            const fragments = cells.get(cell.id) ?? [];
            for (const fragment of cell.blocks) fragments.push(fragment);
            cells.set(cell.id, fragments);
          }
          collect(cell.blocks, depth + 1);
        }
      }
    }
  };
  for (const page of layout.pages) collect(page.fragments, 0);
  const emitted = new Set<string>();
  visits = 0;
  const read = (blocks: readonly BlockFragmentRecord[], depth: number): BlockFragmentRecord[] => {
    if (depth > 64) throw new Error('Column reading depth exceeded.');
    return blocks.flatMap((block): BlockFragmentRecord[] => {
      if (++visits > 100_000) throw new Error('Column reading limit exceeded.');
      if (block.kind === 'paragraph') return separators.has(block.paragraphId) ? [] : [block];
      const geometry = owned.get(block.tableId);
      if (geometry) {
        if (emitted.has(block.tableId)) return [];
        emitted.add(block.tableId);
        return geometry.columns.flatMap((column) =>
          read(cells.get(column.cellId) ?? [], depth + 1)
        );
      }
      return [
        {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({ ...cell, blocks: read(cell.blocks, depth + 1) })),
          })),
        },
      ];
    });
  };
  return {
    ...layout,
    pages: layout.pages.map((page) => ({ ...page, fragments: read(page.fragments, 0) })),
  };
}
