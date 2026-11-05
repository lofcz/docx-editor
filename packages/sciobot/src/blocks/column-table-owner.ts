import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from './block-index';
import { readColumnGeometry } from './column-geometry';

/** Cache only immutable packages. Nested ordinary tables retain their native controls. */
export function createColumnTableRefusal(
  reason: string
): (tableId: string, pkg: OoxmlPackage) => string | null {
  const cache = new WeakMap<OoxmlPackage, Set<string>>();
  return (tableId, pkg) => {
    let tables = cache.get(pkg);
    if (!tables) {
      tables = new Set();
      cache.set(pkg, tables);
      const index = buildBlockIndex(pkg);
      for (const block of index.byId.values()) {
        if (block.payload.type !== 'column_group') continue;
        const geometry = readColumnGeometry(pkg, block, index);
        if (geometry) tables.add(geometry.tableId);
      }
    }
    return tables.has(tableId) ? reason : null;
  };
}
