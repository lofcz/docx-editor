import { findNode, type OoxmlNode } from '@docx-editor.dev/core/store';
import type { DirectListContext } from '../ai/direct-lists';
import { contentWidth, refuseWrite, type SiblingFragmentSeeds } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import { stageNewCells } from './tracked-stage-cells';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

const children = (node: OoxmlNode | null | undefined): readonly OoxmlNode[] =>
  node && node.kind !== 'textValue' ? node.children : [];

/** New table rows own their complete cells; fill only these newly allocated native owners. */
export function stageSiblingTables(
  writer: TrackedStageWriter,
  edit: Extract<CanonicalEdit, { kind: 'siblings' }>,
  seeds: SiblingFragmentSeeds,
  fresh: readonly string[],
  revision: TrackedStageAttribution,
  listContext: DirectListContext
): void {
  for (const seed of seeds.placed) {
    if (seed.kind !== 'table')
      refuseWrite(
        'Tracked application block ownership needs its own stage.',
        'unsupported-operation'
      );
    const anchor = seed.anchor === null ? edit.afterId : fresh[seed.anchor]!;
    const before = currentStagePart(writer, edit.partName);
    if (!anchor || findNode(before, anchor)?.kind !== 'paragraph')
      refuseWrite(
        'A new tracked table needs a following native paragraph.',
        'unsupported-operation'
      );
    const prior = new Set(children(findNode(before, edit.parentId)).map((node) => node.id));
    const columns = seed.rows[0]!.length;
    writer.apply(edit.partName, {
      op: 'insertTable',
      beforeParagraphId: anchor,
      rows: seed.rows.length,
      cols: columns,
      columnWidthTwips: Math.max(240, Math.floor(contentWidth(before) / columns)),
      revision,
    });
    const current = currentStagePart(writer, edit.partName);
    const tables = children(findNode(current, edit.parentId)).filter(
      (node) => node.kind === 'table' && !prior.has(node.id)
    );
    if (tables.length !== 1) refuseWrite('Tracked table allocation is ambiguous.', 'apply-failed');
    const rows = children(tables[0]).filter((node) => node.kind === 'tableRow');
    if (rows.length !== seed.rows.length)
      refuseWrite('The new tracked table row count differs.', 'apply-failed');
    const cells = rows.flatMap((row, r) => {
      const rowCells = children(row).filter((node) => node.kind === 'tableCell');
      if (rowCells.length !== columns)
        refuseWrite('The new tracked table cell count differs.', 'apply-failed');
      return rowCells.map((cell, c) => ({
        cellId: cell.id,
        blocks: seed.rows[r]![c]!,
        properties: seed.alignment[c]
          ? { horizontal: seed.alignment[c] as 'left' | 'center' | 'right' }
          : undefined,
      }));
    });
    stageNewCells(writer, edit.partName, cells, listContext, revision);
  }
}
