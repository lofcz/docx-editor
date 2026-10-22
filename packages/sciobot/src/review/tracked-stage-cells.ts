import { findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { TargetAstNode } from '../projection/parse-target';
import { newParagraphSeeds } from './new-paragraphs';
import { stageParagraphResources } from './tracked-stage-paragraphs';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageWriter } from './tracked-stage-writer';

/** Fill actual newly allocated cells. The native row/cell insertion owns their complete content. */
export function stageNewCells(
  writer: TrackedStageWriter,
  partName: string,
  cells: readonly { cellId: string; blocks: readonly TargetAstNode[] }[]
): void {
  for (const { cellId, blocks } of cells) {
    const seeds = newParagraphSeeds(blocks, { code: true });
    if (!seeds.success) refuseWrite(seeds.error, seeds.code);
    if (!seeds.paragraphs.length) continue;
    const part = currentStagePart(writer, partName),
      cell = findNode(part, cellId);
    if (cell?.kind !== 'tableCell') refuseWrite('A new native cell is missing.', 'apply-failed');
    const placeholders = cell.children.filter((node) => node.kind === 'paragraph');
    if (placeholders.length !== 1 || paragraphTextOf(part, placeholders[0]!.id) !== '')
      refuseWrite('A new native cell needs one empty generated paragraph.', 'apply-failed');
    const old = new Set(cell.children.map((node) => node.id));
    writer.apply(partName, {
      op: 'insertParagraphs',
      parentId: cellId,
      beforeId: placeholders[0]!.id,
      paragraphs: seeds.paragraphs,
    });
    writer.apply(partName, { op: 'deleteBlock', blockId: placeholders[0]!.id });
    const landed = findNode(currentStagePart(writer, partName), cellId);
    if (landed?.kind !== 'tableCell') refuseWrite('A new native cell disappeared.', 'apply-failed');
    const paragraphs = landed.children
      .filter((node) => node.kind === 'paragraph' && !old.has(node.id))
      .map((node) => node.id);
    if (paragraphs.length !== seeds.paragraphs.length)
      refuseWrite('New cell paragraphs did not land one for one.', 'apply-failed');
    stageParagraphResources(writer, partName, paragraphs, seeds);
  }
}
