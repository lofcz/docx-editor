import type { InlineImageRuntime } from '../native/inline-image-runtime';
import type { InsertedInlineAtom } from '../diff/inserted-atoms';
import { inlineAtomsByPath, inlineAtomKey, inlineAtomXml } from '../diff/inline-atom-xml';
import { stageBornAtoms } from './tracked-born-atoms';
import { applyNewCellProperties, type NewCellProperties } from '../diff/new-cell-properties';
import {
  applyDirectListAssignments,
  type DirectListAssignment,
  type DirectListContext,
} from '../ai/direct-lists';
import { newCellListAssignments } from '../diff/new-cell-list-assignments';
import type { InsertedListBinding } from '../diff/inserted-list-types';
import { findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { TargetAstNode } from '../projection/parse-target';
import { newParagraphSeeds } from './new-paragraphs';
import { stageParagraphResources } from './tracked-stage-paragraphs';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Fill actual newly allocated cells. The native row/cell insertion owns their complete content. */
export function stageNewCells(
  writer: TrackedStageWriter,
  partName: string,
  cells: readonly {
    cellId: string;
    blocks: readonly TargetAstNode[];
    lists?: readonly InsertedListBinding[];
    properties?: NewCellProperties;
    atoms?: readonly InsertedInlineAtom[];
  }[],
  listContext?: DirectListContext,
  revision?: TrackedStageAttribution,
  images?: InlineImageRuntime
): void {
  const assignments: DirectListAssignment[] = [];
  for (const { cellId, blocks, lists, atoms = [] } of cells) {
    const bindings = inlineAtomsByPath(atoms);
    const seeds = newParagraphSeeds(blocks, {
      code: true,
      atom: (token, index) => {
        const atom = bindings.get(inlineAtomKey([index], token.path));
        return atom ? inlineAtomXml(atom, images) : null;
      },
    });
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
      ...(revision ? { revision } : {}),
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
    stageBornAtoms(writer, partName, paragraphs, atoms, images);
    assignments.push(...newCellListAssignments(paragraphs, lists));
  }
  for (const { cellId, properties } of cells)
    writer.edit((pkg) => applyNewCellProperties(pkg, partName, cellId, properties));
  if (assignments.length)
    writer.edit((pkg) => applyDirectListAssignments(pkg, partName, assignments, listContext));
}
