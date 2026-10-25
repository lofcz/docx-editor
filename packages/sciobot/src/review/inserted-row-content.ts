import {
  applyNewCellProperties,
  cellPropertiesAt,
  type NewCellProperties,
} from '../diff/new-cell-properties';
import {
  applyDirectListAssignments,
  type DirectListAssignment,
  type DirectListContext,
} from '../ai/direct-lists';
import { newCellListAssignments } from '../diff/new-cell-list-assignments';
import { cellListBindings } from '../diff/inserted-cell-lists';
import { cellAtomBindings } from '../diff/inserted-cell-atoms';
import type { InsertedListBinding } from '../diff/inserted-list-types';
import {
  applyTreeOp,
  ensureHyperlinkRelationship,
  findNode,
  paragraphTextOf,
  withPart,
  type OoxmlPackage,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { InsertedTableRowIntent } from '../diff/inserted-row-types';
import { newParagraphSeeds, type NewParagraphLink } from './new-paragraphs';
import { completeSiblingFragments, siblingFragmentSeeds } from '../diff/fragment-seeds';
import type { InsertedBlockBinding } from '../diff/inserted-blocks';
import type { TargetAstNode } from '../projection/parse-target';

export interface InsertedRowSeeds {
  readonly cells: readonly {
    readonly paragraphs: readonly OoxmlParagraphNode[];
    readonly links: readonly NewParagraphLink[];
    readonly lists?: readonly InsertedListBinding[];
    readonly properties?: NewCellProperties;
  }[];
  readonly heightTwips?: number;
}

/**
 * Seed NEW row cells from paragraph and heading streams. Payload records, tables, code, and
 * opaque atoms need the shared fragment materializer (D28/D29) and refuse before any write.
 */
export function insertedRowSeeds(
  intent: InsertedTableRowIntent
): { success: true; seeds: InsertedRowSeeds } | NativeEditFailure {
  const cells: InsertedRowSeeds['cells'][number][] = [];
  for (const [index, stream] of intent.cells.entries()) {
    const rich = stream.some((node) => node.type !== 'paragraph' && node.type !== 'heading');
    const seeds = rich ? null : newParagraphSeeds(stream);
    if (seeds && !seeds.success) return seeds;
    const payloads = intent.payloads
      .filter((binding) => binding.path[0] === index)
      .map((binding) => ({ ...binding, path: binding.path.slice(1) }));
    const atoms = cellAtomBindings(intent.atoms, index);
    if (rich && (intent.lists ?? []).some((binding) => binding.path[0] === index))
      return {
        success: false,
        code: 'unsupported-operation',
        error: 'Lists inside rich new cells need nested fragment list placement.',
      };
    cells.push({
      paragraphs: seeds?.success ? seeds.paragraphs : [],
      links: seeds?.success ? seeds.links : [],
      ...(rich
        ? {
            fragments: {
              inserted: stream,
              blocks: payloads as InsertedBlockBinding[],
              atoms,
            },
          }
        : { lists: cellListBindings(intent.lists, index) }),
      properties: cellPropertiesAt(intent.cellProperties, index),
    });
  }
  return {
    success: true,
    seeds: {
      cells,
      ...(intent.properties.heightTwips === undefined
        ? {}
        : { heightTwips: intent.properties.heightTwips }),
    },
  };
}

export interface NewCellSeeds {
  readonly properties?: NewCellProperties;
  readonly lists?: readonly InsertedListBinding[];
  readonly paragraphs: readonly OoxmlParagraphNode[];
  readonly links: readonly NewParagraphLink[];
  readonly fragments?: {
    readonly inserted: readonly TargetAstNode[];
    readonly blocks: readonly InsertedBlockBinding[];
    readonly atoms: readonly import('../diff/inserted-atoms').InsertedInlineAtom[];
  };
}

/** Seed one cell stream of paragraphs and headings; other fragments need D28/D29. */
export function newCellSeeds(
  stream: readonly import('../projection/parse-target').TargetAstNode[],
  lists?: readonly InsertedListBinding[],
  properties?: NewCellProperties,
  fragments?: NewCellSeeds['fragments']
): { success: true; seeds: NewCellSeeds } | NativeEditFailure {
  const seeds = newParagraphSeeds(stream);
  return seeds.success
    ? {
        success: true,
        seeds: {
          paragraphs: seeds.paragraphs,
          links: seeds.links,
          ...(lists ? { lists } : {}),
          ...(properties ? { properties } : {}),
          ...(fragments ? { fragments } : {}),
        },
      }
    : seeds;
}

/**
 * Fill cells that this transaction created. The enclosing row or cell insertion revision owns
 * this content, so the writes carry no inline revisions. Uses the same lowering as A01.
 */
export function materializeNewCells(
  input: OoxmlPackage,
  partName: string,
  cells: readonly { readonly cellId: string; readonly seeds: NewCellSeeds }[],
  listContext?: DirectListContext
): OoxmlPackage {
  let pkg = input,
    part = input.parts.get(partName);
  if (!part) throw new Error('The new cell part is missing.');
  const apply = (op: Parameters<typeof applyTreeOp>[1]) => {
    const result = applyTreeOp(part!, op);
    if (!result.ok) throw new Error(`New cell content refused: ${result.reason}.`);
    part = result.part;
    return result;
  };
  const assignments: DirectListAssignment[] = [];
  for (const { cellId, seeds } of cells) {
    const { paragraphs, links } = seeds;
    const cell = findNode(part, cellId);
    if (cell?.kind !== 'tableCell') throw new Error('A new native cell is missing.');
    const placeholders = cell.children.filter((node) => node.kind === 'paragraph');
    if (placeholders.length !== 1 || paragraphTextOf(part, placeholders[0]!.id) !== '')
      throw new Error('A new native cell does not hold one empty generated paragraph.');
    const before = new Set(placeholders.map((node) => node.id));
    if (seeds.fragments) {
      const placeholder = placeholders[0]!;
      const edit = {
        kind: 'siblings' as const,
        partName,
        parentId: cellId,
        beforeId: null,
        afterId: placeholder.id,
        removedIds: [],
        inserted: seeds.fragments.inserted,
        ...(seeds.lists ? { lists: seeds.lists } : {}),
        ...(seeds.fragments.blocks.length ? { blocks: seeds.fragments.blocks } : {}),
        ...(seeds.fragments.atoms.length ? { atoms: seeds.fragments.atoms } : {}),
      };
      const fragmentSeeds = siblingFragmentSeeds(edit, true);
      if (!fragmentSeeds.success)
        throw new Error(`New cell fragments refused: ${fragmentSeeds.error}.`);
      if (fragmentSeeds.seeds.paragraphs.length)
        apply({
          op: 'insertParagraphs',
          parentId: cellId,
          beforeId: placeholder.id,
          paragraphs: fragmentSeeds.seeds.paragraphs,
        });
      const completed = completeSiblingFragments(
        withPart(pkg, part),
        partName,
        edit,
        fragmentSeeds.seeds,
        fragmentSeeds.seeds.paragraphs.map((paragraph) => paragraph.id)
      );
      pkg = completed.pkg;
      part = pkg.parts.get(partName)!;
      continue;
    }
    if (!paragraphs.length) continue;
    // New paragraphs replace the generated seed paragraph.
    apply({ op: 'insertParagraphs', parentId: cellId, beforeId: placeholders[0]!.id, paragraphs });
    apply({ op: 'deleteBlock', blockId: placeholders[0]!.id });
    const landed = findNode(part, cellId);
    const ids =
      landed && landed.kind !== 'textValue'
        ? landed.children
            .filter((node) => node.kind === 'paragraph' && !before.has(node.id))
            .map((node) => node.id)
        : [];
    if (ids.length !== paragraphs.length)
      throw new Error('New cell paragraphs did not land one for one.');
    assignments.push(...newCellListAssignments(ids, seeds.lists));
    for (const link of links) {
      const target: { relationshipId: string } | { anchor: string } = (() => {
        if (link.target.kind === 'bookmark') return { anchor: link.target.anchor };
        const resource = ensureHyperlinkRelationship(pkg, link.target.url, partName);
        if (!resource) throw new Error('A new native hyperlink target was refused.');
        pkg = resource.pkg;
        return { relationshipId: resource.relationshipId };
      })();
      apply({
        op: 'insertHyperlink',
        paragraphId: ids[link.paragraphIndex]!,
        start: link.start,
        end: link.end,
        ...target,
        ...(link.title !== null ? { tooltip: link.title } : {}),
      });
    }
  }
  let filled = withPart(pkg, part);
  for (const { cellId, seeds } of cells)
    filled = applyNewCellProperties(filled, partName, cellId, seeds.properties);
  return assignments.length
    ? applyDirectListAssignments(filled, partName, assignments, listContext)
    : filled;
}

/** Fill the cells of one row that this transaction created, after its exact height. */
export function materializeInsertedRow(
  input: OoxmlPackage,
  partName: string,
  tableId: string,
  rowId: string,
  seeds: InsertedRowSeeds,
  listContext?: DirectListContext
): OoxmlPackage {
  let part = input.parts.get(partName);
  if (!part) throw new Error('The inserted row part is missing.');
  if (seeds.heightTwips !== undefined) {
    const sized = applyTreeOp(part, {
      op: 'setTableRowHeight',
      tableId,
      rowId,
      heightTwips: seeds.heightTwips,
    });
    if (!sized.ok) throw new Error(`New row height refused: ${sized.reason}.`);
    part = sized.part;
  }
  const row = findNode(part, rowId);
  if (row?.kind !== 'tableRow') throw new Error('The inserted native row is missing.');
  const cells = row.children.filter((node) => node.kind === 'tableCell');
  if (cells.length !== seeds.cells.length)
    throw new Error('The inserted native row has a different physical cell count.');
  return materializeNewCells(
    withPart(input, part),
    partName,
    cells.map((cell, index) => ({ cellId: cell.id, seeds: seeds.cells[index]! })),
    listContext
  );
}
