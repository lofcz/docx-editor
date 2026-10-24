import { continueOwnedAtomGapText } from './recompose-atom-text-continuity';
import {
  contentControlsIn,
  contentControlPropertiesOf,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  replaceChildren,
  moveNode,
  removeNode,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import type { OwnedProposalRemovalPlan } from './owned-proposal-plan';
import { refuseWrite } from '../diff/fragment-seeds';
import { readPayloadStore, blockIdFromTag } from '../blocks/payload-opc-codec';

const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const rows = (node: OoxmlNode) => children(node).filter((n) => n.kind === 'tableRow');
const cells = (node: OoxmlNode) => children(node).filter((n) => n.kind === 'tableCell');
const paragraphs = (node: OoxmlNode) =>
  children(node).filter((n): n is OoxmlParagraphNode => n.kind === 'paragraph');
function topology(table: OoxmlNode): string {
  const grid = children(table).find((n) => n.kind !== 'textValue' && n.localName === 'tblGrid');
  return JSON.stringify([
    grid ? children(grid).length : 0,
    rows(table).map((row) =>
      cells(row).map((cell) => {
        const props = children(cell).find((n) => n.kind !== 'textValue' && n.localName === 'tcPr');
        return [
          paragraphs(cell).length,
          ...(props ? children(props) : [])
            .filter((n) => n.kind !== 'textValue' && ['gridSpan', 'vMerge'].includes(n.localName))
            .map((n) => (n.kind === 'textValue' ? null : [n.localName, n.attributes])),
        ];
      })
    ),
  ]);
}
function emptyRecovery(node: OoxmlParagraphNode): boolean {
  return node.children.every(
    (props) =>
      props.kind === 'paragraphProperties' &&
      !props.attributes.length &&
      props.children.every(
        (child) =>
          !child.attributes.length &&
          !child.children.length &&
          ((child.namespaceUri === 'urn:docx-editor.dev:review:retired-paragraph:1' &&
            child.localName === 'retired') ||
            (child.namespaceUri === WML_NAMESPACE_URI && child.localName === 'rPr'))
      )
  );
}
interface Gap {
  readonly parentId: string;
  readonly left: string | null;
  readonly right: string | null;
  readonly tables: OoxmlNode[];
}

/**
 * Keep an owned atom paragraph address through a same-topology table replacement.
 * A late peer update still names that address after the owned atom embeds disappear.
 * Match complete native birth gaps and physical cells, never labels or equal text.
 */
export function continueOwnedCellParagraphs(
  before: OoxmlPackage,
  middle: OoxmlPackage,
  after: OoxmlPackage,
  removal: OwnedProposalRemovalPlan
): OoxmlPackage {
  const selectedBlocks = new Set(removal.propertyDecisions.map((member) => member.blockId));
  if (!selectedBlocks.size) return after;
  let result = after;
  for (const partOps of removal.partOps) {
    const selected = new Set(
      partOps.ops.flatMap((op) => (op.op === 'rejectAllRevisions' ? (op.siteNodeIds ?? []) : []))
    );
    if (!selected.size) continue;
    const original = before.parts.get(partOps.partName),
      baseline = middle.parts.get(partOps.partName),
      current = result.parts.get(partOps.partName);
    if (!original || !baseline || !current) continue;
    const bornRows = new Set<string>(),
      bornParagraphs = new Set<string>();
    for (const id of selected) {
      const node = findNode(original, id);
      if (
        !node ||
        node.kind === 'textValue' ||
        node.namespaceUri !== WML_NAMESPACE_URI ||
        node.localName !== 'ins'
      )
        continue;
      const parent = parentNodeOf(original, id);
      if (parent?.namespaceUri === WML_NAMESPACE_URI && parent.localName === 'trPr') {
        const row = parentNodeOf(original, parent.id);
        if (row?.kind === 'tableRow') bornRows.add(row.id);
      }
      if (parent?.namespaceUri === WML_NAMESPACE_URI && parent.localName === 'rPr') {
        const props = parentNodeOf(original, parent.id),
          paragraph = props && parentNodeOf(original, props.id);
        if (props?.kind === 'paragraphProperties' && paragraph?.kind === 'paragraph')
          bornParagraphs.add(paragraph.id);
      }
    }
    if (!bornRows.size) continue;
    const gaps = new Map<string, Gap>();
    const consider = (node: OoxmlNode) => {
      if (node.kind === 'table' && !findNode(baseline, node.id)) {
        const sourceRows = rows(node);
        if (
          sourceRows.length &&
          sourceRows.every((r) => bornRows.has(r.id)) &&
          sourceRows.every((r) =>
            cells(r).every((c) =>
              children(c).every(
                (n) => n.kind === 'paragraph' || (n.kind !== 'textValue' && n.localName === 'tcPr')
              )
            )
          )
        ) {
          const parent = parentNodeOf(original, node.id),
            baselineParent = parent && findNode(baseline, parent.id);
          if (parent && baselineParent && baselineParent.kind !== 'textValue') {
            const at = parent.children.findIndex((n) => n.id === node.id);
            const survives = (sibling: OoxmlNode) =>
              baselineParent.children.some((n) => n.id === sibling.id) &&
              !bornParagraphs.has(sibling.id);
            const left = parent.children.slice(0, at).filter(survives).at(-1)?.id ?? null;
            const right = parent.children.slice(at + 1).find(survives)?.id ?? null;
            const key = JSON.stringify([parent.id, left, right]);
            const gap = gaps.get(key) ?? { parentId: parent.id, left, right, tables: [] };
            gap.tables.push(node);
            gaps.set(key, gap);
          }
        }
      }
    };
    const candidateTables = new Map<string, OoxmlNode>();
    for (const row of bornRows) {
      const table = parentNodeOf(original, row);
      if (table?.kind === 'table') candidateTables.set(table.id, table);
    }
    candidateTables.forEach(consider);
    const replacements = new Map<string, string>(),
      slots = new Set<string>();
    for (const gap of gaps.values()) {
      const parent = findNode(current, gap.parentId);
      if (!parent || parent.kind === 'textValue') continue;
      const left = gap.left === null ? -1 : parent.children.findIndex((n) => n.id === gap.left);
      const right =
        gap.right === null
          ? parent.children.length
          : parent.children.findIndex((n) => n.id === gap.right);
      if ((gap.left !== null && left < 0) || right <= left) continue;
      const fresh = parent.children
        .slice(left + 1, right)
        .filter((n) => n.kind === 'table' && !findNode(baseline, n.id));
      if (
        fresh.length !== gap.tables.length ||
        fresh.some((n, i) => topology(n) !== topology(gap.tables[i]!))
      )
        continue;
      for (const [index, source] of gap.tables.entries()) {
        const oldCells = rows(source).flatMap(cells),
          newCells = rows(fresh[index]!).flatMap(cells);
        for (const [cellIndex, cell] of oldCells.entries()) {
          const targets = paragraphs(newCells[cellIndex]!);
          for (const [paragraphIndex, sourceParagraph] of paragraphs(cell).entries()) {
            const slot = findNode(baseline, sourceParagraph.id),
              liveSlot = findNode(current, sourceParagraph.id),
              target = targets[paragraphIndex]!;
            const ownsAtom = contentControlsIn(sourceParagraph).some(({ node }) => {
              const blockId = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
              return blockId !== null && selectedBlocks.has(blockId);
            });
            if (
              !ownsAtom ||
              !bornParagraphs.has(sourceParagraph.id) ||
              slot?.kind !== 'paragraph' ||
              liveSlot?.kind !== 'paragraph' ||
              paragraphTextOf(baseline, slot.id) !== '' ||
              !emptyRecovery(slot) ||
              paragraphTextOf(current, liveSlot.id) !== '' ||
              !emptyRecovery(liveSlot) ||
              parentNodeOf(current, liveSlot.id)?.id !== gap.parentId ||
              findNode(baseline, target.id) ||
              slots.has(slot.id)
            )
              continue;
            replacements.set(target.id, slot.id);
            slots.add(slot.id);
          }
        }
      }
    }
    if (!replacements.size) continue;
    const sourcePayloads = readPayloadStore(before),
      targetPayloads = readPayloadStore(after);
    if (sourcePayloads.status !== 'ready' || targetPayloads.status !== 'ready')
      refuseWrite('Atom continuity needs complete native payload stores.', 'invalid-document');
    let part: OoxmlPart = current;
    // Explicit moves preserve native parent ownership. Rebuilding parent lists in a
    // different order can unlink a source after the destination already lists it.
    // This edit owns a transaction boundary; validate the complete resulting part once.
    const deferred = { deferValidation: true };
    const apply = (edit: ReturnType<typeof moveNode>) => {
      if (!edit.ok)
        refuseWrite('Owned paragraph continuity violates the native tree.', 'apply-failed');
      part = edit.part;
    };
    for (const [targetId, sourceId] of replacements) {
      const target = findNode(part, targetId),
        parent = parentNodeOf(part, targetId);
      if (target?.kind !== 'paragraph' || parent?.kind !== 'tableCell')
        refuseWrite('A replacement paragraph cell disappeared.', 'apply-failed');
      apply(
        moveNode(
          part,
          sourceId,
          parent.id,
          parent.children.findIndex((n) => n.id === targetId),
          deferred
        )
      );
      const slot = findNode(part, sourceId);
      const oldProps =
        slot?.kind === 'paragraph' && slot.children.find((n) => n.kind === 'paragraphProperties');
      const newProps = target.children.find((n) => n.kind === 'paragraphProperties');
      if (
        !oldProps ||
        oldProps.kind !== 'paragraphProperties' ||
        newProps?.kind !== 'paragraphProperties' ||
        newProps.attributes.length
      )
        refuseWrite('Owned paragraph continuity needs exact property holders.', 'apply-failed');
      apply(replaceChildren(part, oldProps.id, [], deferred));
      for (const [index, child] of newProps.children.entries())
        apply(moveNode(part, child.id, oldProps.id, index, deferred));
      for (const [index, child] of target.children
        .filter((n) => n.kind !== 'paragraphProperties')
        .entries())
        apply(moveNode(part, child.id, sourceId, index + 1, deferred));
      apply(removeNode(part, targetId, deferred));
      const source = findNode(original, sourceId);
      if (source?.kind !== 'paragraph')
        refuseWrite('The original atom paragraph disappeared.', 'apply-failed');
      part = continueOwnedAtomGapText(
        part,
        source,
        sourceId,
        (id) => {
          const decoded = sourcePayloads.entries.get(id)?.decoded;
          return decoded?.status === 'supported' ? decoded.value.type : undefined;
        },
        (id) => {
          const decoded = targetPayloads.entries.get(id)?.decoded;
          return decoded?.status === 'supported' ? decoded.value.type : undefined;
        },
        selectedBlocks
      );
    }
    apply(replaceChildren(part, part.root.id, part.root.children));
    result = withPart(result, part);
  }
  return result;
}
