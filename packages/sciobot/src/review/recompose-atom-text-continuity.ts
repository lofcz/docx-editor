import { activeMdxRegistry } from '../registry/active';
import {
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  contentControlLevelOf,
  findNode,
  replaceNode,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import { blockIdFromTag } from '../blocks/payload-opc-codec';
import { refuseWrite } from '../diff/fragment-seeds';

/** Inline components are the atoms a paragraph's text gaps sit between. */
const isAtom = (type: string | undefined) =>
  type !== undefined && activeMdxRegistry().components.get(type)?.kind === 'inline';
type Text = Extract<OoxmlNode, { kind: 'text' }>;
interface Inventory {
  readonly atoms: readonly string[];
  readonly gaps: readonly (readonly Text[])[];
}
/** Native text gaps are addressed by ordered atom slots, never by matching labels. */
export function ownedAtomGapInventory(
  paragraph: OoxmlParagraphNode,
  typeOf: (id: string) => string | undefined,
  selected?: ReadonlySet<string>
): Inventory | null {
  const atoms: string[] = [],
    gaps: Text[][] = [[]];
  const visit = (node: OoxmlNode): boolean => {
    if (node.kind === 'paragraphProperties' || node.kind === 'runProperties') return true;
    if (node.kind === 'contentControl') {
      const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
      const type = id === null ? undefined : typeOf(id);
      if (
        id === null ||
        !type ||
        !isAtom(type) ||
        contentControlLevelOf(node) !== 'inline' ||
        (selected && !selected.has(id))
      )
        return false;
      atoms.push(type);
      gaps.push([]);
      return true;
    }
    if (node.kind === 'text') {
      if (node.children.length !== 1 || node.children[0]!.kind !== 'textValue') return false;
      gaps.at(-1)!.push(node);
      return true;
    }
    if (!['paragraph', 'run', 'revisionInsert', 'hyperlink'].includes(node.kind)) return false;
    return node.kind !== 'textValue' && node.children.every(visit);
  };
  return visit(paragraph) && atoms.length ? { atoms, gaps } : null;
}

/**
 * Preserve the native text identities of unchanged gaps in the continued paragraph.
 * Position and atom kinds establish the correspondence; full native text equality only
 * verifies it. Target run formatting and revision wrappers remain newly staged content.
 */
export function continueOwnedAtomGapText(
  part: OoxmlPart,
  source: OoxmlParagraphNode,
  targetId: string,
  sourceType: (id: string) => string | undefined,
  targetType: (id: string) => string | undefined,
  selected: ReadonlySet<string>
): OoxmlPart {
  const target = findNode(part, targetId);
  if (target?.kind !== 'paragraph')
    refuseWrite('The continued atom paragraph disappeared.', 'apply-failed');
  const before = ownedAtomGapInventory(source, sourceType, selected),
    after = ownedAtomGapInventory(target, targetType);
  if (!before || !after || JSON.stringify(before.atoms) !== JSON.stringify(after.atoms))
    return part;
  let result = part;
  for (const [gapIndex, oldGap] of before.gaps.entries()) {
    const newGap = after.gaps[gapIndex]!;
    if (
      oldGap.length !== newGap.length ||
      oldGap.some(
        (node, index) =>
          canonicalOoxmlFingerprint(node) !== canonicalOoxmlFingerprint(newGap[index]!)
      )
    )
      continue;
    for (const [index, node] of oldGap.entries()) {
      const next = newGap[index]!;
      if (
        node.id === next.id ||
        findNode(result, node.id) ||
        node.kind !== 'text' ||
        node.children.some((child) => findNode(result, child.id))
      )
        continue;
      const replaced = replaceNode(result, next.id, node, { deferValidation: true });
      if (!replaced.ok)
        refuseWrite('Atom text continuity violates the native tree.', 'apply-failed');
      result = replaced.part;
    }
  }
  return result;
}
