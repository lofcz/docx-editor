import { parentNodeOf } from '@docx-editor.dev/core/store';
import { createListSourceReader, type ListPlanSource } from '../diff/lists';
import type { ListStyle } from '../diff/list-types';
const formats = new Map<string, ListStyle>([
  ['decimal', 'decimal'],
  ['lowerLetter', 'lower-alpha'],
  ['upperLetter', 'upper-alpha'],
  ['lowerRoman', 'lower-roman'],
  ['upperRoman', 'upper-roman'],
]);
const bullets = new Map<string, ListStyle>([
  ['•', 'disc'],
  ['○', 'circle'],
  ['■', 'square'],
]);

/** Source collapsed toggles change their same-level original-style siblings. */
export function listToggleSiblings(source: ListPlanSource, paragraphId: string): readonly string[] {
  const part = source.pkg.parts.get(source.partName)!;
  const reader = createListSourceReader(source.pkg, part);
  const describe = (id: string) => {
    const read = reader.read(id);
    if (!read.success) throw new Error(read.error);
    const state = read.state;
    if (state.checkbox) return { depth: state.checkbox.level, style: 'todo' };
    if (state.numbering) {
      const level = reader.level(state.numbering.numId, state.numbering.level)?.level;
      const style =
        level?.numFmt === 'bullet' ? bullets.get(level.lvlText) : formats.get(level?.numFmt ?? '');
      return { depth: state.numbering.level, style: style ?? null };
    }
    const left = reader.paragraph(id)?.indent.left ?? 0;
    const depth = left / 18 - 1;
    return { depth: Number.isInteger(depth) && depth >= 0 ? depth : null, style: null };
  };
  const current = describe(paragraphId),
    parent = parentNodeOf(part, paragraphId);
  if (!parent || !current.style) return [paragraphId];
  const at = parent.children.findIndex((node) => node.id === paragraphId),
    ids = new Set([paragraphId]);
  for (const direction of [-1, 1] as const)
    for (
      let index = at + direction;
      index >= 0 && index < parent.children.length;
      index += direction
    ) {
      const node = parent.children[index]!;
      if (node.kind !== 'paragraph') break;
      const sibling = describe(node.id);
      if (sibling.depth === null || sibling.depth < current.depth!) break;
      if (sibling.depth !== current.depth) continue;
      if (sibling.style !== current.style) break;
      ids.add(node.id);
    }
  return parent.children.filter((node) => ids.has(node.id)).map((node) => node.id);
}
