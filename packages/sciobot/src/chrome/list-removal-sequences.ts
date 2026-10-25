import { listCounterAncestry } from '../diff/list-counter-ancestry';
import { parentNodeOf } from '@docx-editor.dev/core/store';
import {
  createListSourceReader,
  type ListMembershipPatch,
  type ListPlanSource,
} from '../diff/lists';
import type { ListTarget } from '../diff/list-types';

/** Fork the following counter when an unindented paragraph splits its source sequence. */
export function resetRemovedListSequences(
  source: ListPlanSource,
  input: readonly ListMembershipPatch[]
): readonly ListMembershipPatch[] {
  const part = source.pkg.parts.get(source.partName)!,
    reader = createListSourceReader(source.pkg, part);
  const patches = new Map(input.map((patch) => [patch.paragraphId, patch.target]));
  const read = (id: string) => {
    const result = reader.read(id);
    if (!result.success) throw new Error(result.error);
    return result.state;
  };
  let ancestry: ReturnType<typeof listCounterAncestry> | undefined;
  for (const patch of input) {
    if (patch.target.kind !== 'none') continue;
    const old = read(patch.paragraphId).numbering;
    if (!old) continue;
    const oldLevel = reader.level(old.numId, old.level)?.level;
    if (!oldLevel || oldLevel.numFmt === 'bullet' || oldLevel.numFmt === 'none') continue;
    const parent = parentNodeOf(part, patch.paragraphId);
    if (!parent) continue;
    const at = parent.children.findIndex((child) => child.id === patch.paragraphId);
    let key: string | undefined;
    let ancestorStarts: readonly number[] | undefined;
    let heading: boolean | undefined;
    for (let index = at + 1; index < parent.children.length; index++) {
      const node = parent.children[index]!;
      if (node.kind !== 'paragraph') break;
      const current = read(node.id),
        selected = patches.get(node.id);
      if (selected?.kind === 'none') break;
      if (selected) continue;
      const member = current.numbering;
      const depth = current.checkbox?.level ?? member?.level;
      if (depth === undefined || depth < old.level) break;
      if (!member || member.numId !== old.numId) continue;
      const layout = reader.paragraph(node.id),
        isHeading = layout?.outlineLevel !== null && layout?.outlineLevel !== undefined;
      if (member.level === old.level) {
        if (heading !== undefined && heading !== isHeading) break;
        heading = isHeading;
        if (!key) {
          key = `removed:${source.revision.revision}:${node.id}`;
          if (old.level > 0) {
            ancestry ??= listCounterAncestry(source, input);
            ancestorStarts = ancestry.get(node.id)?.slice(0, old.level);
            if (ancestorStarts?.length !== old.level)
              throw new Error('The removed list tail has no complete native parent counters.');
          }
        }
      }
      if (!key) continue;
      const level = reader.level(member.numId, member.level)?.level;
      if (!level || level.numFmt === 'none')
        throw new Error('The removed list tail has an unsupported counter level.');
      const target: ListTarget = {
        kind: level.numFmt === 'bullet' ? 'bullet' : 'ordered',
        level: member.level,
        instance: {
          kind: 'new',
          key,
          templateNumId: old.numId,
          ...(ancestorStarts?.length ? { ancestorStarts } : {}),
          ...(member.level === old.level ? { start: 1 } : {}),
        },
      };
      patches.set(node.id, target);
    }
  }
  return [...patches].map(([paragraphId, target]) => ({ paragraphId, target }));
}
