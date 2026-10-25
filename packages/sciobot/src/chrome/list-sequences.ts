import { resetRemovedListSequences } from './list-removal-sequences';
import { parentNodeOf, type OoxmlNode } from '@docx-editor.dev/core/store';
import {
  createListSourceReader,
  planListMembership,
  type ListMembershipPatch,
  type ListPlanSource,
} from '../diff/lists';
import type { ListSourceState, ListStyle, ListTarget } from '../diff/list-types';

type Numbered = Extract<ListTarget, { kind: 'ordered' | 'bullet' }>;
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

/** Plan the affected sequence as one native list gesture; counter definitions stay immutable. */
export function planManualListSequences(
  source: ListPlanSource,
  input: readonly ListMembershipPatch[]
) {
  const validated = planListMembership(source, input);
  if (!validated.success) return validated;
  const part = source.pkg.parts.get(source.partName)!;
  const reader = createListSourceReader(source.pkg, part);
  const selected = new Set(input.map((patch) => patch.paragraphId));
  const patches = new Map(input.map((patch) => [patch.paragraphId, patch.target]));
  const states = new Map<string, ListSourceState>();
  const parents = new Map<string, OoxmlNode>();
  for (const patch of input) {
    const parent = parentNodeOf(part, patch.paragraphId);
    if (parent) parents.set(parent.id, parent);
  }
  const stateOf = (id: string) => {
    let value = states.get(id);
    if (!value) {
      const read = reader.read(id);
      if (!read.success) throw new Error(read.error);
      value = read.state;
      states.set(id, value);
    }
    return value;
  };
  const styleOf = (state: ListSourceState): ListStyle | 'todo' | null => {
    if (state.checkbox) return 'todo';
    if (!state.numbering) return null;
    const level = reader.level(state.numbering.numId, state.numbering.level)?.level;
    if (!level || level.picBulletId !== undefined || level.vanish) return null;
    return level.numFmt === 'bullet'
      ? (bullets.get(level.lvlText) ?? null)
      : (formats.get(level.numFmt) ?? null);
  };
  const describe = (node: OoxmlNode | undefined) => {
    if (node?.kind !== 'paragraph') return null;
    const state = stateOf(node.id),
      patch = patches.get(node.id);
    const layout = reader.paragraph(node.id);
    const heading = layout?.outlineLevel !== null && layout?.outlineLevel !== undefined;
    const nativeLevel = state.checkbox?.level ?? state.numbering?.level;
    const plain = (layout?.indent.left ?? 0) / 18 - 1;
    const depth = patch
      ? patch.kind === 'none'
        ? null
        : patch.level
      : (nativeLevel ?? (Number.isInteger(plain) && plain >= 0 ? plain : null));
    const style = patch
      ? patch.kind === 'none'
        ? null
        : patch.kind === 'checklist'
          ? 'todo'
          : (patch.style ??
            styleOf({
              ...state,
              checkbox: null,
              numbering:
                patch.instance.kind === 'existing'
                  ? { numId: patch.instance.numId, level: patch.level }
                  : state.numbering,
            }))
      : styleOf(state);
    const target: ListTarget | null =
      patch ??
      (state.checkbox
        ? { kind: 'checklist', level: state.checkbox.level, checked: state.checkbox.checked }
        : state.numbering && style && style !== 'todo'
          ? {
              kind: ['disc', 'circle', 'square'].includes(style) ? 'bullet' : 'ordered',
              level: state.numbering.level,
              instance: { kind: 'existing', numId: state.numbering.numId },
            }
          : null);
    return { node, state, depth, style, heading, target };
  };
  type Description = NonNullable<ReturnType<typeof describe>>;
  const numbered = (row: Description): Numbered | null =>
    row.target && (row.target.kind === 'bullet' || row.target.kind === 'ordered')
      ? row.target
      : null;
  const restart = (row: Description) => {
    const member = row.state.numbering;
    const native = member && reader.level(member.numId, member.level);
    return native
      ? (native.startOverride ?? (native.level.start !== 1 ? native.level.start : undefined))
      : undefined;
  };
  const sameDefinition = (a: Description, b: Description) => {
    const first = numbered(a),
      second = numbered(b);
    if (
      !first ||
      !second ||
      first.instance.kind !== 'existing' ||
      second.instance.kind !== 'existing'
    )
      return false;
    if (first.instance.numId === second.instance.numId) return true;
    const shape = reader.levelShape(first.instance.numId, first.level);
    if (!shape || shape !== reader.levelShape(second.instance.numId, second.level)) return false;
    return (
      JSON.stringify(reader.level(first.instance.numId, first.level)?.level) ===
      JSON.stringify(reader.level(second.instance.numId, second.level)?.level)
    );
  };
  const assign = (row: Description, reference: Description) => {
    const target = numbered(reference)!;
    patches.set(row.node.id, {
      kind: target.kind,
      level: row.depth!,
      instance: target.instance,
      ...(target.instance.kind === 'new' ? { style: reference.style as ListStyle } : {}),
    });
  };
  try {
    for (const parent of parents.values()) {
      if (parent.kind === 'textValue') continue;
      const nodes = parent.children;
      const sibling = (at: number, direction: -1 | 1, current: Description) => {
        for (let index = at + direction; index >= 0 && index < nodes.length; index += direction) {
          const row = describe(nodes[index]);
          if (!row || row.depth === null || row.depth < current.depth!) return null;
          if (row.depth === current.depth && row.style === current.style) {
            if (row.heading !== current.heading) return null;
            return { index, row };
          }
        }
        return null;
      };
      for (let index = 0; index < nodes.length; index++) {
        const node = nodes[index]!;
        if (!selected.has(node.id)) continue;
        let row = describe(node)!;
        const target = numbered(row);
        if (!target) continue;
        const previous = sibling(index, -1, row);
        const originalStyle = styleOf(row.state);
        const own = row.state.numbering;
        if (
          previous &&
          numbered(previous.row) &&
          !(
            restart(row) !== undefined &&
            own?.numId !==
              (numbered(previous.row)!.instance.kind === 'existing'
                ? (numbered(previous.row)!.instance as { kind: 'existing'; numId: string }).numId
                : null)
          ) &&
          (originalStyle !== row.style ||
            sameDefinition(
              {
                ...row,
                target: own
                  ? { ...target, instance: { kind: 'existing', numId: own.numId } }
                  : target,
              },
              previous.row
            ))
        ) {
          const prior = numbered(previous.row)!;
          const next = prior.instance.kind === 'new' ? sibling(index, 1, row) : null;
          if (
            next &&
            !selected.has(next.row.node.id) &&
            numbered(next.row)?.instance.kind === 'existing' &&
            restart(next.row) === undefined
          ) {
            const key = prior.instance.kind === 'new' ? prior.instance.key : null;
            const reference = numbered(next.row)!;
            for (const [id, patch] of patches) {
              if (
                (patch.kind === 'ordered' || patch.kind === 'bullet') &&
                patch.instance.kind === 'new' &&
                patch.instance.key === key
              )
                patches.set(id, {
                  kind: patch.kind,
                  level: patch.level,
                  instance: reference.instance,
                });
            }
            assign(row, next.row);
          } else assign(row, previous.row);
        } else if (own && originalStyle === row.style) {
          const { style: _style, ...base } = target;
          patches.set(node.id, { ...base, instance: { kind: 'existing', numId: own.numId } });
        } else {
          const next = sibling(index, 1, row);
          if (
            next &&
            !patches.has(next.row.node.id) &&
            numbered(next.row)?.instance.kind === 'existing' &&
            restart(next.row) === undefined
          )
            assign(row, next.row);
          else
            patches.set(node.id, {
              ...target,
              instance: { kind: 'new', key: `manual:${source.revision.revision}:${node.id}` },
            });
        }
        row = describe(node)!;
        let at = index;
        for (;;) {
          const next = sibling(at, 1, row);
          if (
            !next ||
            selected.has(next.row.node.id) ||
            !numbered(next.row) ||
            !sameDefinition(row, next.row)
          )
            break;
          const nextTarget = numbered(next.row)!,
            chosen = numbered(row)!;
          if (
            restart(next.row) !== undefined &&
            !(
              nextTarget.instance.kind === 'existing' &&
              chosen.instance.kind === 'existing' &&
              nextTarget.instance.numId === chosen.instance.numId
            )
          )
            break;
          assign(next.row, row);
          at = next.index;
        }
      }
    }
    return planListMembership(
      source,
      resetRemovedListSequences(
        source,
        [...patches].map(([paragraphId, target]) => ({ paragraphId, target }))
      )
    );
  } catch (error) {
    return {
      success: false as const,
      code: 'unsupported-operation' as const,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
