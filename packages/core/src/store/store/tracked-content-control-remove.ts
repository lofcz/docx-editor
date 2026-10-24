import {
  createNodeIdAllocator,
  findNode,
  parentNodeOf,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { contentControlLevelOf } from '../package/content-control-nodes.ts';
import { validateRemoveContentControl } from './tree-op-validate-controls.ts';
import { contentControlEffect } from './tree-op-content-controls.ts';
import { fromEdit } from './tree-op-nodes.ts';
import { revisionKey } from './tree-op-tracked-adjacency.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';
import {
  invalidRevisionAttribution,
  type TreeDocOp,
  type TreeOpRejection,
  type TreeOpResult,
} from './tree-op-types.ts';

type Removal = Extract<TreeDocOp, { op: 'removeContentControl' }>;
const revisionNames = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
]);

/** Whole inline controls keep their identities until the reviewer decides their native deletion. */
export function trackedControlRemovalRefusal(part: OoxmlPart, op: Removal): TreeOpRejection | null {
  if (op.revision === undefined || invalidRevisionAttribution(op.revision))
    return 'invalid-property-value';
  if (op.keepContent !== false) return 'unsupported';
  const ordinary = validateRemoveContentControl(part, op.controlId);
  if (ordinary) return ordinary;
  const control = findNode(part, op.controlId);
  if (!control || control.kind !== 'contentControl') return 'unknown-content-control';
  const level = contentControlLevelOf(control);
  if (level !== 'inline' && level !== 'empty') return 'unsupported';
  let owner = parentNodeOf(part, control.id);
  while (owner && owner.kind !== 'paragraph') {
    if (owner.kind !== 'hyperlink' && !(owner.kind === 'generic' && owner.localName === 'smartTag'))
      return 'indivisible-content';
    owner = parentNodeOf(part, owner.id);
  }
  if (!owner) return 'not-a-paragraph';
  // An enclosing block proposal or move also owns this paragraph's existence.
  for (
    let ancestor = parentNodeOf(part, owner.id);
    ancestor;
    ancestor = parentNodeOf(part, ancestor.id)
  )
    if (
      ancestor.kind.startsWith('revision') ||
      (ancestor.namespaceUri === WML_NAMESPACE_URI && revisionNames.has(ancestor.localName))
    )
      return 'indivisible-content';
  const pending = [{ node: control as OoxmlNode, depth: 0 }];
  let visited = 0;
  while (pending.length) {
    const { node, depth } = pending.pop()!;
    if (++visited > 100_000 || depth > 256) return 'indivisible-content';
    if (node.kind === 'textValue') continue;
    if (
      node.kind.startsWith('revision') ||
      (node.namespaceUri === WML_NAMESPACE_URI &&
        (revisionNames.has(node.localName) || node.localName.endsWith('Change'))) ||
      node.namespaceUri === 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    )
      return 'indivisible-content';
    for (const child of node.children) pending.push({ node: child, depth: depth + 1 });
  }
  return null;
}

export function applyTrackedControlRemoval(
  part: OoxmlPart,
  op: Removal,
  options?: EditOptions
): TreeOpResult {
  const refusal = trackedControlRemovalRefusal(part, op);
  if (refusal) return { ok: false, reason: refusal };
  const control = findNode(part, op.controlId)!,
    owner = parentNodeOf(part, control.id)!;
  const mint = createNodeIdAllocator(part);
  const markDeleted = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue') return node;
    const children = node.children.map(markDeleted);
    if (node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'instrText')
      return { ...node, kind: 'generic', localName: 'delInstrText', children } as OoxmlNode;
    return node.kind === 'text'
      ? ({ ...node, kind: 'deletedText', localName: 'delText', children } as OoxmlNode)
      : ({ ...node, children } as OoxmlNode);
  };
  const revisionId = options?.trackedRevisionIds?.mint() ?? nextRevisionId(part)();
  options?.trackedRevisionIds?.wrote(
    revisionKey(revisionId, op.revision!.author, op.revision!.date)
  );
  const wrapper = build(
    mint(),
    'revisionDelete',
    'del',
    revisionAttributes(revisionId, op.revision!),
    [markDeleted(control)]
  );
  const result = replaceChildren(
    part,
    owner.id,
    owner.children.map((node) => (node.id === control.id ? wrapper : node)),
    options
  );
  return fromEdit(result, contentControlEffect(owner.id, 'flow-structural'));
}
