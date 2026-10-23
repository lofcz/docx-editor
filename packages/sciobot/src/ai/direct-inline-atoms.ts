import {
  applyTreeOp,
  contentControlsIn,
  contentControlPropertiesOf,
  paragraphTextOf,
  withPart,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { insertedPayloadProblem } from '../diff/inserted-payloads';
import { inlineAtomXml } from '../diff/inline-atom-xml';
import type { CanonicalEdit } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';

/** Native operations address the retained paragraph. Fresh control IDs resolve only at completion. */
export function lowerRetainedAtoms(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  edit: Extract<CanonicalEdit, { kind: 'inline-atoms' }>
): { readonly ops: readonly TreeDocOp[]; readonly complete: (pkg: OoxmlPackage) => OoxmlPackage } {
  if (
    paragraphTextOf(part, edit.paragraphId)?.slice(
      edit.offset,
      edit.offset + edit.before.length
    ) !== edit.before
  )
    refuseWrite('The retained atom insertion text no longer matches.', 'stale-revision');
  const controls = edit.atoms.map((atom) => {
    if (atom.kind !== 'inline-control' || !inlineAtomXml(atom))
      refuseWrite(
        'Inserted atoms need supported native payload and resource reconciliation.',
        'unsupported-operation'
      );
    return atom as Extract<typeof atom, { kind: 'inline-control' }>;
  });
  const problem = insertedPayloadProblem(
    withPart(pkg, part),
    controls.map((op) => ({
      path: op.path,
      payload: op.payload,
      label: op.label,
    }))
  );
  if (problem) refuseWrite(problem.error, problem.code);
  const ops: TreeDocOp[] = [];
  let preview = part;
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      refuseWrite(`Retained atom operation refused: ${result.reason}.`, 'apply-failed');
    preview = (result as { part: OoxmlPart }).part;
    ops.push(op);
  };
  for (const atom of controls) {
    apply({
      op: 'insertInlineContentControl',
      paragraphId: edit.paragraphId,
      offset: atom.offset,
      text: atom.fallback,
      tag: blockTag(atom.payload.id),
      alias: atom.label,
    });
    if (atom.marks.length && atom.fallback.length)
      apply({
        op: 'setRunProperties',
        paragraphId: edit.paragraphId,
        start: atom.offset,
        end: atom.offset + atom.fallback.length,
        properties: atom.marks.map((mark) => ({
          localName: mark === 'bold' ? 'b' : mark === 'italic' ? 'i' : 'strike',
          attributes: {},
        })),
      });
  }
  return {
    ops,
    complete(working) {
      let next = working;
      for (const atom of controls) {
        const stored = withPayloadRecord(next, atom.payload, atom.label);
        if (!stored.ok)
          refuseWrite(`Inline payload storage refused: ${stored.reason}.`, 'apply-failed');
        next = stored.pkg;
        const part = next.parts.get(edit.partName)!;
        const control = contentControlsIn(part.root).find(
          ({ node }) => contentControlPropertiesOf(node).tag === blockTag(atom.payload.id)
        )?.node;
        if (!control) refuseWrite('The retained inline control was not allocated.', 'apply-failed');
        const locked = applyTreeOp(part, {
          op: 'setContentControlProperties',
          controlId: control!.id,
          lock: 'sdtContentLocked',
        });
        if (!locked.ok)
          refuseWrite(`Inline atom locking refused: ${locked.reason}.`, 'apply-failed');
        next = withPart(next, (locked as { part: OoxmlPart }).part);
      }
      return next;
    },
  };
}
