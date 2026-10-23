import type { InlineImageRuntime } from '../native/inline-image-runtime';
import { retainedAtomLinks } from './direct-atom-links';
import {
  inlineCodeRunChanges,
  inlineCodeCharacterStyles,
  hyperlinkCharacterStyle,
} from '../native/inline-code';
import {
  applyTreeOp,
  AUTHORABLE_RUN_PROPERTIES,
  authoredProperties,
  contentControlsIn,
  contentControlPropertiesOf,
  paragraphTextOf,
  propertyContainer,
  runsCovering,
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
  edit: Extract<CanonicalEdit, { kind: 'inline-atoms' }>,
  images?: InlineImageRuntime,
  allocatedCodeStyles: readonly string[] = []
): { readonly ops: readonly TreeDocOp[]; readonly complete: (pkg: OoxmlPackage) => OoxmlPackage } {
  if (
    paragraphTextOf(part, edit.paragraphId)?.slice(
      edit.offset,
      edit.offset + edit.before.length
    ) !== edit.before
  )
    refuseWrite('The retained atom insertion text no longer matches.', 'stale-revision');
  const controls = edit.atoms.flatMap((atom) => {
    if (atom.kind === 'drawing' && images) return [];
    if (atom.kind !== 'inline-control' || !inlineAtomXml(atom))
      refuseWrite(
        'Inserted atoms need supported native payload and resource reconciliation.',
        'unsupported-operation'
      );
    return [atom as Extract<typeof atom, { kind: 'inline-control' }>];
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
  const links = retainedAtomLinks(pkg, part, edit);
  const codeStyles = new Set([...inlineCodeCharacterStyles(pkg), ...allocatedCodeStyles]);
  let preview = part;
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      refuseWrite(`Retained atom operation refused: ${result.reason}.`, 'apply-failed');
    preview = (result as { part: OoxmlPart }).part;
    ops.push(op);
  };
  for (const op of links.unlink) apply(op);
  for (const atom of edit.atoms) {
    if (atom.kind === 'drawing') {
      apply(images!.reservation(atom, edit.paragraphId));
      continue;
    }
    apply({
      op: 'insertInlineContentControl',
      paragraphId: edit.paragraphId,
      offset: atom.offset,
      text: atom.fallback,
      tag: blockTag(atom.payload.id),
      alias: atom.label,
    });
    if (atom.fallback.length) {
      const run = runsCovering(
        preview,
        edit.paragraphId,
        atom.offset,
        atom.offset + atom.fallback.length
      )[0];
      const inherited = authoredProperties(
        propertyContainer(run, 'runProperties', 'rPr'),
        AUTHORABLE_RUN_PROPERTIES
      ).filter((property) => !['b', 'i', 'strike'].includes(property.localName));
      apply({
        op: 'setRunProperties',
        paragraphId: edit.paragraphId,
        start: atom.offset,
        end: atom.offset + atom.fallback.length,
        properties: [
          ...inherited,
          ...atom.marks
            .filter((mark) => mark !== 'code')
            .map((mark) => ({
              localName: mark === 'bold' ? 'b' : mark === 'italic' ? 'i' : 'strike',
              attributes: {},
            })),
        ],
      });
      const code = inlineCodeRunChanges(
        preview,
        {
          paragraphId: edit.paragraphId,
          start: atom.offset,
          end: atom.offset + atom.fallback.length,
          code: atom.marks.includes('code'),
        },
        { codeStyles: [...codeStyles], hyperlinkStyle: hyperlinkCharacterStyle(pkg) }
      );
      if (!code.success) refuseWrite(code.error, code.code);
      for (const change of code.changes) {
        if (change.after !== null && atom.marks.includes('code')) codeStyles.add(change.after);
        apply({
          op: 'setRunProperties',
          paragraphId: edit.paragraphId,
          start: change.start,
          end: change.end,
          properties: change.properties,
          characterStyle: change.after,
          ...(change.targetRunIds ? { targetRunIds: change.targetRunIds } : {}),
        });
      }
    }
  }
  return {
    ops,
    complete(working) {
      let next = links.complete(working);
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
