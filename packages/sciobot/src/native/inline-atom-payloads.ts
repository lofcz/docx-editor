import { withInlineControlId } from './inline-image-control-id';
import {
  applyTreeOp,
  contentControlsIn,
  contentControlPropertiesOf,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { InsertedInlineAtom } from '../diff/inserted-atoms';

/** Finalize each newly allocated control once, after its native paragraph identity exists. */
export function completeInlineAtomPayloads(
  input: OoxmlPackage,
  partName: string,
  atoms: readonly InsertedInlineAtom[]
): OoxmlPackage {
  let pkg = input;
  for (const atom of atoms) {
    if (atom.kind !== 'inline-control') continue;
    const stored = withPayloadRecord(pkg, atom.payload, atom.label);
    if (!stored.ok) throw Error(`Inline payload storage refused: ${stored.reason}.`);
    pkg = stored.pkg;
    const part = pkg.parts.get(partName)!;
    const controls = contentControlsIn(part.root).filter(
      ({ node }) => contentControlPropertiesOf(node).tag === blockTag(atom.payload.id)
    );
    if (controls.length !== 1) throw Error('The inserted atom needs one native control.');
    pkg = withInlineControlId(pkg, partName, controls[0]!.node.id);
    const locked = applyTreeOp(pkg.parts.get(partName)!, {
      op: 'setContentControlProperties',
      controlId: controls[0]!.node.id,
      lock: 'sdtContentLocked',
    });
    if (!locked.ok) throw Error(`Inline atom locking refused: ${locked.reason}.`);
    pkg = withPart(pkg, locked.part);
  }
  return pkg;
}
