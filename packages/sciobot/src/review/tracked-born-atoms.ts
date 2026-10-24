import type { InlineImageRuntime } from '../native/inline-image-runtime';
import {
  contentControlsIn,
  contentControlPropertiesOf,
  parentNodeOf,
} from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { refuseWrite } from '../diff/fragment-seeds';
import type { InsertedInlineAtom } from '../diff/inserted-atoms';
import type { TrackedStageWriter } from './tracked-stage-writer';

/** Finalize controls inside paragraphs born in this stage, after native resources are complete. */
export function stageBornAtoms(
  writer: TrackedStageWriter,
  partName: string,
  paragraphIds: readonly string[],
  atoms: readonly InsertedInlineAtom[],
  images?: InlineImageRuntime
): void {
  for (const atom of atoms) {
    if (atom.kind === 'drawing') {
      if (!images)
        refuseWrite('Tracked image births need resource ownership.', 'unsupported-operation');
      continue; // The same admitted runtime finalizes the owned reservation before presence capture.
    }
    const part = writer.currentPackage().parts.get(partName)!;
    const controls = contentControlsIn(part.root).filter(
      ({ node }) => contentControlPropertiesOf(node).tag === blockTag(atom.payload.id)
    );
    if (controls.length !== 1)
      refuseWrite('A born atom needs one new native control.', 'ambiguous-target');
    const control = controls[0]!.node;
    let owner = parentNodeOf(part, control.id);
    while (owner && owner.kind !== 'paragraph') owner = parentNodeOf(part, owner.id);
    if (!owner || !paragraphIds.includes(owner.id))
      refuseWrite(
        'The atom does not belong to a newly allocated paragraph.',
        'unsupported-operation'
      );
    writer.edit((pkg) => {
      const stored = withPayloadRecord(pkg, atom.payload, atom.label);
      if (!stored.ok) refuseWrite('The born atom payload could not be stored.', 'apply-failed');
      return stored.pkg;
    });
    writer.apply(partName, {
      op: 'setContentControlProperties',
      controlId: control.id,
      lock: 'sdtContentLocked',
    });
  }
}
