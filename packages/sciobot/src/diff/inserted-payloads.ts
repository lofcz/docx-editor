import {
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { blockIdFromTag, readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { NativeEditFailure } from '../contracts/document';
import type { InsertedCellPayloadBinding } from './inserted-cell-fragments';

/** Validate new ownership and storage on a detached package, without a live transaction. */
export function insertedPayloadProblem(
  pkg: OoxmlPackage,
  payloads: readonly InsertedCellPayloadBinding[]
): NativeEditFailure | null {
  const fail = (
    error: string,
    code: NativeEditFailure['code'] = 'invalid-input'
  ): NativeEditFailure => ({ success: false, code, error });
  const store = readPayloadStore(pkg);
  if (store.status === 'invalid')
    return fail(`Invalid payload store: ${store.reason}.`, 'invalid-document');
  const proposedIds = new Set(payloads.map((binding) => binding.payload.id));
  if (proposedIds.size) {
    for (const part of pkg.parts.values()) {
      if (!isStoryPart(part)) continue;
      for (const { node } of contentControlsIn(part.root)) {
        const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
        if (id && proposedIds.has(id)) return fail('Inserted application block IDs must be new.');
      }
    }
  }
  let candidate = pkg;
  for (const binding of payloads) {
    if (store.status === 'ready' && store.entries.has(binding.payload.id))
      return fail('Inserted application block IDs must be new.');
    const prepared = withPayloadRecord(candidate, binding.payload, binding.label);
    if (!prepared.ok) return fail(`Inserted payload storage refused: ${prepared.reason}.`);
    candidate = prepared.pkg;
  }
  return null;
}
