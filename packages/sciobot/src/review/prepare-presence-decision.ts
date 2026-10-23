import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import {
  replaceNode,
  withPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { bindBlockPresenceRecord } from './block-presence-plan';
import { decodeBlockPresenceRecord } from './block-presence-schema';
import { readPropertyReviewStore } from './property-record-store';
import { payloadGroupDecisionKey, readPayloadGroupDecisions } from './payload-group-decisions';

/** Prepare the exact native owner and nested payload removals; receipts use the normal group transaction. */
export function preparePresenceDecision(
  pkg: OoxmlPackage,
  proposalId: string,
  blockId: string,
  decision: 'accept' | 'reject'
):
  | {
      readonly success: true;
      readonly partName: string;
      readonly ops: readonly TreeDocOp[];
      readonly apply: (pkg: OoxmlPackage) => OoxmlPackage;
    }
  | NativeEditFailure {
  const fail = (
    error: string,
    code: NativeEditFailure['code'] = 'invalid-document'
  ): NativeEditFailure => ({ success: false, error, code });
  const store = readPropertyReviewStore(pkg),
    entry = store.status === 'ready' ? store.records.get(proposalId) : undefined;
  const decoded = entry && decodeBlockPresenceRecord(entry.raw);
  if (decoded?.status !== 'supported')
    return fail('The presence proposal is missing or unsupported.', 'unsupported-operation');
  const block = decoded.value.blocks.find((block) => block.blockId === blockId);
  if (!block) return fail('The selected presence root is missing.');
  const dispositions = readPayloadGroupDecisions(pkg);
  if (dispositions.status === 'invalid') return fail(dispositions.reason);
  if (dispositions.decisions.has(payloadGroupDecisionKey(proposalId, blockId)))
    return fail('The presence group is already resolved or conflicting.', 'stale-revision');
  const bound = bindBlockPresenceRecord(pkg, { ...decoded.value, blocks: [block] });
  if (!bound.success) return bound;
  const present = (decision === 'accept' ? block.after : block.before).present;
  const payloads = readPayloadStore(pkg);
  if (payloads.status !== 'ready') return fail('The presence payload store is missing.');
  return {
    success: true,
    partName: block.partName,
    ops: present
      ? []
      : [
          {
            op: 'removeContentControl',
            controlId: bound.controlIds.get(blockId)!,
            keepContent: false,
          },
        ],
    apply(working) {
      if (present) return working;
      const index = buildBlockIndex(working);
      if (block.members.some((member) => index.byId.has(member.payload.id)))
        throw Error('A removed presence control still owns payloads.');
      let next = working;
      for (const member of block.members) {
        const part = next.parts.get(payloads.part.partName)!;
        const node = part.root.children.find(
          (child) =>
            child.kind !== 'textValue' &&
            child.attributes.some(
              (attribute) =>
                attribute.namespaceUri === '' &&
                attribute.localName === 'id' &&
                attribute.value === 'scio_' + member.payload.id
            )
        );
        if (!node) throw Error('An owned presence payload is missing.');
        // Retire the exact record by identity. Shared custom XML can sort adopted
        // siblings differently from storage. Retain an empty shell for peer-preserving
        // history, as the payload codec already permits for native undo.
        if (node.kind === 'textValue') throw Error('Invalid presence payload node.');
        const removed = replaceNode(part, node.id, { ...node, attributes: [], children: [] });
        if (!removed.ok) throw Error('Presence payload removal failed.');
        next = withPart(next, removed.part);
      }
      const read = readPayloadStore(next);
      if (
        read.status !== 'ready' ||
        block.members.some((member) => read.entries.has(member.payload.id))
      )
        throw Error('Presence payload removal failed.');
      return next;
    },
  };
}
