import type { MdxComponentRegistry } from '../registry/contract';
import { activeMdxRegistry } from '../registry/active';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { encodeBlockPayload, isBlockId, type BlockPayload } from '../blocks/payload-schema';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import { MDX_BLOCK, MDX_INLINE, serializeOpaqueMdx } from '../projection/mdx';
import { domainDiffStrategy } from './domain-strategies';
import { createEditPlan, EDIT_PLAN_LIMITS, type EditPlanResult } from './plan';
import { jsonDataKey, targetAstKey } from './target-key';

export interface AtomicBlockSource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly blockId: string;
}

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const opaque = (type: string) => type === MDX_BLOCK || type === MDX_INLINE;
const atomic = (type: string, registry: MdxComponentRegistry) =>
  opaque(type) || domainDiffStrategy({ type }, registry)?.kind === 'atomic';

function payloadKey(payload: BlockPayload): string {
  return jsonDataKey(
    opaque(payload.type) && payload.mdast
      ? { ...payload, mdast: targetAstKey(payload.mdast) }
      : payload
  );
}

/**
 * Replace one complete atomic payload through an empty payload path. No descendants are paired.
 * Supply the complete decoded target, including reconciled assets and opaque metadata.
 * Equal targets keep the original stored payload and parser locations without any edit.
 * Executors must treat a root payload change as one block replacement and update its native fallback.
 */
export function planAtomicBlock(
  source: AtomicBlockSource,
  target: BlockPayload,
  registry: MdxComponentRegistry = activeMdxRegistry()
): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing atomic block source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'pkg', 'blockId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Atomic block source fields must be plain values.');
  const problem = jsonDataProblem(
    { revision: source.revision, blockId: source.blockId },
    EDIT_PLAN_LIMITS
  );
  if (problem || !isBlockId(source.blockId)) return fail(problem ?? 'Invalid atomic block ID.');
  let after: BlockPayload;
  try {
    after = JSON.parse(encodeBlockPayload(target)) as BlockPayload;
    if (opaque(after.type)) serializeOpaqueMdx(after);
  } catch (error) {
    return fail(error instanceof Error ? error.message : 'Invalid atomic target payload.');
  }
  const block = buildBlockIndex(source.pkg).byId.get(source.blockId);
  if (!block) return fail('The native atomic block does not exist.', 'invalid-document');
  const before = block.payload;
  if (!atomic(before.type, registry))
    return fail(
      'The native block requires container or structural planning.',
      'unsupported-operation'
    );
  if (after.id !== before.id || after.type !== before.type)
    return fail('Atomic payload replacement must retain its block ID and payload type.');
  try {
    if (opaque(before.type)) serializeOpaqueMdx(before);
  } catch (error) {
    return fail(
      error instanceof Error ? error.message : 'Invalid stored atomic payload.',
      'invalid-document'
    );
  }
  if (payloadKey(before) === payloadKey(after)) return createEditPlan(source.revision, []);
  return createEditPlan(source.revision, [
    {
      kind: 'payload',
      partName: block.partName,
      blockId: block.id,
      controlId: block.controlId,
      path: [],
      before: { present: true, value: before },
      after: { present: true, value: after },
    },
  ]);
}
