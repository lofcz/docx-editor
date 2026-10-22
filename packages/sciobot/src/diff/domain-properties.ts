import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { encodeBlockPayload, isBlockId } from '../blocks/payload-schema';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem, type JsonObject } from '../data/json';
import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
import { domainDiffStrategy, sameContainerIdentity } from './domain-strategies';
import { createEditPlan, EDIT_PLAN_LIMITS, type CanonicalEdit, type EditPlanResult } from './plan';

export interface DomainPropertySource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly blockId: string;
}

/** Canonical property names, after target decoding. Omitted names retain their native values. */
export interface DomainPropertyTarget {
  readonly set?: JsonObject;
  readonly remove?: readonly string[];
}

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const record = (value: unknown): value is JsonObject =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

/**
 * Plan independent payload deltas on a retained container. No child text enters this patch.
 * Executors must verify the revision and old values, then apply all deltas in one transaction.
 */
export function planDomainProperties(
  source: DomainPropertySource,
  target: DomainPropertyTarget
): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing domain property source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const key of ['revision', 'pkg', 'blockId'])
    if (!descriptors[key] || !('value' in descriptors[key]!) || !descriptors[key]!.enumerable)
      return fail('Domain property source fields must be plain values.');
  const input = { revision: source.revision, blockId: source.blockId, target };
  const problem = jsonDataProblem(input, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(input);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Domain property input exceeds its byte limit.');
  if (
    !isBlockId(source.blockId) ||
    !record(target) ||
    Object.keys(target).some((key) => !['set', 'remove'].includes(key)) ||
    (target.set !== undefined && !record(target.set)) ||
    (target.remove !== undefined && !Array.isArray(target.remove))
  )
    return fail('Invalid domain property target.');
  const set = target.set ?? {},
    remove = target.remove ?? [],
    names = [...Object.keys(set), ...remove];
  if (
    names.some((name) => typeof name !== 'string' || name.length === 0) ||
    new Set(names).size !== names.length
  )
    return fail('Domain property names must be unique nonempty strings.');
  const block = buildBlockIndex(source.pkg).byId.get(source.blockId);
  if (!block) return fail('The native domain block does not exist.', 'invalid-document');
  const definition = DOMAIN_BLOCK_TAGS.get(block.payload.type),
    strategy = domainDiffStrategy({ type: block.payload.type });
  if (definition?.kind !== 'container' || strategy?.kind !== 'container')
    return fail(
      'Property planning requires a registered domain container.',
      'unsupported-operation'
    );
  // Query the registered serializer for canonical property ownership. Probe values expose optional fields.
  const probe = Object.fromEntries(
    names.map((name) => [name, name === 'bookId' ? 1 : name === 'trust' ? 'verified' : 'probe'])
  );
  const supported = new Set(
    definition
      .attributes({ ...block.payload, attributes: probe })
      .map((binding) => binding.property)
  );
  if (names.some((name) => !supported.has(name)))
    return fail(
      'The property is not a serialized attribute of this container.',
      'unsupported-operation'
    );
  const attributes = block.payload.attributes,
    removed = new Set(remove);
  const after = Object.fromEntries(
    Object.entries(attributes)
      .filter(([name]) => !removed.has(name))
      .concat(Object.entries(set))
  );
  if (
    !sameContainerIdentity(
      { ...attributes, type: block.payload.type },
      { ...after, type: block.payload.type }
    )
  )
    return fail(
      'The property change requires a structural domain replacement.',
      'unsupported-operation'
    );
  try {
    encodeBlockPayload({ ...block.payload, attributes: after });
  } catch (error) {
    return fail(error instanceof Error ? error.message : 'Invalid domain payload result.');
  }
  const operations: CanonicalEdit[] = names.map((name) => ({
    kind: 'payload',
    partName: block.partName,
    blockId: block.id,
    controlId: block.controlId,
    path: ['attributes', name],
    before: Object.hasOwn(attributes, name)
      ? { present: true, value: attributes[name]! }
      : { present: false },
    after: Object.hasOwn(set, name) ? { present: true, value: set[name]! } : { present: false },
  }));
  return createEditPlan(source.revision, operations);
}
