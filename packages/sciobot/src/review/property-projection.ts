import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import {
  DEFAULT_PAYLOAD_LIMITS,
  decodeBlockPayload,
  type BlockPayload,
} from '../blocks/payload-schema';
import type { NativeEditFailure, RevisionToken, DocumentView } from '../contracts/document';
import { copyBoundedJson, type JsonObject } from '../data/json';
import {
  createEditPlan,
  decodeEditPlan,
  type CanonicalEdit,
  type CanonicalEditPlan,
  type PlannedValue,
} from '../diff/plan';
import { planDomainProperties, type DomainPropertyIntent } from '../diff/domain-properties';
import { planAtomicBlock } from '../diff/atomic-blocks';
import { domainDiffStrategy } from '../diff/domain-strategies';
import { jsonDataKey } from '../diff/target-key';
import { MDX_BLOCK, MDX_INLINE, serializeOpaqueMdx } from '../projection/mdx';
import { ATOMIC_PAYLOAD_CHANGE } from './property-record-schema';

/** Detached property/atomic data. R12/R13 and R14-R16 own persistence and live decisions. */
export interface PendingPropertyProposal {
  readonly id: string;
  readonly plan: CanonicalEditPlan;
  readonly intent?: DomainPropertyIntent;
}
export type PropertyProposalResult =
  | { readonly success: true; readonly proposal: PendingPropertyProposal }
  | NativeEditFailure;
export type PropertyProjectionResult =
  | {
      readonly success: true;
      /** A derived read candidate. Never install it as a replacement live document. */
      readonly pkg: OoxmlPackage;
    }
  | NativeEditFailure;

const MAX_PROPOSALS = 1024;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
type PayloadEdit = Extract<CanonicalEdit, { kind: 'payload' }>;
const propertyEdit = (edit: CanonicalEdit): edit is PayloadEdit =>
  edit.kind === 'payload' &&
  edit.path.length === 2 &&
  edit.path[0] === 'attributes' &&
  typeof edit.path[1] === 'string' &&
  edit.path[1].length > 0;
function atomicEdit(edit: CanonicalEdit): edit is PayloadEdit {
  if (edit.kind !== 'payload' || edit.path.length || !edit.before.present || !edit.after.present)
    return false;
  const payloads = [edit.before.value, edit.after.value].map((value) =>
    decodeBlockPayload(JSON.stringify(value))
  );
  if (payloads.some((value) => value.status !== 'supported')) return false;
  const [before, after] = payloads.map((value) =>
    value.status === 'supported' ? value.value : null
  ) as [BlockPayload, BlockPayload];
  if (before.id !== edit.blockId || after.id !== before.id || after.type !== before.type)
    return false;
  if ([MDX_BLOCK, MDX_INLINE].includes(before.type)) {
    try {
      serializeOpaqueMdx(before);
      serializeOpaqueMdx(after);
      return true;
    } catch {
      return false;
    }
  }
  return domainDiffStrategy({ type: before.type })?.kind === 'atomic';
}
const sameState = (a: PlannedValue, b: PlannedValue) =>
  a.present === b.present &&
  (!a.present || (b.present && jsonDataKey(a.value) === jsonDataKey(b.value)));
const stateAt = (attributes: JsonObject, key: string): PlannedValue =>
  Object.hasOwn(attributes, key) ? { present: true, value: attributes[key]! } : { present: false };
const withStates = (
  attributes: JsonObject,
  states: ReadonlyMap<string, PlannedValue>
): JsonObject =>
  Object.fromEntries([
    ...Object.entries(attributes).filter(([key]) => !states.has(key)),
    ...[...states].flatMap(([key, state]) => (state.present ? [[key, state.value]] : [])),
  ]);

/** Validate and freeze D13 attribute deltas or D14 complete atomic payload pairs. */
export function createPendingPropertyProposal(
  id: string,
  plan: CanonicalEditPlan,
  intent?: DomainPropertyIntent
): PropertyProposalResult {
  const copied = copyBoundedJson(
    { id, plan, ...(intent === undefined ? {} : { intent }) },
    DEFAULT_PAYLOAD_LIMITS
  );
  if (!copied.ok) return fail(copied.reason);
  const data = copied.value as unknown as PendingPropertyProposal;
  if (data.intent !== undefined && data.intent !== 'field-edit')
    return fail('Invalid property edit intent.');
  if (typeof data.id !== 'string' || !data.id.trim() || data.id.length > 128)
    return fail('Property proposal IDs require nonempty strings of at most 128 characters.');
  const checked = decodeEditPlan(JSON.stringify(data.plan));
  if (!checked.success) return checked;
  if (
    !checked.plan.operations.length ||
    !checked.plan.operations.every((edit) => propertyEdit(edit) || atomicEdit(edit))
  )
    return fail(
      'Application proposals require nonempty property or complete atomic payload plans.',
      'unsupported-operation'
    );
  return {
    success: true,
    proposal: Object.freeze({
      id: data.id,
      plan: checked.plan,
      ...(data.intent ? { intent: data.intent } : {}),
    }),
  };
}

interface BlockChain {
  readonly controlId: string;
  readonly partName: string;
  readonly keys: Map<string, PlannedValue[]>;
  readonly steps: { edits: PayloadEdit[]; intent?: DomainPropertyIntent }[];
  readonly atomic: boolean;
}

/**
 * Overlay old/proposed payload data on an immutable package for a read. Native story trees and labels
 * remain unchanged. Ordered same-property proposals must form an exact presence/value chain.
 */
export function projectPendingProperties(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  pending: readonly PendingPropertyProposal[],
  kind: DocumentView
): PropertyProjectionResult {
  const copied = copyBoundedJson({ revision, pending, kind }, DEFAULT_PAYLOAD_LIMITS);
  if (!copied.ok) return fail(copied.reason);
  const data = copied.value as unknown as {
    revision: RevisionToken;
    pending: PendingPropertyProposal[];
    kind: DocumentView;
  };
  const token = createEditPlan(data.revision, []);
  if (!token.success) return token;
  if (!['committed', 'working'].includes(data.kind) || !Array.isArray(data.pending))
    return fail('Invalid property view or pending proposal list.');
  if (data.pending.length > MAX_PROPOSALS)
    return fail('Property proposals exceed their count limit.');
  if (!data.pending.length) return { success: true, pkg };
  const ids = new Set<string>(),
    chains = new Map<string, BlockChain>();
  for (const raw of data.pending) {
    if (
      !raw ||
      typeof raw !== 'object' ||
      Object.keys(raw).some((key) => !['id', 'plan', 'intent'].includes(key))
    )
      return fail('Invalid property proposal record.');
    const checked = createPendingPropertyProposal(raw.id, raw.plan, raw.intent);
    if (!checked.success) return checked;
    const proposal = checked.proposal;
    if (ids.has(proposal.id)) return fail('Duplicate property proposal IDs.', 'ambiguous-target');
    ids.add(proposal.id);
    if (proposal.plan.revision.documentId !== data.revision.documentId)
      return fail('A property proposal belongs to another document.', 'stale-revision');
    const touched = new Set<string>();
    const step = new Map<string, PayloadEdit[]>();
    for (const edit of proposal.plan.operations as readonly PayloadEdit[]) {
      const root = edit.path.length === 0,
        key = root ? ATOMIC_PAYLOAD_CHANGE : (edit.path[1] as string),
        address = JSON.stringify([edit.blockId, key]);
      if (touched.has(address))
        return fail(
          'One property proposal addresses a property more than once.',
          'ambiguous-target'
        );
      touched.add(address);
      let chain = chains.get(edit.blockId);
      if (!chain) {
        chain = {
          controlId: edit.controlId,
          partName: edit.partName,
          keys: new Map(),
          steps: [],
          atomic: root,
        };
        chains.set(edit.blockId, chain);
      }
      if (chain.controlId !== edit.controlId || chain.partName !== edit.partName)
        return fail('Property proposals disagree on the native block owner.', 'ambiguous-target');
      if (chain.atomic !== root)
        return fail(
          'Atomic and attribute proposals need coordinated group ownership.',
          'ambiguous-target'
        );
      const states = chain.keys.get(key);
      if (states) {
        if (!sameState(states.at(-1)!, edit.before))
          return fail(
            'Pending property values do not form a continuous chain.',
            'ambiguous-target'
          );
        states.push(edit.after);
      } else chain.keys.set(key, [edit.before, edit.after]);
      const blockEdits = step.get(edit.blockId) ?? [];
      blockEdits.push(edit);
      step.set(edit.blockId, blockEdits);
    }
    for (const [blockId, edits] of step)
      chains
        .get(blockId)!
        .steps.push({ edits, ...(proposal.intent ? { intent: proposal.intent } : {}) });
  }
  const index = buildBlockIndex(pkg),
    store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return fail('Pending properties need a valid native payload store.', 'invalid-document');
  let projected = pkg;
  for (const [blockId, chain] of chains) {
    const block = index.byId.get(blockId),
      entry = store.entries.get(blockId);
    if (
      !block ||
      !entry ||
      block.controlId !== chain.controlId ||
      block.partName !== chain.partName
    )
      return fail('The pending property block has no exact native owner.', 'invalid-document');
    if (chain.atomic) {
      const states = chain.keys.get(ATOMIC_PAYLOAD_CHANGE)!;
      if (
        !states.some(
          (state) => state.present && jsonDataKey(state.value) === jsonDataKey(block.payload)
        )
      )
        return fail(
          'A pending atomic payload differs from its retained history.',
          'stale-revision'
        );
      let working = projected;
      for (let position = 0; position < states.length; position++) {
        const state = states[position]!;
        if (!state.present) return fail('An atomic replacement needs both payloads.');
        const payload = state.value as BlockPayload;
        if (position) {
          const checked = planAtomicBlock(
            { pkg: working, revision: data.revision, blockId },
            payload
          );
          if (!checked.success) return checked;
        }
        const written = withPayloadRecord(working, payload, entry.label);
        if (!written.ok) return fail(written.reason, 'invalid-document');
        working = written.pkg;
        if (data.kind === 'committed' && position === 0) projected = working;
      }
      if (data.kind === 'working') projected = working;
      continue;
    }
    const before = new Map<string, PlannedValue>(),
      after = new Map<string, PlannedValue>();
    for (const [key, states] of chain.keys) {
      const current = stateAt(block.payload.attributes, key);
      if (!states.some((state) => sameState(state, current)))
        return fail('A pending property differs from its native value history.', 'stale-revision');
      before.set(key, states[0]!);
      after.set(key, states.at(-1)!);
    }
    const baseline = withPayloadRecord(
      projected,
      { ...block.payload, attributes: withStates(block.payload.attributes, before) },
      entry.label
    );
    if (!baseline.ok) return fail(baseline.reason, 'invalid-document');
    // Validate complete proposal steps. Some native container rules require related updates together.
    let working = baseline.pkg,
      attributes = withStates(block.payload.attributes, before);
    for (const step of chain.steps) {
      const targets = new Map(step.edits.map((edit) => [edit.path[1] as string, edit.after])),
        checked = planDomainProperties(
          { pkg: working, revision: data.revision, blockId },
          {
            set: Object.fromEntries(
              [...targets].flatMap(([key, state]) => (state.present ? [[key, state.value]] : []))
            ),
            remove: [...targets].filter(([, state]) => !state.present).map(([key]) => key),
          },
          step.intent
        );
      if (!checked.success) return checked;
      attributes = withStates(attributes, targets);
      const next = withPayloadRecord(working, { ...block.payload, attributes }, entry.label);
      if (!next.ok) return fail(next.reason, 'invalid-document');
      working = next.pkg;
    }
    const selected = data.kind === 'committed' ? before : after;
    if (
      [...selected].every(([key, state]) =>
        sameState(stateAt(block.payload.attributes, key), state)
      )
    )
      continue;
    projected = data.kind === 'committed' ? baseline.pkg : working;
  }
  return { success: true, pkg: projected };
}
