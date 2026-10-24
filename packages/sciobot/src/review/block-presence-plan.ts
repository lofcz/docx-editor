import {
  bindInlinePresenceNative,
  inlinePresenceExtension,
  readInlinePresenceExtension,
  type BoundInlinePresenceNative,
  type InlinePresenceNative,
} from './inline-presence-native';
import { contentControlLevelOf, findNode } from '@docx-editor.dev/core/store';
import {
  contentControlsIn,
  contentControlPropertiesOf,
  isStoryPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { isBlockId } from '../blocks/payload-schema';
import { readPayloadGroupDecisions, payloadGroupDecisionKey } from './payload-group-decisions';
import { blockIdFromTag, readPayloadStore } from '../blocks/payload-opc-codec';
import { copyBoundedJson } from '../data/json';
import type { NativeEditFailure } from '../contracts/document';
import { jsonDataKey } from '../diff/target-key';
import { readPropertyReviewStore } from './property-record-store';
import { readPresenceOwners } from './block-presence-owner';
import {
  BLOCK_PRESENCE_EXECUTOR,
  blockPresenceAttributionProblem,
  BLOCK_PRESENCE_LIMITS,
  decodeBlockPresenceRecord,
  encodeBlockPresenceRecord,
  type BlockPresenceRecord,
} from './block-presence-schema';

export interface BlockPresenceIntent {
  readonly blockId: string;
  readonly before: { readonly present: boolean };
  readonly after: { readonly present: boolean };
}
export interface BlockPresenceAttribution {
  readonly id: string;
  readonly author: string;
  readonly date: string;
}
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, error, code });
/** Canonical storage contains the present side. No Word wrapper is invented for a block SDT. */
export function createBlockPresenceRecord(
  before: OoxmlPackage,
  canonical: OoxmlPackage,
  intents: readonly BlockPresenceIntent[],
  attribution: BlockPresenceAttribution,
  options: {
    readonly nativeInline?: true;
    readonly bornInline?: readonly InlinePresenceNative[];
  } = {}
): { readonly success: true; readonly record: BlockPresenceRecord } | NativeEditFailure {
  const copied = copyBoundedJson({ intents, attribution, options }, BLOCK_PRESENCE_LIMITS);
  if (!copied.ok) return fail(copied.reason);
  const data = copied.value as unknown as {
    intents: BlockPresenceIntent[];
    attribution: BlockPresenceAttribution;
    options: {
      readonly nativeInline?: true;
      readonly bornInline?: readonly InlinePresenceNative[];
    };
  };
  if (
    !Array.isArray(data.intents) ||
    !data.intents.length ||
    data.intents.length > BLOCK_PRESENCE_LIMITS.maxBlocks
  )
    return fail('Invalid presence intentions.');
  if (
    !data.attribution ||
    typeof data.attribution !== 'object' ||
    Object.keys(data.attribution).some((key) => !['id', 'author', 'date'].includes(key))
  )
    return fail('Invalid presence attribution.');
  const attributionProblem = blockPresenceAttributionProblem(
    data.attribution.id,
    data.attribution.author,
    data.attribution.date
  );
  if (attributionProblem) return fail(attributionProblem);
  for (const intent of data.intents)
    if (
      !intent ||
      typeof intent !== 'object' ||
      Object.keys(intent).some((key) => !['blockId', 'before', 'after'].includes(key)) ||
      !isBlockId(intent.blockId) ||
      !intent.before ||
      !intent.after ||
      typeof intent.before.present !== 'boolean' ||
      typeof intent.after.present !== 'boolean' ||
      intent.before.present === intent.after.present
    )
      return fail('Invalid presence intention.');
  if (
    !data.options ||
    typeof data.options !== 'object' ||
    Array.isArray(data.options) ||
    Object.keys(data.options).some((key) => !['nativeInline', 'bornInline'].includes(key)) ||
    (data.options.bornInline !== undefined &&
      (!Array.isArray(data.options.bornInline) || data.options.nativeInline !== true)) ||
    (data.options.nativeInline !== undefined && data.options.nativeInline !== true)
  )
    return fail('Invalid native presence options.');
  if (
    data.options.bornInline?.some(
      (owner) =>
        !owner ||
        typeof owner !== 'object' ||
        Array.isArray(owner) ||
        !isBlockId(owner.blockId) ||
        typeof owner.partName !== 'string'
    )
  )
    return fail('Invalid paragraph birth ownership options.');
  const current = readPresenceOwners(
    canonical,
    data.intents.map((intent) => intent.blockId),
    data.options.nativeInline === true
  );
  if (!current.success) return current;
  const store = readPropertyReviewStore(before),
    payloads = readPayloadStore(before);
  if (store.status === 'invalid' || payloads.status === 'invalid')
    return fail('The starting application metadata is invalid.', 'invalid-document');
  if (store.status === 'ready' && store.retainedRecords.has(data.attribution.id))
    return fail('The presence proposal ID is already reserved.', 'stale-revision');
  const reserved = new Set<string>();
  for (const part of before.parts.values())
    if (isStoryPart(part))
      for (const { node } of contentControlsIn(part.root)) {
        const id = blockIdFromTag(contentControlPropertiesOf(node).tag ?? '');
        if (id) reserved.add(id);
      }
  if (store.status === 'ready')
    for (const entry of store.retainedRecords.values()) {
      if (entry.decoded.status === 'supported')
        for (const block of entry.decoded.value.blocks) reserved.add(block.blockId);
      else {
        const presence = decodeBlockPresenceRecord(entry.raw);
        if (presence.status !== 'supported')
          return fail(
            'Unknown retained presence ownership needs a supported reader.',
            'unsupported-operation'
          );
        for (const block of presence.value.blocks)
          for (const member of block.members) reserved.add(member.payload.id);
      }
    }
  const decisions = readPayloadGroupDecisions(before);
  if (decisions.status === 'invalid') return fail(decisions.reason, 'invalid-document');
  const targets = new Set(
    [...current.owners.values()].flatMap((owner) =>
      owner.members.map((member) => member.payload.id)
    )
  );
  if (store.status === 'ready')
    for (const entry of store.records.values()) {
      let blocks: readonly { readonly blockId: string; readonly owned: readonly string[] }[];
      if (entry.decoded.status === 'supported')
        blocks = entry.decoded.value.blocks.map((block) => ({
          blockId: block.blockId,
          owned: [block.blockId],
        }));
      else {
        const presence = decodeBlockPresenceRecord(entry.raw);
        if (presence.status !== 'supported')
          return fail('Unknown pending presence ownership.', 'unsupported-operation');
        blocks = presence.value.blocks.map((block) => ({
          blockId: block.blockId,
          owned: block.members.map((member) => member.payload.id),
        }));
      }
      for (const block of blocks) {
        const disposition = decisions.decisions.get(
          payloadGroupDecisionKey(entry.id, block.blockId)
        );
        if (
          block.owned.some((id) => targets.has(id)) &&
          (!disposition || disposition.status === 'conflict')
        )
          return fail(
            'Presence ownership overlaps a pending application proposal.',
            'unsupported-operation'
          );
      }
    }
  const blocks: BlockPresenceRecord['blocks'][number][] = [];
  for (const intent of data.intents) {
    const owner = current.owners.get(intent.blockId)!;
    if (intent.before?.present === false && intent.after?.present === true) {
      for (const member of owner.members)
        if (
          reserved.has(member.payload.id) ||
          (payloads.status === 'ready' && payloads.entries.has(member.payload.id))
        )
          return fail('An inserted block or payload identity already existed.', 'stale-revision');
    } else if (intent.before?.present === true && intent.after?.present === false) {
      const original = readPresenceOwners(
        before,
        [intent.blockId],
        data.options.nativeInline === true
      );
      if (!original.success) return original;
      const source = original.owners.get(intent.blockId)!;
      if (
        source.partName !== owner.partName ||
        source.nativeFingerprint !== owner.nativeFingerprint ||
        jsonDataKey(source.members) !== jsonDataKey(owner.members)
      )
        return fail('A removal proposal changed its retained native owner.', 'stale-revision');
    } else return fail('Presence intentions require opposite explicit states.');
    blocks.push({
      blockId: intent.blockId,
      partName: owner.partName,
      before: intent.before,
      after: intent.after,
      members: owner.members,
      nativeFingerprint: owner.nativeFingerprint,
    });
  }
  const native: BoundInlinePresenceNative[] = [];
  try {
    for (const block of blocks) {
      const owner = current.owners.get(block.blockId)!;
      const control = findNode(canonical.parts.get(owner.partName)!, owner.controlId)!;
      if (contentControlLevelOf(control) !== 'block')
        native.push(
          bindInlinePresenceNative(
            canonical,
            block.blockId,
            block.partName,
            owner.controlId,
            block.after.present ? 'ins' : 'del',
            data.options.bornInline?.find((owner) => owner.blockId === block.blockId)
          )
        );
    }
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error), 'unsupported-operation');
  }
  if (
    data.options.bornInline?.some(
      (owner, at, entries) =>
        !native.some((bound) => bound.blockId === owner.blockId && bound.gap) ||
        entries.findIndex((candidate) => candidate.blockId === owner.blockId) !== at
    )
  )
    return fail('Paragraph birth ownership contains unmatched or duplicate roots.');
  const record: BlockPresenceRecord = {
    schemaVersion: native.some((owner) => owner.siblings)
      ? 6
      : native.some((owner) => owner.cell)
        ? 5
        : native.some((owner) => owner.gap)
          ? 4
          : native.length
            ? 3
            : 2,
    executor: BLOCK_PRESENCE_EXECUTOR,
    ...data.attribution,
    predecessors: store.status === 'ready' ? [...store.records.keys()].sort() : [],
    blocks,
    ...(native.length ? { extension: inlinePresenceExtension(native) } : {}),
  };
  try {
    const decoded = decodeBlockPresenceRecord(encodeBlockPresenceRecord(record));
    return decoded.status === 'supported'
      ? { success: true, record: decoded.value }
      : fail('Invalid presence record.');
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
}
/** Re-read after each revision. Seen native, payload, label, and topology edits refuse before writes. */
export function bindBlockPresenceRecord(
  pkg: OoxmlPackage,
  input: BlockPresenceRecord
):
  | {
      readonly success: true;
      readonly record: BlockPresenceRecord;
      readonly controlIds: ReadonlyMap<string, string>;
      readonly nativeSites: ReadonlyMap<string, BoundInlinePresenceNative>;
    }
  | NativeEditFailure {
  let record: BlockPresenceRecord;
  try {
    const decoded = decodeBlockPresenceRecord(encodeBlockPresenceRecord(input));
    if (decoded.status !== 'supported')
      return fail('Unsupported presence record.', 'unsupported-operation');
    record = decoded.value;
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
  let inline: ReturnType<typeof readInlinePresenceExtension>;
  try {
    inline = readInlinePresenceExtension(record);
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
  const current = readPresenceOwners(
    pkg,
    record.blocks.map((block) => block.blockId),
    inline !== null
  );
  if (!current.success) return current;
  const controlIds = new Map<string, string>();
  const nativeSites = new Map<string, BoundInlinePresenceNative>();
  for (const block of record.blocks) {
    const owner = current.owners.get(block.blockId)!;
    if (
      owner.partName !== block.partName ||
      owner.nativeFingerprint !== block.nativeFingerprint ||
      jsonDataKey(owner.members) !== jsonDataKey(block.members)
    )
      return fail('A pending presence owner changed.', 'stale-revision');
    const control = findNode(pkg.parts.get(owner.partName)!, owner.controlId)!;
    if (contentControlLevelOf(control) !== 'block') {
      const expected = inline?.get(block.blockId);
      if (!expected) return fail('The inline presence native proof is missing.');
      try {
        nativeSites.set(
          block.blockId,
          bindInlinePresenceNative(
            pkg,
            block.blockId,
            block.partName,
            owner.controlId,
            block.after.present ? 'ins' : 'del',
            expected
          )
        );
      } catch (error) {
        return fail(error instanceof Error ? error.message : String(error), 'stale-revision');
      }
    } else if (inline?.has(block.blockId))
      return fail('Inline native ownership cannot claim a block root.');
    controlIds.set(block.blockId, owner.controlId);
  }
  return { success: true, record, controlIds, nativeSites };
}
