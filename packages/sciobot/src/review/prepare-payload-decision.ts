import {
  applyTreeOp,
  collectRevisionSites,
  type TreeDocOp,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload } from '../blocks/payload-schema';
import type { JsonObject } from '../data/json';
import { jsonDataKey } from '../diff/target-key';
import { type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { readPropertyReviewStore } from './property-record-store';
import { ATOMIC_PAYLOAD_CHANGE, ATOMIC_REVIEW_EXECUTOR } from './property-record-schema';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { projectPendingProperties } from './property-projection';
import {
  nativePropertyAddress,
  paragraphOwner,
  propertyAddressKey,
  propertyControlRefusal,
  propertyFallbackParagraph,
  atomicFallbackParagraph,
} from './property-fallback';

import { readPayloadGroupDecisions, payloadGroupDecisionKey } from './payload-group-decisions';

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Pure preflight shared by one-group and legacy whole-proposal decisions. */
export function preparePayloadProposalDecision(
  pkg: OoxmlPackage,
  current: RevisionToken,
  proposalId: string,
  decision: 'accept' | 'reject',
  executor: 'properties' | 'atomic',
  blockIds?: ReadonlySet<string>
):
  | {
      readonly success: true;
      readonly ops: readonly TreeDocOp[];
      readonly partName: string;
      readonly apply: (pkg: OoxmlPackage) => OoxmlPackage;
    }
  | NativeEditFailure {
  const atomic = executor === 'atomic',
    data = { proposalId, decision };
  const part = pkg.parts.get(pkg.mainDocumentPart)!,
    store = readPropertyReviewStore(pkg),
    payloads = readPayloadStore(pkg);
  if (store.status !== 'ready' || payloads.status !== 'ready')
    return fail('The native property store is not ready.', 'invalid-document');
  const entry = store.records.get(data.proposalId);
  if (!entry || entry.decoded.status !== 'supported')
    return fail(
      'The pending property proposal is missing or unsupported.',
      'unsupported-operation'
    );
  const dispositions = readPayloadGroupDecisions(pkg);
  if (dispositions.status === 'invalid') return fail(dispositions.reason, 'invalid-document');
  const sourceRecord = entry.decoded.value;
  if (
    sourceRecord.blocks.some(
      (block) =>
        (!blockIds || blockIds.has(block.blockId)) &&
        dispositions.decisions.get(payloadGroupDecisionKey(sourceRecord.id, block.blockId))
          ?.status === 'conflict'
    )
  )
    return fail('Conflicting property group decisions require reconciliation.', 'ambiguous-target');
  const selected = sourceRecord.blocks.filter((block) => {
    const state = dispositions.decisions.get(
      payloadGroupDecisionKey(sourceRecord.id, block.blockId)
    );
    return (!blockIds || blockIds.has(block.blockId)) && !state;
  });
  if (!selected.length || (blockIds && selected.length !== blockIds.size))
    return fail(
      'The requested property groups are missing, resolved, or conflicting.',
      'unsupported-operation'
    );
  const record = { ...sourceRecord, blocks: selected },
    index = buildBlockIndex(pkg),
    extension = record.extension,
    labels = extension?.fallbackLabels;
  if (
    extension?.executor !== (atomic ? ATOMIC_REVIEW_EXECUTOR : 'sciobot-properties-v1') ||
    !Array.isArray(labels) ||
    labels.length !== sourceRecord.blocks.length
  )
    return fail(
      'The property proposal has no supported fallback execution metadata.',
      'unsupported-operation'
    );
  if (atomic) {
    const pending = readStoredPropertyProposals(pkg, current);
    if (!pending.success) return pending;
    const checked = projectPendingProperties(pkg, current, pending.pending, 'working');
    if (!checked.success) return checked;
  }
  const labelValues = new Map<
    string,
    { before: string; after: string; textBefore: string; textAfter: string }
  >();
  for (const raw of labels) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw))
      return fail('Invalid stored fallback labels.', 'invalid-document');
    const value = raw as JsonObject;
    if (
      Object.keys(value).some(
        (key) => !['blockId', 'before', 'after', 'textBefore', 'textAfter'].includes(key)
      ) ||
      typeof value.blockId !== 'string' ||
      typeof value.before !== 'string' ||
      typeof value.after !== 'string' ||
      typeof value.textBefore !== 'string' ||
      typeof value.textAfter !== 'string' ||
      labelValues.has(value.blockId)
    )
      return fail('Stored fallback labels need exact unique block ownership.', 'invalid-document');
    labelValues.set(value.blockId, {
      before: value.before,
      after: value.after,
      textBefore: value.textBefore,
      textAfter: value.textAfter,
    });
  }
  for (const id of labelValues.keys())
    if (!record.blocks.some((block) => block.blockId === id)) labelValues.delete(id);
  const paragraphIds = new Set<string>(),
    changedParagraphs = new Set<string>();
  for (const saved of record.blocks) {
    if (atomic && (saved.changes.length !== 1 || saved.changes[0]!.name !== ATOMIC_PAYLOAD_CHANGE))
      return fail(
        'Atomic proposals require one retained payload pair per block.',
        'invalid-document'
      );
    const owner = index.byId.get(saved.blockId),
      stored = payloads.entries.get(saved.blockId);
    if (
      !owner ||
      owner.payload.type !== saved.type ||
      owner.partName !== part.name ||
      !stored ||
      !labelValues.has(saved.blockId)
    )
      return fail('A property decision has no exact current native owner.', 'invalid-document');
    const paragraph = atomic
      ? atomicFallbackParagraph(part, owner.controlId)
      : propertyFallbackParagraph(part, owner.controlId);
    if (!paragraph)
      return fail('A generated property fallback is missing or ambiguous.', 'invalid-document');
    const restriction = propertyControlRefusal(part, paragraph.id);
    if (restriction) return fail(restriction, 'unsupported-operation');
    paragraphIds.add(paragraph.id);
    const fallback = labelValues.get(saved.blockId)!;
    if ((atomic || data.decision === 'accept') && stored.label !== fallback.before)
      return fail(
        'The stored fallback label changed while this proposal was pending.',
        'stale-revision'
      );
    if (fallback.textBefore !== fallback.textAfter) changedParagraphs.add(paragraph.id);
    for (const change of saved.changes) {
      const attributes = owner.payload.attributes;
      if (
        (atomic ? true : Object.hasOwn(attributes, change.name)) !== change.before.present ||
        (change.before.present &&
          jsonDataKey(atomic ? owner.payload : attributes[change.name]!) !==
            jsonDataKey(change.before.value))
      )
        return fail(
          'The committed property changed while this proposal was pending.',
          'stale-revision'
        );
    }
  }
  for (const other of store.records.values()) {
    if (other.id === record.id) continue;
    if (other.decoded.status === 'unsupported-version')
      return fail('Another property proposal has an unsupported version.', 'unsupported-operation');
    if (
      other.decoded.value.blocks.some(
        (block) =>
          labelValues.has(block.blockId) &&
          (!dispositions.decisions.has(payloadGroupDecisionKey(other.id, block.blockId)) ||
            dispositions.decisions.get(payloadGroupDecisionKey(other.id, block.blockId))?.status ===
              'conflict')
      )
    )
      return fail(
        'Overlapping pending property decisions need group reconciliation.',
        'unsupported-operation'
      );
  }
  const sites = collectRevisionSites(part),
    ops: TreeDocOp[] = [];
  if (
    atomic &&
    sites.some((site) => {
      const address = nativePropertyAddress(site.node);
      return (
        address &&
        paragraphIds.has(paragraphOwner(part, site.node) ?? '') &&
        address.author === record.author &&
        address.date === record.date &&
        !record.nativeRevisions.some(
          (saved) =>
            propertyAddressKey(saved.partName, saved) === propertyAddressKey(part.name, address)
        )
      );
    })
  )
    return fail('An atomic fallback has an unrecorded native revision.', 'stale-revision');
  for (const address of record.nativeRevisions) {
    if (
      address.partName !== part.name ||
      address.author !== record.author ||
      address.date !== record.date
    )
      return fail('A stored fallback revision has foreign ownership.', 'invalid-document');
    const key = propertyAddressKey(address.partName, address),
      matches = sites.filter((site) => {
        const actual = nativePropertyAddress(site.node);
        return (
          actual &&
          propertyAddressKey(part.name, actual) === key &&
          paragraphIds.has(paragraphOwner(part, site.node) ?? '')
        );
      });
    if (
      matches.some(
        (site) =>
          site.refused ||
          !['ins', 'del'].includes(site.node.localName) ||
          !paragraphIds.has(paragraphOwner(part, site.node) ?? '')
      )
    )
      return fail(
        'A native property fallback revision is missing, resolved, or outside its owned label.',
        'stale-revision'
      );
    if (!matches.length) {
      const matchesElsewhere = sites.some((site) => {
        const actual = nativePropertyAddress(site.node);
        return actual && propertyAddressKey(part.name, actual) === key;
      });
      const hasResolvedBlock = sourceRecord.blocks.some((block) => {
        const state = dispositions.decisions.get(
          payloadGroupDecisionKey(sourceRecord.id, block.blockId)
        );
        return state && state.status !== 'conflict';
      });
      if (!matchesElsewhere && !hasResolvedBlock)
        return fail('A retained native property revision is missing.', 'stale-revision');
      continue;
    }
    ops.push({
      op: data.decision === 'accept' ? 'acceptRevision' : 'rejectRevision',
      revision: {
        id: address.id,
        author: address.author,
        ...(address.date !== undefined ? { date: address.date } : {}),
      },
      siteNodeIds: matches.map((site) => site.node.id),
    });
  }
  for (const id of changedParagraphs) {
    const owned = sites.some((site) => {
      const address = nativePropertyAddress(site.node);
      return (
        paragraphOwner(part, site.node) === id &&
        address &&
        record.nativeRevisions.some(
          (saved) =>
            propertyAddressKey(saved.partName, saved) === propertyAddressKey(part.name, address)
        )
      );
    });
    if (!owned)
      return fail(
        'Changed fallback text has no retained native revision receipt.',
        'stale-revision'
      );
  }
  let preview = part;
  for (const op of ops) {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      return fail(`Native property decision refused: ${result.reason}.`, 'unsupported-operation');
    preview = result.part;
  }
  return {
    success: true,
    ops,
    partName: part.name,
    apply(working) {
      let next = working;
      if (data.decision === 'accept')
        for (const saved of record.blocks) {
          const owner = buildBlockIndex(next).byId.get(saved.blockId)!,
            changed = new Set(saved.changes.map((change) => change.name)),
            attributes = Object.fromEntries([
              ...Object.entries(owner.payload.attributes).filter(([key]) => !changed.has(key)),
              ...saved.changes.flatMap((change) =>
                change.after.present ? [[change.name, change.after.value]] : []
              ),
            ]);
          const written = withPayloadRecord(
            next,
            atomic && saved.changes[0]!.after.present
              ? (saved.changes[0]!.after.value as BlockPayload)
              : { ...owner.payload, attributes },
            labelValues.get(saved.blockId)!.after
          );
          if (!written.ok) throw new Error(written.reason);
          next = written.pkg;
        }

      return next;
    },
  };
}
