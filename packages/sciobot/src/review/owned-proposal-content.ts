import { ownedStructuralContent } from './owned-structural-content';
import { sha256FontBytes } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  findNode,
  parentNodeOf,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { EDIT_PLAN_LIMITS } from '../diff/plan';
import { jsonDataKey } from '../diff/target-key';
import { GROUP_ID_LIMIT } from './group-markers';
import { readReviewGroups } from './groups';
import { ownedProposalSelectionProblem } from './owned-proposal-input';
import type { OwnedProposalGroup } from './owned-proposal-plan';
import { propertyAddressKey } from './property-fallback';
import { payloadFallback } from './payload-fallback';
import { ATOMIC_REVIEW_EXECUTOR } from './property-record-schema';
import { readPropertyReviewStore } from './property-record-store';
import { decodeBlockPresenceRecord } from './block-presence-schema';
import { bindBlockPresenceRecord } from './block-presence-plan';
import { currentRevisionGuardContent } from './revision-property-content';

export interface OwnedProposalContentProof {
  readonly schemaVersion: 1;
  readonly groupId: string;
  readonly fingerprint: `sha256:${string}`;
}

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const object = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const exactKeys = (value: object, names: readonly string[]) =>
  Object.keys(value).length === names.length &&
  Object.keys(value).every((key) => names.includes(key));
const id = (value: unknown): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= GROUP_ID_LIMIT;
const propertyKey = (value: { proposalId: string; blockId: string }) =>
  JSON.stringify([value.proposalId, value.blockId]);
const equalSets = (left: ReadonlySet<string>, right: ReadonlySet<string>) =>
  left.size === right.size && [...left].every((value) => right.has(value));
// Canonical trees have a separate depth bound from caller-owned JSON receipt data.
const MAX_CANONICAL_DEPTH = 256;

/** Read-only semantic content proof. No inverse previews or native write primitives run here. */
export function captureOwnedProposalContentProofs(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  ownership: readonly OwnedProposalGroup[]
):
  | { readonly success: true; readonly contentProofs: readonly OwnedProposalContentProof[] }
  | NativeEditFailure {
  try {
    const copied = copyBoundedJson({ revision, ownership }, EDIT_PLAN_LIMITS);
    if (!copied.ok) return fail(copied.reason);
    const input = copied.value as unknown as {
      revision: RevisionToken;
      ownership: readonly OwnedProposalGroup[];
    };
    const token = input.revision;
    if (
      !object(token) ||
      !exactKeys(token, ['documentId', 'epoch', 'revision']) ||
      !id(token.documentId) ||
      !id(token.epoch) ||
      !Number.isSafeInteger(token.revision) ||
      token.revision < 0
    )
      return fail('Invalid ownership content revision.');
    const problem = ownedProposalSelectionProblem(input.ownership);
    if (problem) return fail(problem);
    const read = readReviewGroups(pkg);
    if (!read.success) return read;
    const groups = new Map(read.groups.map((group) => [group.id, group]));
    const hasProperties = input.ownership.some((owned) => owned.properties.length > 0);
    const records = hasProperties ? readPropertyReviewStore(pkg) : null;
    const blocks = hasProperties ? buildBlockIndex(pkg) : null;
    const proofs: OwnedProposalContentProof[] = [];
    let visited = 0,
      bytes = 0;
    const semantic = (root: OoxmlNode): string => {
      const pending = [{ node: root, depth: 0 }];
      while (pending.length) {
        const { node, depth } = pending.pop()!;
        if (++visited > EDIT_PLAN_LIMITS.maxNodes || depth > MAX_CANONICAL_DEPTH)
          throw new Error('Owned proposal content exceeds its traversal limit.');
        if (node.kind === 'textValue') bytes += node.value.length;
        else {
          bytes += node.localName.length + node.namespaceUri.length;
          for (const attribute of node.attributes)
            bytes +=
              attribute.localName.length + attribute.namespaceUri.length + attribute.value.length;
          for (const child of node.children) pending.push({ node: child, depth: depth + 1 });
        }
        if (bytes > EDIT_PLAN_LIMITS.maxBytes)
          throw new Error('Owned proposal content exceeds its byte limit.');
      }
      return canonicalOoxmlFingerprint(root);
    };
    for (const owned of input.ownership) {
      const group = groups.get(owned.groupId);
      if (!group) return fail('An owned pending group is missing or resolved.', 'stale-revision');
      if (group.readOnly)
        return fail(
          'An owned group needs reconciliation before content capture.',
          'unsupported-operation'
        );
      const expectedNative = new Set(
        owned.native.flatMap((member) =>
          member.addresses.map((address) => propertyAddressKey(member.partName, address))
        )
      );
      if (
        !equalSets(
          expectedNative,
          new Set(group.native.map((member) => propertyAddressKey(member.partName, member.address)))
        ) ||
        !equalSets(
          new Set(owned.properties.map(propertyKey)),
          new Set(group.properties.map(propertyKey))
        )
      )
        return fail('Owned receipts do not match the complete current group.', 'ambiguous-target');
      const selectedSites = new Map<string, Set<string>>();
      for (const member of group.native) {
        const ids = selectedSites.get(member.partName) ?? new Set<string>();
        member.siteNodeIds.forEach((id) => ids.add(id));
        selectedSites.set(member.partName, ids);
      }
      const native: string[] = [];
      const normalizedParts = new Map<string, OoxmlPart>();
      for (const member of group.native) {
        const part = pkg.parts.get(member.partName);
        if (!part || !member.siteNodeIds.length)
          return fail('An owned native site is missing.', 'invalid-document');
        const sites: string[] = [];
        for (const nodeId of member.siteNodeIds) {
          const node = findNode(part, nodeId);
          if (!node) return fail('An owned native site is missing.', 'invalid-document');
          const parent = parentNodeOf(part, nodeId);
          let guarded: OoxmlNode | null = node;
          if (parent?.kind === 'paragraph') {
            const key = JSON.stringify([part.name, parent.id]);
            let normalized = normalizedParts.get(key);
            if (!normalized) {
              const root = ownedStructuralContent(parent, selectedSites.get(member.partName)!);
              if (root.kind === 'textValue')
                return fail('Invalid native proof root.', 'invalid-document');
              normalized = { ...part, root };
              normalizedParts.set(key, normalized);
            }
            guarded = findNode(normalized, nodeId);
          }
          // Coalesced fragments retain their first site identity. The original site
          // must still exist above; only the read-only proof projection can omit it.
          if (!guarded) continue;
          const saved = semantic(guarded);
          // An inverse restores current property bags or removes a marked paragraph.
          // Committed suffixes of independently marked paragraphs survive the inverse.
          // All other content outside the revision subtree remains guarded.
          const currentContent = currentRevisionGuardContent(part, node);
          sites.push(
            currentContent
              ? JSON.stringify([
                  saved,
                  semantic(
                    ownedStructuralContent(currentContent, selectedSites.get(member.partName)!)
                  ),
                ])
              : saved
          );
        }
        native.push(
          JSON.stringify([propertyAddressKey(member.partName, member.address), sites.sort()])
        );
      }
      const properties: string[] = [];
      for (const member of group.properties) {
        const record =
          records?.status === 'ready' ? records.records.get(member.proposalId) : undefined;
        const block = blocks?.byId.get(member.blockId);
        if (!record || !block)
          return fail('An owned property record or block is missing.', 'invalid-document');
        if (record.decoded.status !== 'supported') {
          const decoded = decodeBlockPresenceRecord(record.raw);
          if (decoded.status !== 'supported')
            return fail('An owned presence record is unsupported.', 'unsupported-operation');
          const saved = decoded.value.blocks.find((root) => root.blockId === member.blockId);
          if (!saved) return fail('An owned presence root is missing.', 'invalid-document');
          const bound = bindBlockPresenceRecord(pkg, { ...decoded.value, blocks: [saved] });
          if (!bound.success) return bound;
          const part = pkg.parts.get(block.partName),
            node = part && findNode(part, block.controlId);
          if (!node) return fail('An owned presence control is missing.', 'invalid-document');
          const bounded = copyBoundedJson(bound.record, EDIT_PLAN_LIMITS);
          if (!bounded.ok) return fail(bounded.reason);
          properties.push(
            JSON.stringify([propertyKey(member), jsonDataKey(bounded.value), semantic(node)])
          );
          continue;
        }
        const saved = record.decoded.value.blocks.find((entry) => entry.blockId === member.blockId);
        const part = pkg.parts.get(block.partName);
        if (!saved || !part) return fail('An owned property block is missing.', 'invalid-document');
        const atomic = record.decoded.value.extension?.executor === ATOMIC_REVIEW_EXECUTOR;
        // Property decisions own their named attributes and native revision sites, not an
        // unrelated suffix in the generated label. Atomic decisions own the complete fallback.
        const fallback = atomic ? payloadFallback(part, block.controlId, true)?.node : null;
        const current = atomic
          ? block.payload
          : saved.changes.map((change) => [
              change.name,
              Object.hasOwn(block.payload.attributes, change.name)
                ? { present: true, value: block.payload.attributes[change.name] }
                : { present: false },
            ]);
        const bounded = copyBoundedJson(
          { record: record.decoded.value, current },
          EDIT_PLAN_LIMITS
        );
        if (!bounded.ok) return fail(bounded.reason);
        properties.push(
          JSON.stringify([
            propertyKey(member),
            jsonDataKey(bounded.value),
            fallback ? semantic(fallback) : null,
          ])
        );
      }
      const content = JSON.stringify([1, group.id, native.sort(), properties.sort()]);
      const encoded = new TextEncoder().encode(content);
      bytes += encoded.byteLength;
      if (bytes > EDIT_PLAN_LIMITS.maxBytes)
        return fail('Owned proposal content exceeds its byte limit.');
      proofs.push(
        Object.freeze({
          schemaVersion: 1,
          groupId: group.id,
          fingerprint: sha256FontBytes(encoded) as `sha256:${string}`,
        })
      );
    }
    return Object.freeze({ success: true, contentProofs: Object.freeze(proofs) });
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'invalid-document');
  }
}

/** Verify active selections only. Valid inactive proofs remain available to the host for undo. */
export function verifyOwnedProposalContentProofs(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  activeSelection: readonly OwnedProposalGroup[],
  proofs: readonly OwnedProposalContentProof[]
): NativeEditFailure | null {
  try {
    const copied = copyBoundedJson(proofs, EDIT_PLAN_LIMITS);
    if (!copied.ok) return fail(copied.reason);
    if (!Array.isArray(copied.value)) return fail('Ownership content proofs must be an array.');
    const byId = new Map<string, string>();
    for (const proof of copied.value) {
      if (
        !object(proof) ||
        !exactKeys(proof, ['schemaVersion', 'groupId', 'fingerprint']) ||
        proof.schemaVersion !== 1 ||
        !id(proof.groupId) ||
        byId.has(proof.groupId) ||
        typeof proof.fingerprint !== 'string' ||
        !/^sha256:[a-f0-9]{64}$/.test(proof.fingerprint)
      )
        return fail('Invalid or duplicate ownership content proof.');
      byId.set(proof.groupId, proof.fingerprint);
    }
    const current = captureOwnedProposalContentProofs(pkg, revision, activeSelection);
    if (!current.success) return current;
    for (const proof of current.contentProofs) {
      if (!byId.has(proof.groupId))
        return fail(
          'An active owned proposal needs an explicit content proof.',
          'ambiguous-target'
        );
      if (byId.get(proof.groupId) !== proof.fingerprint)
        return fail(
          'Owned proposal content changed after its receipt. Read and reconcile the manual edits.',
          'ambiguous-target'
        );
    }
    return null;
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'invalid-input');
  }
}
