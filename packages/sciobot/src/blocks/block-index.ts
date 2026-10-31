import { activeStorage } from '../registry/active';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { BlockPayload } from './payload-schema';
import { blockIdFromTag, readPayloadStore } from './payload-opc-codec';

export interface IndexedBlock {
  readonly id: string;
  readonly controlId: string;
  readonly partName: string;
  readonly parentId: string | null;
  readonly payload: BlockPayload;
}

export interface BlockIndexDiagnostic {
  readonly code:
    | 'invalid-store'
    | 'malformed-reference'
    | 'duplicate-reference'
    | 'missing-payload'
    | 'unsupported-version'
    | 'orphan-payload';
  readonly id?: string;
  readonly controlId?: string;
  readonly detail?: string;
}

export interface BlockIndex {
  readonly byId: ReadonlyMap<string, IndexedBlock>;
  readonly diagnostics: readonly BlockIndexDiagnostic[];
  readonly omittedDiagnostics: number;
}

/** Resolve immutable package data. Invalid references never trigger a document rewrite. */
export function buildBlockIndex(pkg: OoxmlPackage, maxDiagnostics = 64): BlockIndex {
  if (!Number.isSafeInteger(maxDiagnostics) || maxDiagnostics < 0 || maxDiagnostics > 256)
    throw new TypeError('Invalid diagnostic limit');
  const diagnostics: BlockIndexDiagnostic[] = [];
  let omittedDiagnostics = 0;
  const report = (diagnostic: BlockIndexDiagnostic) => {
    if (diagnostics.length < maxDiagnostics) diagnostics.push(diagnostic);
    else omittedDiagnostics++;
  };
  const candidates: { id: string; controlId: string; partName: string; parents: string[] }[] = [];
  const counts = new Map<string, number>();
  for (const [partName, part] of pkg.parts) {
    if (!isStoryPart(part)) continue;
    for (const { node, ancestors } of contentControlsIn(part.root)) {
      const tag = contentControlPropertiesOf(node).tag ?? '';
      const id = blockIdFromTag(tag);
      if (id === null) {
        if (tag.startsWith(activeStorage().tagPrefix))
          report({ code: 'malformed-reference', controlId: node.id });
        continue;
      }
      counts.set(id, (counts.get(id) ?? 0) + 1);
      const parents = ancestors
        .map((ancestor) => blockIdFromTag(contentControlPropertiesOf(ancestor).tag ?? ''))
        .filter((parent): parent is string => parent !== null);
      candidates.push({ id, controlId: node.id, partName, parents });
    }
  }
  const store = readPayloadStore(pkg);
  const byId = new Map<string, IndexedBlock>();
  if (store.status === 'invalid') report({ code: 'invalid-store', detail: store.reason });
  const duplicated = new Set<string>();
  const accepted: { candidate: (typeof candidates)[number]; payload: BlockPayload }[] = [];
  for (const candidate of candidates) {
    if (counts.get(candidate.id)! > 1) {
      if (!duplicated.has(candidate.id)) report({ code: 'duplicate-reference', id: candidate.id });
      duplicated.add(candidate.id);
      continue;
    }
    const stored = store.status === 'ready' ? store.entries.get(candidate.id) : undefined;
    if (!stored) {
      report({ code: 'missing-payload', id: candidate.id, controlId: candidate.controlId });
      continue;
    }
    if (stored.decoded.status === 'unsupported-version') {
      report({ code: 'unsupported-version', id: candidate.id, controlId: candidate.controlId });
      continue;
    }
    accepted.push({ candidate, payload: stored.decoded.value });
  }
  const knownIds = new Set(accepted.map(({ candidate }) => candidate.id));
  for (const { candidate, payload } of accepted)
    byId.set(candidate.id, {
      id: candidate.id,
      controlId: candidate.controlId,
      partName: candidate.partName,
      payload,
      parentId: [...candidate.parents].reverse().find((id) => knownIds.has(id)) ?? null,
    });
  if (store.status === 'ready')
    for (const id of store.entries.keys()) {
      if (!counts.has(id)) report({ code: 'orphan-payload', id });
    }
  return { byId, diagnostics, omittedDiagnostics };
}
