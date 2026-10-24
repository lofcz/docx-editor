/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import {
  IMAGE_RELATIONSHIP_TYPE,
  relsPartNameFor,
  resolveInternalTarget,
  type OoxmlPart,
  type OoxmlNode,
  type RelationshipRecord,
} from '@docx-editor.dev/core/store';
import type { DocumentRegistry } from './registry.ts';
import type { BlobBytesStore } from './seed.ts';
import { ImageReferenceIndex, type ImageReferenceState } from './image-reference-index.ts';
import {
  retiredImageRecords,
  retainedRelationshipsPart,
  storageKeyOf,
} from './retired-image-records.ts';
import { emptyRelsPart, relationshipChildrenOf } from './materialize-rels.ts';
import { withRelsChildren, countBlobBytes } from './materialize-freeze.ts';
interface Patch {
  readonly references: ImageReferenceState;
  readonly baseRelationships: ReadonlyMap<string, readonly RelationshipRecord[]>;
  readonly baseBytes: ReadonlyMap<string, Uint8Array>;
  readonly relationships: ReadonlyMap<string, readonly RelationshipRecord[]>;
  readonly bytes: Map<string, Uint8Array>;
  readonly parts: ReadonlyMap<string, OoxmlPart>;
  readonly overrides: ReadonlyMap<string, string>;
  readonly sourceParts: ReadonlyMap<string, OoxmlPart | undefined>;
}
/** Retired descriptors become visible only through actual live canonical references. */
export class RetiredImageResources {
  private readonly references = new ImageReferenceIndex();
  private records: ReturnType<typeof retiredImageRecords> | null = null;
  private patch: Patch | null = null;
  private readonly xmlNodes = new Set<string>();
  constructor(
    private readonly registry: DocumentRegistry,
    private readonly blobs: BlobBytesStore
  ) {}
  invalidate(): void {
    this.records = null;
    this.patch = null;
    this.xmlNodes.clear();
  }
  invalidateNodes(ids: ReadonlySet<string>): void {
    for (const id of ids) {
      if (this.xmlNodes.has(id)) {
        this.patch = null;
        return;
      }
    }
  }
  private watchXml(node: OoxmlNode): void {
    this.xmlNodes.add(node.id);
    if (node.kind !== 'textValue') for (const child of node.children) this.watchXml(child);
  }
  recover(
    parts: Map<string, OoxmlPart>,
    baseRelationships: ReadonlyMap<string, readonly RelationshipRecord[]>,
    baseBytes: Map<string, Uint8Array>,
    overrides: Map<string, string>
  ): {
    readonly relationships: ReadonlyMap<string, readonly RelationshipRecord[]>;
    readonly bytes: Map<string, Uint8Array>;
  } | null {
    const records = (this.records ??= retiredImageRecords(this.registry));
    if (!records.binaries.size && !records.relationships.length)
      return { relationships: baseRelationships, bytes: baseBytes };
    const references = this.references.read(parts);
    if (
      !this.patch ||
      this.patch.references !== references ||
      this.patch.baseRelationships !== baseRelationships ||
      this.patch.baseBytes !== baseBytes ||
      [...this.patch.sourceParts].some(([name, part]) => parts.get(name) !== part)
    ) {
      this.xmlNodes.clear();
      const sourceParts = new Map<string, OoxmlPart | undefined>();
      const byOwner = new Map(baseRelationships),
        bytes = new Map(baseBytes),
        repairedParts = new Map<string, OoxmlPart>(),
        restoredTypes = new Map<string, string>();
      const activeNames = new Set([...bytes.keys()].map(storageKeyOf));
      for (const [owner, refs] of references) {
        const wanted = new Set(refs.ids),
          active = baseRelationships.get(owner) ?? [];
        const candidates = [
          ...active.filter((record) => record.type === IMAGE_RELATIONSHIP_TYPE),
          ...records.relationships.filter(
            (record) => record.ownerPart === owner && !active.some((live) => live.id === record.id)
          ),
        ];
        const restored: RelationshipRecord[] = [];
        for (const record of candidates) {
          if (!wanted.has(record.id) || record.targetMode !== 'Internal') continue;
          const target = resolveInternalTarget(owner, record.rawTarget);
          if (!target.ok) continue;
          const key = storageKeyOf(target.partName);
          const binary = records.binaries.get(key);
          if (!activeNames.has(key)) {
            if (!binary) continue;
            const data = this.blobs.get(binary.digest);
            if (!data || data.byteLength !== binary.size) return null;
            countBlobBytes(data.byteLength);
            bytes.set(binary.storageKey, data);
            activeNames.add(key);
            if (binary.retainedOverride) restoredTypes.set(key, binary.mediaType);
          }
          if (!active.some((live) => live.id === record.id)) restored.push(record);
        }
        if (!restored.length) continue;
        const merged = [...active, ...restored].sort(
          (a, b) => a.order - b.order || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0)
        );
        byOwner.set(owner, merged);
        const name = relsPartNameFor(owner);
        sourceParts.set(name, parts.get(name));
        const original = retainedRelationshipsPart(this.registry, name);
        if (original) this.watchXml(original.root);
        const base = parts.get(name) ?? original ?? emptyRelsPart(name);
        this.watchXml(base.root);
        // The original children retain extensions on a retired relationship row.
        const previous = original
          ? [
              ...base.root.children,
              ...original.root.children.filter(
                (child) => !base.root.children.some((existing) => existing.id === child.id)
              ),
            ]
          : base.root.children;
        repairedParts.set(
          name,
          withRelsChildren(base, relationshipChildrenOf(previous, merged, name))
        );
      }
      this.patch = {
        references,
        sourceParts,
        baseRelationships,
        baseBytes,
        relationships: byOwner,
        bytes,
        parts: repairedParts,
        overrides: restoredTypes,
      };
    }
    for (const [name, part] of this.patch.parts) parts.set(name, part);
    for (const [name, type] of this.patch.overrides) overrides.set(name, type);
    return { relationships: this.patch.relationships, bytes: this.patch.bytes };
  }
}
