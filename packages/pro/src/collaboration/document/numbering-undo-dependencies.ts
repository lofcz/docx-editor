/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { DocumentRegistry } from './registry.ts';
import { nodeRecordDeleteFilter } from './undo-delete-filter.ts';
import { retainsNumberingInfrastructure } from './undo-numbering.ts';
import { RESTORE_ORIGIN } from './paragraph-text-restore.ts';
import { neededNumberingMembers } from './numbering-live-references.ts';
import type { LogicalId } from './identity.ts';
import {
  attributeMapKey,
  isElementRecord,
  namespaceIdOf,
  parseAttributeMapKey,
  parseBindingMapKey,
  nodeRecordUndone,
  NODE_UNDONE_FIELD,
  type ElementRecord,
} from './schema.ts';
import { holderItem, mapEntry } from './yjs-items.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
type Step = Y.UndoManager['undoStack'][number];

/** Retain allocation data for delayed references; the materializer controls its visibility. */
export class NumberingUndoDependencies {
  private readonly kept = new Set<LogicalId>();
  private required: ReadonlySet<LogicalId> | null = null;
  private readonly invalidate = () => {
    this.required = null;
  };
  private readonly nodeChanges = (events: Y.YEvent<Y.AbstractType<unknown>>[]) => {
    if (events.some((event) => !(event.target instanceof Y.Text))) this.invalidate();
  };
  private rescuing = false;
  private readonly after = () => {
    this.clear();
    if (
      this.rescuing ||
      !this.ready() ||
      !this.registry.partEntries().some((part) => part.name === '/word/numbering.xml')
    )
      return;
    const needed = (this.required ??= neededNumberingMembers(this.registry));
    const restored = [...needed].filter((id) =>
      nodeRecordUndone(this.registry.schema.nodes.get(id))
    );
    if (!restored.length) return;
    // Restore availability using the existing undo mark, so every reader sees the same
    // live resource. The data was retained even if this reference arrived after undo.
    this.rescuing = true;
    try {
      this.registry.doc.transact(() => {
        for (const id of restored) this.registry.schema.nodes.get(id)?.delete(NODE_UNDONE_FIELD);
      }, RESTORE_ORIGIN);
    } finally {
      this.rescuing = false;
    }
  };
  constructor(
    private readonly registry: DocumentRegistry,
    private readonly ready: () => boolean
  ) {
    registry.doc.on('beforeTransaction', this.before);
    registry.doc.on('afterTransaction', this.after);
    registry.schema.attributes.observe(this.invalidate);
    registry.schema.nodes.observeDeep(this.nodeChanges);
    registry.schema.parts.observeDeep(this.invalidate);
    registry.schema.namespaces.observe(this.invalidate);
  }
  reconcile(): void {
    this.invalidate();
    this.after();
  }
  deleteFilter(): (item: Y.Item) => boolean {
    // No descriptor maps: keeping the attributes of nodes an undo un-creates leaves those nodes
    // shown after the undo of a paste, a list, or a block birth in this fork's history.
    const nodeFilter = nodeRecordDeleteFilter(this.registry.schema.nodes);
    return (item) =>
      !this.retains(item) &&
      !retainsNumberingInfrastructure(this.registry.schema, item) &&
      nodeFilter(item);
  }
  destroy(): void {
    this.registry.doc.off('beforeTransaction', this.before);
    this.registry.doc.off('afterTransaction', this.after);
    this.registry.schema.attributes.unobserve(this.invalidate);
    this.registry.schema.nodes.unobserveDeep(this.nodeChanges);
    this.registry.schema.parts.unobserveDeep(this.invalidate);
    this.registry.schema.namespaces.unobserve(this.invalidate);
    this.required = null;
    this.clear();
  }
  private readonly clear = () => {
    this.kept.clear();
  };
  private readonly before = (transaction: Y.Transaction) => {
    this.clear();
    const manager = transaction.origin;
    if (!(manager instanceof Y.UndoManager) || !manager.undoing) return;
    const step = manager.undoStack.at(-1);
    if (!step || !this.registry.partEntries().some((part) => part.name === '/word/numbering.xml'))
      return;
    const elements = this.registry.allLogicalIds().flatMap((id) => {
      const record = this.registry.record(id);
      return record &&
        isElementRecord(record) &&
        record.namespaceUri === W &&
        ['num', 'abstractNum', 'numId'].includes(record.localName) &&
        !this.registry.isTombstoned(id)
        ? [record]
        : [];
    });
    const nums = elements.filter((node) => node.localName === 'num');
    const abstracts = elements.filter((node) => node.localName === 'abstractNum');
    const allocated = (node: ElementRecord, name: string) => {
      const birth = holderItem(this.registry.schema.nodes.get(node.logicalId));
      return (
        (birth !== null && Y.isDeleted(step.insertions, birth.id)) ||
        this.insertedAttribute(step, node.logicalId, name)
      );
    };
    if (
      !abstracts.some((node) => allocated(node, 'abstractNumId')) &&
      !nums.some((node) => allocated(node, 'numId'))
    )
      return;
    // Keep allocation data for references that can arrive after this undo. Visibility is
    // controlled by the existing undone marks and the live numbering dependency view.
    for (const num of nums) if (allocated(num, 'numId')) this.retainSubtree(num.logicalId);
    for (const abstract of abstracts)
      if (allocated(abstract, 'abstractNumId')) this.retainSubtree(abstract.logicalId);
  };
  private insertedAttribute(step: Step, id: LogicalId, name: string): boolean {
    const item = mapEntry(
      this.registry.schema.attributes,
      attributeMapKey(id, namespaceIdOf(W), name)
    );
    return item !== undefined && Y.isDeleted(step.insertions, item.id);
  }
  private retainSubtree(id: LogicalId): void {
    if (this.kept.has(id)) return;
    this.kept.add(id);
    const record = this.registry.record(id);
    if (record && isElementRecord(record))
      for (const child of record.childIds) this.retainSubtree(child);
  }
  retains(item: Y.Item): boolean {
    if (!this.kept.size) return false;
    if (item.parent === this.registry.schema.attributes && item.parentSub) {
      const key = parseAttributeMapKey(item.parentSub);
      return key !== null && this.kept.has(key.logicalId);
    }
    if (item.parent === this.registry.schema.bindings && item.parentSub) {
      const key = parseBindingMapKey(item.parentSub);
      return key !== null && this.kept.has(key.logicalId);
    }
    const container = holderItem(item.parent);
    const node = holderItem(container?.parent);
    return (
      node?.parent === this.registry.schema.nodes &&
      node.parentSub !== null &&
      this.kept.has(node.parentSub as LogicalId)
    );
  }
}
