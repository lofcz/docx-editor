/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { DocumentRegistry } from './registry.ts';
import { asLogicalId, type LogicalId } from './identity.ts';
import { deleteSetRuns, holderItem } from './yjs-items.ts';
import {
  parseAttributeMapKey,
  parseBindingMapKey,
  isElementRecord,
  isNodeMap,
  NODE_CHILDREN_FIELD,
  NODE_UNDONE_FIELD,
  namespaceIdOf,
} from './schema.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const BIRTH_GEOMETRY_NAMES = new Set([
  'gridCol',
  'gridSpan',
  'vMerge',
  'tcW',
  'tblW',
  'gridBefore',
  'gridAfter',
]);

/** Keep the ancestry of a birth available for foreign content that can arrive after undo. */
export class BirthUndoDependencies {
  private readonly born = new Set<LogicalId>();
  constructor(private readonly registry: DocumentRegistry) {
    registry.doc.on('beforeTransaction', this.before);
    registry.doc.on('afterTransaction', this.after);
  }
  destroy(): void {
    this.registry.doc.off('beforeTransaction', this.before);
    this.registry.doc.off('afterTransaction', this.after);
    this.born.clear();
  }
  private readonly after = () => this.born.clear();
  private readonly before = (transaction: Y.Transaction) => {
    this.born.clear();
    const manager = transaction.origin;
    if (!(manager instanceof Y.UndoManager) || !manager.undoing) return;
    const step = manager.undoStack.at(-1);
    if (!step) return;
    const root = this.registry
      .partEntries()
      .find((part) => part.name === this.registry.mainDocumentPart())?.rootLogicalId;
    if (!root) return;
    const inBody = (id: LogicalId) => {
      const seen = new Set<LogicalId>();
      for (
        let at: LogicalId | null = id;
        at && seen.size < this.registry.limits.maxTreeDepth;
        at = this.registry.parentOf(at)
      ) {
        if (at === root) return true;
        if (seen.has(at)) return false;
        seen.add(at);
      }
      return false;
    };
    const changed: Y.Item[] = [];
    for (const run of deleteSetRuns(step.insertions)) {
      for (let clock = run.clock; clock < run.clock + run.length; ) {
        const item = Y.getItem(this.registry.doc.store, Y.createID(run.client, clock));
        clock = item.id.clock + item.length;
        // A collected character range must not hide later births in the same history step.
        if (!(item instanceof Y.Item)) continue;
        changed.push(item);
        if (item.parent === this.registry.schema.nodes && item.parentSub) {
          const node = this.registry.schema.nodes.get(item.parentSub);
          if (isNodeMap(node) && holderItem(node) === item && inBody(asLogicalId(item.parentSub)))
            this.born.add(asLogicalId(item.parentSub));
        }
      }
    }
    // This mutation belongs to the undo transaction. Redo reverses these markers too.
    for (const id of this.born) this.registry.schema.nodes.get(id)?.set(NODE_UNDONE_FIELD, true);
    // Split mixed array chunks, so retaining a born child never retains an unrelated move.
    for (const item of changed) {
      if (!(item.parent instanceof Y.Array) || item.deleted || !item.countable) continue;
      const values = item.content.getContent();
      const length = item.length;
      for (let offset = 1; offset < length; offset++) {
        const retained = (value: unknown) =>
          typeof value === 'string' && this.born.has(asLogicalId(value));
        if (retained(values[offset - 1]) !== retained(values[offset]))
          Y.getItemCleanStart(transaction, Y.createID(item.id.client, item.id.clock + offset));
      }
    }
  };
  retains(item: Y.Item): boolean {
    if (!this.born.size) return false;
    if (item.parent instanceof Y.Array) {
      const field = holderItem(item.parent);
      const node = holderItem(field?.parent);
      if (field?.parentSub !== NODE_CHILDREN_FIELD || node?.parent !== this.registry.schema.nodes)
        return false;
      const values = item.content.getContent();
      return (
        values.length > 0 &&
        values.every((value) => typeof value === 'string' && this.born.has(asLogicalId(value)))
      );
    }
    if (item.parent === this.registry.schema.bindings && item.parentSub) {
      const key = parseBindingMapKey(item.parentSub);
      return !!key && this.born.has(key.logicalId);
    }
    if (item.parent === this.registry.schema.attributes && item.parentSub) {
      const key = parseAttributeMapKey(item.parentSub);
      if (!key || !this.born.has(key.logicalId)) return false;
      const record = this.registry.record(key.logicalId);
      return (
        !!record &&
        isElementRecord(record) &&
        ((record.namespaceUri === W && BIRTH_GEOMETRY_NAMES.has(record.localName)) ||
          (key.namespaceId === namespaceIdOf(W14) && ['paraId', 'textId'].includes(key.localName)))
      );
    }
    // Record containers and their immutable shell already use the ordinary tombstone filter.
    return false;
  }
}
