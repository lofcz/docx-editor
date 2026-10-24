/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type * as Y from 'yjs';
import type { CanonicalPrimitiveJournal } from '@docx-editor.dev/core/collaboration/replication';
import { relationshipKey } from './relationship-store.ts';
import type { DocumentRegistry } from './registry.ts';
const HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
/** Native URL allocation is persistent package metadata, not a second user undo gesture. */
export function withoutPersistentHyperlinkUndo<T>(
  manager: Y.UndoManager,
  origin: unknown,
  registry: Pick<DocumentRegistry, 'doc' | 'relationships'> & {
    readonly schema: Pick<DocumentRegistry['schema'], 'relationships'>;
  },
  journal: Pick<CanonicalPrimitiveJournal, 'effects'>,
  run: () => T
): T {
  const effects = journal.effects;
  if (
    !effects.length ||
    !effects.every(
      (e) =>
        e.kind === 'putRelationship' &&
        e.record.type === HYPERLINK &&
        e.record.targetMode === 'External'
    )
  )
    return run();
  const existing = new Set(registry.relationships().map((r) => `${r.ownerPart}\0${r.id}`));
  if (
    effects.some((e) => e.kind === 'putRelationship' && existing.has(`${e.owner}\0${e.record.id}`))
  )
    return run();
  const previous = manager.captureTransaction;
  let mint: Y.Transaction | null = null;
  const before = (transaction: Y.Transaction) => {
    if (transaction.origin !== origin || mint) return;
    mint = transaction;
    registry.doc.off('beforeTransaction', before);
  };
  const keys = new Set(
    effects.flatMap((e) =>
      e.kind === 'putRelationship' ? [relationshipKey(e.owner, e.record.id)] : []
    )
  );
  registry.doc.on('beforeTransaction', before);
  manager.captureTransaction = (transaction) => {
    if (!previous.call(manager, transaction)) return false;
    if (transaction !== mint) return true;
    const root = registry.schema.relationships;
    const comparableRoot: unknown = root;
    const holders = new Set<unknown>([...keys].map((key) => root.get(key)));
    for (const [type, changed] of transaction.changed) {
      if (type === comparableRoot) {
        if ([...changed].some((key) => key === null || !keys.has(key))) return true;
      } else if (!holders.has(type)) return true;
    }
    return false;
  };
  try {
    return run();
  } finally {
    manager.captureTransaction = previous;
    registry.doc.off('beforeTransaction', before);
  }
}
