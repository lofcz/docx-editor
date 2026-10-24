/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { relationshipKey } from '../document/relationship-store.ts';
import { withoutPersistentHyperlinkUndo } from '../document/persistent-hyperlink-history.ts';
const record = {
  id: 'r1',
  ownerPart: '/word/document.xml',
  type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  target: 'https://example.com',
  rawTarget: 'https://example.com',
  order: 0,
  targetMode: 'External' as const,
};
const journal = {
  effects: [{ kind: 'putRelationship' as const, owner: record.ownerPart, record }],
};
test('persistent URL mints preserve origin and do not hide a reentrant local content edit from undo', () => {
  const doc = new Y.Doc(),
    relationships = doc.getMap<Y.Map<Y.Map<unknown>>>('relationships'),
    text = doc.getText('text'),
    origin = { kind: 'local' };
  const manager = new Y.UndoManager([relationships, text], { trackedOrigins: new Set([origin]) });
  let once = false;
  doc.on('afterTransaction', (transaction) => {
    if (transaction.origin === origin && !once) {
      once = true;
      doc.transact(() => text.insert(0, 'Peer'), origin);
    }
  });
  const registry = { doc, schema: { relationships }, relationships: () => [] };
  withoutPersistentHyperlinkUndo(manager, origin, registry, journal, () =>
    doc.transact(
      () =>
        relationships.set(
          relationshipKey(record.ownerPart, record.id),
          new Y.Map<Y.Map<unknown>>()
        ),
      origin
    )
  );
  expect(once).toBe(true);
  expect(text.toString()).toBe('Peer');
  expect(manager.undoStack).toHaveLength(1);
  manager.undo();
  expect(text.toString()).toBe('');
  expect(relationships.has(relationshipKey(record.ownerPart, record.id))).toBe(true);
  manager.redo();
  expect(text.toString()).toBe('Peer');
  manager.destroy();
  doc.destroy();
});
test('editing an existing relationship retains its ordinary undo step', () => {
  const doc = new Y.Doc(),
    relationships = doc.getMap<Y.Map<Y.Map<unknown>>>('relationships'),
    origin = { kind: 'local' };
  const manager = new Y.UndoManager(relationships, { trackedOrigins: new Set([origin]) });
  const registry = { doc, schema: { relationships }, relationships: () => [record] };
  withoutPersistentHyperlinkUndo(manager, origin, registry, journal, () =>
    doc.transact(() => relationships.set('existing', new Y.Map<Y.Map<unknown>>()), origin)
  );
  expect(manager.undoStack).toHaveLength(1);
  manager.undo();
  expect(relationships.has('existing')).toBe(false);
  manager.destroy();
  doc.destroy();
});
