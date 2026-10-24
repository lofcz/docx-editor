import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { attachOwnedTextAttributes } from './owned-text';
import { attachOwnedReviewHistory } from './review-history';
import { OWNED_TEXT_OPERATIONS } from './ownership-codec';
import { OWNED_STRUCTURE_FIELD } from './owned-structure';
import type { EditorCollaborationSession } from '@docx-editor.dev/core/collaboration';

test('structural attributes follow native Yjs while authored formatting keeps insertion ownership', () => {
  const a = new Y.Doc(),
    b = new Y.Doc();
  a.clientID = 41;
  b.clientID = 73;
  const scopes = [a, b].map((doc) => attachOwnedTextAttributes(doc, { protocol: 2 }));
  try {
    scopes.forEach((scope) => scope.setStructuralAttributes(['node']));
    const x = a.getText('paragraph');
    x.insert(0, 'AB', { node: 'old', bold: false });
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
    const y = b.getText('paragraph');
    x.format(0, 2, { node: 'new', bold: true });
    y.insert(1, ' peer');
    Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
    expect(x.toString()).toBe('A peerB');
    expect(y.toString()).toBe('A peerB');
    const native = (text: Y.Text) =>
      text.toDelta().map((op: { insert: unknown; attributes?: Record<string, unknown> }) => ({
        insert: op.insert,
        node: op.attributes?.node,
        bold: op.attributes?.bold,
      }));
    expect(native(x)).toEqual(native(y));
    expect(native(x)).toEqual([
      { insert: 'A', node: 'new', bold: true },
      { insert: ' peer', node: 'new', bold: false },
      { insert: 'B', node: 'new', bold: true },
    ]);
  } finally {
    scopes.forEach((scope) => scope.dispose());
    a.destroy();
    b.destroy();
  }
});

test('version 2 refuses legacy rooms and version 1 refuses version 2 rooms before binding', () => {
  for (const [oldVersion, nextVersion] of [
    [1, 2],
    [2, 1],
  ] as const) {
    const doc = new Y.Doc(),
      old = attachOwnedTextAttributes(doc, { protocol: oldVersion });
    doc.getText('text').insert(0, 'Seed');
    old.dispose();
    const before = Y.encodeStateAsUpdate(doc);
    expect(() => attachOwnedTextAttributes(doc, { protocol: nextVersion })).toThrow(
      'Unsupported owned-text protocol'
    );
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    doc.destroy();
  }
});

test('provider policy must agree with the room snapshot and cannot change after binding', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedTextAttributes(doc, { protocol: 2 });
  try {
    scope.setStructuralAttributes(['text', 'run']);
    scope.setStructuralAttributes(['run', 'text']);
    expect(doc.getMap(OWNED_TEXT_OPERATIONS).get(OWNED_STRUCTURE_FIELD)).toBe('["run","text"]');
    expect(() => scope.setStructuralAttributes(['run'])).toThrow('Incompatible structural');
    for (const keys of [['run', 'run'], ['__proto__'], [''], ['sciobot-owned:v1:fake']])
      expect(() => scope.setStructuralAttributes(keys)).toThrow('Invalid structural');
    expect(() => doc.getMap(OWNED_TEXT_OPERATIONS).set(OWNED_STRUCTURE_FIELD, '["other"]')).toThrow(
      'Incompatible structural'
    );
    expect(() => scope.assertHealthy()).toThrow('Incompatible structural');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('complete history refuses a provider without advertised structural keys before writing a policy', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewHistory(doc, { documentId: 'room' });
  try {
    expect(() => scope.bindSession({ documentId: 'room' } as EditorCollaborationSession)).toThrow(
      'provider structural'
    );
    expect(doc.getMap(OWNED_TEXT_OPERATIONS).has(OWNED_STRUCTURE_FIELD)).toBe(false);
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`mixed owned-text headers remain refused after snapshot reopening: ${clients}`, () => {
    const current = new Y.Doc(),
      legacy = new Y.Doc();
    current.clientID = clients[0];
    legacy.clientID = clients[1];
    const scope = attachOwnedTextAttributes(current, { protocol: 2 });
    legacy.getMap(OWNED_TEXT_OPERATIONS).set('protocol', '1');
    try {
      expect(() => Y.applyUpdate(current, Y.encodeStateAsUpdate(legacy))).toThrow(
        'Unsupported owned-text protocol'
      );
      expect(() => scope.assertHealthy()).toThrow('Unsupported owned-text protocol');
      scope.dispose();
      const reopened = new Y.Doc();
      try {
        Y.applyUpdate(reopened, Y.encodeStateAsUpdate(current));
        expect(() => attachOwnedTextAttributes(reopened, { protocol: 2 })).toThrow(
          'Unsupported owned-text protocol'
        );
      } finally {
        reopened.destroy();
      }
    } finally {
      scope.dispose();
      current.destroy();
      legacy.destroy();
    }
  });
