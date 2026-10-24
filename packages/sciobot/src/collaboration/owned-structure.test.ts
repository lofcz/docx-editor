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
    if (oldVersion === 2) old.setStructuralAttributes([]);
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

test('complete history refuses a provider without advertised structural keys before binding writes', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    });
  try {
    const before = Y.encodeStateAsUpdate(doc);
    expect(() => scope.bindSession({ documentId: 'room' } as EditorCollaborationSession)).toThrow(
      'provider structural'
    );
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
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
    scope.setStructuralAttributes([]);
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

test('protocol 2 refuses text writes until structural policy binds, before changing native text', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedTextAttributes(doc, { protocol: 2 }),
    text = doc.getText('p');
  try {
    const before = Y.encodeStateAsUpdate(doc);
    expect(() => text.insert(0, 'X')).toThrow('Bind a structural');
    expect(() => text.applyDelta([{ insert: 'X' }])).toThrow('Bind a structural');
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    scope.setStructuralAttributes([]);
    text.insert(0, 'X');
    expect(text.toString()).toBe('X');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('invalid complete-history structural keys refuse before room bootstrap writes', () => {
  for (const keys of [['run', 'run'], ['__proto__'], ['']]) {
    const doc = new Y.Doc(),
      before = Y.encodeStateAsUpdate(doc);
    try {
      expect(() =>
        attachOwnedReviewHistory(doc, {
          documentId: 'room',
          bootstrap: 'create',
          structuralTextAttributeKeys: keys,
        })
      ).toThrow('Invalid structural');
      expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
    } finally {
      doc.destroy();
    }
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`conflicting policy advertisements remain refused after reopening: ${clients}`, () => {
    const docs = [new Y.Doc(), new Y.Doc()];
    docs.forEach((doc, i) => {
      doc.clientID = clients[i]!;
    });
    const scopes = docs.map((doc) => attachOwnedTextAttributes(doc, { protocol: 2 }));
    scopes[0]!.setStructuralAttributes(['run']);
    scopes[1]!.setStructuralAttributes(['other']);
    try {
      expect(() => Y.applyUpdate(docs[0]!, Y.encodeStateAsUpdate(docs[1]!))).toThrow(
        'Incompatible structural'
      );
      expect(() => scopes[0]!.assertHealthy()).toThrow('Incompatible structural');
      const reopened = new Y.Doc();
      try {
        Y.applyUpdate(reopened, Y.encodeStateAsUpdate(docs[0]!));
        expect(() => attachOwnedTextAttributes(reopened, { protocol: 2 })).toThrow(
          'Incompatible structural'
        );
      } finally {
        reopened.destroy();
      }
    } finally {
      scopes.forEach((scope) => scope.dispose());
      docs.forEach((doc) => doc.destroy());
    }
  });

test('a joining scope makes no metadata writes and waits for complete bootstrap dependencies', () => {
  const creator = new Y.Doc(),
    joining = new Y.Doc(),
    errors: Error[] = [];
  const a = attachOwnedTextAttributes(creator, { protocol: 2 });
  a.setStructuralAttributes(['node']);
  creator.getText('p').insert(0, 'AB', { node: 'run', bold: false });
  const seed = Y.encodeStateAsUpdate(creator),
    vector = Y.encodeStateVector(creator);
  creator.getText('p').insert(2, 'X');
  const later = Y.encodeStateAsUpdate(creator, vector);
  const before = Y.encodeStateAsUpdate(joining);
  const b = attachOwnedTextAttributes(joining, {
    protocol: 2,
    bootstrap: 'join',
    onError: (error) => errors.push(error),
  });
  try {
    b.setStructuralAttributes(['node']);
    expect(Y.encodeStateAsUpdate(joining)).toEqual(before);
    expect(() => Y.applyUpdate(joining, later)).not.toThrow();
    expect(errors).toEqual([]);
    Y.applyUpdate(joining, seed);
    b.assertHealthy();
    expect(joining.getText('p').toString()).toBe('ABX');
    expect(Y.encodeStateVector(joining)).toEqual(Y.encodeStateVector(creator));
    expect(errors).toEqual([]);
  } finally {
    a.dispose();
    b.dispose();
    creator.destroy();
    joining.destroy();
  }
});
