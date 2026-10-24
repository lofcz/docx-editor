import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { EditorCollaborationSession } from '@docx-editor.dev/core/collaboration';
import { attachOwnedReviewHistory } from './review-history';
import { reviewGroupHistoryRefusal, REVIEW_HISTORY_CAPABILITY } from '../review/group-history';
import { OWNED_TEXT_OPERATIONS } from './ownership-codec';
import { REVIEW_FLAG_TARGETS } from './review-flag-codec';

class Transport {
  readonly structuralTextAttributeKeys = Object.freeze([]);
  #ready = true;
  constructor(readonly documentId = 'room') {}
  get attached() {
    return this.#ready;
  }
  status() {
    return this.#ready ? 'ready' : 'destroyed';
  }
  close() {
    this.#ready = false;
  }
}
const session = (source: unknown = new Transport()) =>
  source as unknown as EditorCollaborationSession;

test('owned review sessions retain native method receivers, stable functions, frozen transports, and repeat binding', () => {
  const doc = new Y.Doc(),
    scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    }),
    transport = Object.freeze(new Transport());
  try {
    const native = session(transport),
      bound = scope.bindSession(native);
    expect(bound).not.toBe(native);
    expect(bound.documentId).toBe('room');
    expect(bound.attached).toBe(true);
    expect(bound.status()).toBe('ready');
    expect(bound.status).toBe(bound.status);
    expect(scope.bindSession(native)).toBe(bound);
    expect(scope.bindSession(bound)).toBe(bound);
    expect(reviewGroupHistoryRefusal(bound)).toBeNull();
    expect(reviewGroupHistoryRefusal(null)).toBeNull();
    expect(
      (native as unknown as Record<symbol, unknown>)[REVIEW_HISTORY_CAPABILITY]
    ).toBeUndefined();
    transport.close();
    expect(bound.status()).toBe('destroyed');
    expect(bound.attached).toBe(false);
    const frozen = Object.freeze({
      documentId: 'room',
      structuralTextAttributeKeys: Object.freeze([]),
      status() {
        return this.documentId === 'room' ? 'ready' : 'error';
      },
    }) as unknown as EditorCollaborationSession;
    expect(scope.bindSession(frozen).status()).toBe('ready');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('missing, foreign, disposed, and destroyed history scopes refuse with explicit reasons', () => {
  const doc = new Y.Doc(),
    otherDoc = new Y.Doc(),
    scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    }),
    other = attachOwnedReviewHistory(otherDoc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    });
  try {
    expect(reviewGroupHistoryRefusal(session())).toContain('before room bootstrap');
    expect(() => scope.bindSession(session(new Transport('other')))).toThrow('document identity');
    const bound = scope.bindSession(session());
    expect(() => other.bindSession(bound)).toThrow('another review history scope');
    expect(() =>
      attachOwnedReviewHistory(doc, {
        documentId: 'room',
        bootstrap: 'create',
        structuralTextAttributeKeys: [],
      })
    ).toThrow('already attached');
    scope.dispose();
    scope.dispose();
    expect(reviewGroupHistoryRefusal(bound)).toContain('no longer active');
    expect(() => scope.bindSession(session())).toThrow('no longer active');
    const boundOther = other.bindSession(session());
    otherDoc.destroy();
    expect(reviewGroupHistoryRefusal(boundOther)).toContain('no longer active');
  } finally {
    scope.dispose();
    other.dispose();
    doc.destroy();
    otherDoc.destroy();
  }
});

test('configuration validates before attachment, and legacy rooms require an explicit upgrade before authoring', () => {
  const doc = new Y.Doc();
  try {
    for (const documentId of ['', ' ', 'x\u0000y', 'x'.repeat(4097)])
      expect(() =>
        attachOwnedReviewHistory(doc, {
          documentId,
          bootstrap: 'create',
          structuralTextAttributeKeys: [],
        })
      ).toThrow('document ID');
    const scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    });
    scope.dispose();
    doc.getText('legacy').insert(0, 'Native content');
    expect(() =>
      attachOwnedReviewHistory(doc, {
        documentId: 'room',
        bootstrap: 'create',
        structuralTextAttributeKeys: [],
      })
    ).toThrow('before legacy room content');
    expect(doc.getText('legacy').toString()).toBe('Native content');
  } finally {
    doc.destroy();
  }
});

test('poisoned ownership data refuses future review operations without replacing the protocol error', () => {
  const doc = new Y.Doc(),
    errors: Error[] = [],
    scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
      onError: (error) => errors.push(error),
    }),
    bound = scope.bindSession(session());
  try {
    expect(() => doc.getMap(OWNED_TEXT_OPERATIONS).set('protocol', 'unsupported')).toThrow(
      'Unsupported owned-text protocol'
    );
    expect(errors).toHaveLength(1);
    expect(() => scope.assertHealthy()).toThrow('Unsupported owned-text protocol');
    expect(reviewGroupHistoryRefusal(bound)).toContain('Unsupported owned-text protocol');
  } finally {
    scope.dispose();
    doc.destroy();
  }
});

test('review flag protocol failures release partial attachment and refuse stale scope capabilities', () => {
  const doc = new Y.Doc();
  try {
    doc.getMap(OWNED_TEXT_OPERATIONS).set('protocol', '2');
    doc.getMap(REVIEW_FLAG_TARGETS).set('protocol', '1');
    expect(() =>
      attachOwnedReviewHistory(doc, {
        documentId: 'room',
        bootstrap: 'create',
        structuralTextAttributeKeys: [],
      })
    ).toThrow('Unsupported review flag protocol');
    doc.getMap(REVIEW_FLAG_TARGETS).delete('protocol');
    const scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    });
    try {
      expect(reviewGroupHistoryRefusal(scope.bindSession(session()))).toBeNull();
      for (const version of [1, 2, 3, 4]) {
        const old = Object.assign(new Transport(), {
          [Symbol.for(`sciobot.docx.review-history.${version}`)]: () => {},
        });
        expect(reviewGroupHistoryRefusal(session(old))).toContain('before room bootstrap');
      }
    } finally {
      scope.dispose();
    }
  } finally {
    doc.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test('mixed legacy review headers refuse live updates and reopened snapshots; ' + clients, () => {
    const doc = new Y.Doc(),
      legacy = new Y.Doc();
    doc.clientID = clients[0];
    legacy.clientID = clients[1];
    const scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    });
    legacy.getMap(OWNED_TEXT_OPERATIONS).set('protocol', '1');
    legacy.getMap(REVIEW_FLAG_TARGETS).set('protocol', '1');
    try {
      expect(() => Y.applyUpdate(doc, Y.encodeStateAsUpdate(legacy), 'relay')).toThrow();
      expect(() => scope.assertHealthy()).toThrow();
      scope.dispose();
      const reopened = new Y.Doc();
      try {
        Y.applyUpdate(reopened, Y.encodeStateAsUpdate(doc));
        expect(() =>
          attachOwnedReviewHistory(reopened, {
            documentId: 'room',
            bootstrap: 'create',
            structuralTextAttributeKeys: [],
          })
        ).toThrow('protocol');
      } finally {
        reopened.destroy();
      }
    } finally {
      scope.dispose();
      doc.destroy();
      legacy.destroy();
    }
  });

test('a bound session shares its native undo authority without creating a second history owner', async () => {
  const { registerUndoHistoryPosition } =
    await import('@docx-editor.dev/core/collaboration/replication');
  const { undoHistoryPosition } =
    await import('../../../core/src/store/store/undo-history-position');
  const doc = new Y.Doc(),
    scope = attachOwnedReviewHistory(doc, {
      documentId: 'room',
      bootstrap: 'create',
      structuralTextAttributeKeys: [],
    }),
    native = session(),
    key = {};
  let splits = 0;
  registerUndoHistoryPosition(native, {
    current: () => key,
    split: () => {
      splits++;
    },
  });
  try {
    const bound = scope.bindSession(native);
    expect(undoHistoryPosition(bound)).toBe(undoHistoryPosition(native));
    expect(undoHistoryPosition(bound)!.current()).toBe(key);
    undoHistoryPosition(bound)!.split();
    expect(splits).toBe(1);
  } finally {
    scope.dispose();
    doc.destroy();
  }
});
