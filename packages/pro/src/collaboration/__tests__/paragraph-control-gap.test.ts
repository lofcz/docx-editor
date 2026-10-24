/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { diffTokens, type Token } from '../document/paragraph-text-diff';
import { concurrentRevisionTail } from '../document/concurrent-revision-tail';
import { textIdentities } from '../document/paragraph-text-identity';
import { DEFAULT_DOCUMENT_LIMITS } from '../document/limits';
import { asLogicalId } from '../document/identity';
import { keepFormattingMarkers } from '../document/yjs-items';

function chars(value: string, holder: string): Token[] {
  return [...value].map((insert) => ({
    key: 'c' + insert,
    strong: 'c' + insert + '\0' + holder,
    insert,
    attributes: { t: holder },
  }));
}
test('complete native gap text keeps its identities between replaced atom labels', () => {
  const before = [
    ...chars('Ada', 'old-mention'),
    ...chars(' ', 'gap'),
    ...chars('Paper', 'old-source'),
  ];
  const after = [
    ...chars('Grace', 'new-mention'),
    ...chars(' ', 'gap'),
    ...chars('Updated', 'new-source'),
  ];
  expect(diffTokens(before, after).filter((step) => step.op === 'eq')).toEqual([
    { op: 'eq', before: 3, after: 5 },
  ]);
});
test('equal text in a newly created gap grants no native identity', () => {
  const before = [
    ...chars('Ada', 'old-mention'),
    ...chars(' ', 'old-gap'),
    ...chars('Paper', 'old-source'),
  ];
  const after = [
    ...chars('Grace', 'new-mention'),
    ...chars(' ', 'new-gap'),
    ...chars('Updated', 'new-source'),
  ];
  expect(diffTokens(before, after).some((step) => step.op === 'eq' && step.before === 3)).toBe(
    false
  );
});

const tracked = {
  r: '{"i":"native-run"}',
  t: '{"i":"native-text"}',
  'v:review': JSON.stringify({
    e: {
      i: 'review',
      l: 'ins',
      k: 'revisionInsert',
      a: [
        ['id', '7'],
        ['author', 'AI'],
      ],
    },
    d: 1,
  }),
};
for (const causal of [true, false])
  for (const gc of [true, false])
    for (const boundary of [false, true])
      test(`terminal revision ordering respects observed right boundaries: causal=${causal}, gc=${gc}, boundary=${boundary}`, () => {
        const alice = new Y.Doc({ gc }),
          bob = new Y.Doc({ gc }),
          cold = new Y.Doc({ gc });
        alice.clientID = 73;
        bob.clientID = 41;
        const stops = [
          keepFormattingMarkers(alice),
          keepFormattingMarkers(bob),
          keepFormattingMarkers(cold),
        ];
        try {
          const left = alice.getText('paragraph');
          const base = boundary ? 1 : 0;
          if (boundary) left.insertEmbed(0, { p: 1 });
          left.insert(base, 'A ');
          if (causal) left.insert(base + 2, 'Paper', tracked);
          Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
          if (!causal) left.insert(base + 2, 'Paper', tracked);
          bob.getText('paragraph').insert(base + 2, 'Peer', {});
          Y.applyUpdate(alice, Y.encodeStateAsUpdate(bob));
          Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
          Y.applyUpdate(cold, Y.encodeStateAsUpdate(alice));
          for (const doc of [alice, bob, cold]) {
            const text = doc.getText('paragraph');
            expect(text.toString()).toBe('A PeerPaper');
            const shown = {
              identities: textIdentities(text),
              hidden: new Set<number>(),
              incoming: [],
            };
            const result = concurrentRevisionTail(text, shown, asLogicalId('paragraph'), {
              ...DEFAULT_DOCUMENT_LIMITS,
              maxTextLength: 11,
            });
            expect(result.incoming).toHaveLength(causal ? 0 : 1);
            expect(
              concurrentRevisionTail(text, shown, asLogicalId('paragraph'), {
                ...DEFAULT_DOCUMENT_LIMITS,
                maxTextLength: 10,
              }).incoming
            ).toEqual([]);
            expect([...result.hidden]).toEqual(
              causal ? [] : [6, 7, 8, 9, 10].map((offset) => offset + base)
            );
          }
        } finally {
          stops.forEach((stop) => stop());
          alice.destroy();
          bob.destroy();
          cold.destroy();
        }
      });

test('collected formatting clocks cannot change terminal append ordering', () => {
  const alice = new Y.Doc({ gc: false }),
    bob = new Y.Doc({ gc: false }),
    cold = new Y.Doc({ gc: true });
  alice.clientID = 73;
  bob.clientID = 41;
  const stops = [
    keepFormattingMarkers(alice),
    keepFormattingMarkers(bob),
    keepFormattingMarkers(cold),
  ];
  try {
    const text = alice.getText('paragraph');
    text.insert(0, 'A ');
    Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
    text.insert(2, 'Paper', tracked);
    bob.getText('paragraph').insert(2, 'Peer', {});
    Y.applyUpdate(alice, Y.encodeStateAsUpdate(bob));
    let marker = text._start;
    while (marker && !(marker.content instanceof Y.ContentFormat && marker.content.key === 't'))
      marker = marker.right;
    if (!marker) throw Error('Missing native formatting clock.');
    const deleted = marker;
    Y.transact(alice, (transaction) => deleted.delete(transaction));
    Y.applyUpdate(cold, Y.encodeStateAsUpdate(alice));
    const incoming = [alice, cold].map((doc) => {
      const current = doc.getText('paragraph');
      expect(current.toString()).toBe('A PeerPaper');
      const shown = {
        identities: textIdentities(current),
        hidden: new Set<number>(),
        incoming: [],
      };
      return concurrentRevisionTail(
        current,
        shown,
        asLogicalId('paragraph'),
        DEFAULT_DOCUMENT_LIMITS
      ).incoming;
    });
    expect(incoming).toEqual([[], []]);
  } finally {
    stops.forEach((stop) => stop());
    alice.destroy();
    bob.destroy();
    cold.destroy();
  }
});
