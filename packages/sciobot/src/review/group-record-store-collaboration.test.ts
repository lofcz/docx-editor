import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlsIn,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import type { ReviewGroupRecord } from './group-record-schema';
import {
  readReviewGroupStore,
  withReviewGroupRecord,
  withoutReviewGroupRecord,
  type ReviewGroupStoreWrite,
} from './group-record-store';

const record = (id: string, author: string): ReviewGroupRecord => ({
  schemaVersion: 1,
  id,
  members: [
    {
      partName: '/word/document.xml',
      revision: { id: author === 'Alice' ? '41' : '73', author, date: '2026-10-20T10:00:00.000Z' },
    },
  ],
});
const documentBytes = () =>
  docxFixture(
    '<w:sdt><w:sdtPr><w:tag w:val="ordinary-owner"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Owned</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
  );
function source(editor: DocxEditorInstance) {
  return editor.surface!.session.currentPackage();
}
function store(editor: DocxEditorInstance) {
  const value = readReviewGroupStore(source(editor));
  if (value.status !== 'ready') throw new Error('Missing initialized review group store');
  return value;
}
const ids = (editor: DocxEditorInstance) => [...store(editor).records.keys()].sort();
function mutate(editor: DocxEditorInstance, change: (pkg: OoxmlPackage) => ReviewGroupStoreWrite) {
  const result = editor.surface!.applyAutomationOps(
    () => [],
    { kind: 'body' },
    [
      (pkg) => {
        const next = change(pkg);
        if (!next.ok) throw new Error(next.reason);
        return next.pkg;
      },
    ],
    false,
    true
  );
  expect(result.rejected).toBe(false);
  expect(result.committed).toBe(true);
}
function type(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const surface = editor.surface!;
  surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`preinitialized group records preserve peer creation, removal, text, history, geometry, and saved convergence; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ document: documentBytes(), clientIds: clients });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          tailId = alice
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Tail')!,
          initialPart = store(alice).part,
          body = canonicalOoxmlFingerprint(alice.surface!.session.part()),
          geometry = caretAt(alice.surface!.layout(), { paragraphId: tailId, offset: 0 });
        expect(geometry).not.toBeNull();
        for (const editor of pair.editors) {
          expect(ids(editor)).toEqual([]);
          expect(editor.surface!.session.canUndo()).toBe(false);
          expect(store(editor).part.partName).toBe(initialPart.partName);
        }
        pair.pause();
        mutate(alice, (pkg) => withReviewGroupRecord(pkg, record('rg_alice', 'Alice')));
        expect(canonicalOoxmlFingerprint(alice.surface!.session.part())).toBe(body);
        expect(caretAt(alice.surface!.layout(), { paragraphId: tailId, offset: 0 })).toEqual(
          geometry
        );
        mutate(bob, (pkg) => withReviewGroupRecord(pkg, record('rg_bob', 'Bob')));
        type(bob, tailId, 'Bob: ');
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(ids(editor)).toEqual(['rg_alice', 'rg_bob']);
          expect(store(editor).part).toEqual(initialPart);
          expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Tail');
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(tailId);
        expect(
          caretAt(alice.surface!.layout(), { paragraphId: tailId, offset: remote!.head.offset })
        ).not.toBeNull();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(ids(editor)).toEqual(['rg_bob']);
          expect(store(editor).part).toEqual(initialPart);
          expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Bob: Tail');
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(ids(editor)).toEqual(['rg_alice', 'rg_bob']);
        pair.pause();
        mutate(alice, (pkg) => withoutReviewGroupRecord(pkg, 'rg_bob'));
        mutate(bob, (pkg) => withReviewGroupRecord(pkg, record('rg_later', 'Bob')));
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(ids(editor)).toEqual(['rg_alice', 'rg_later']);
          expect([...store(editor).removedIds]).toEqual(['rg_bob']);
        }
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(ids(editor)).toEqual(['rg_alice', 'rg_bob', 'rg_later']);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(ids(editor)).toEqual(['rg_alice', 'rg_later']);
          expect(withReviewGroupRecord(source(editor), record('rg_bob', 'Bob'))).toMatchObject({
            ok: false,
            reason: 'review-group-id-reserved',
          });
          const read = readOoxmlPackage(writeOoxmlPackage(source(editor)));
          if (!read.ok) throw new Error(read.reason);
          const reopened = readReviewGroupStore(read.package);
          if (reopened.status !== 'ready') throw new Error('Missing saved group store');
          expect([...reopened.records.keys()].sort()).toEqual(['rg_alice', 'rg_later']);
          expect([...reopened.removedIds]).toEqual(['rg_bob']);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    }, 30_000);

for (const receiver of [0, 1] as const)
  test(`owner deletion during a peer record write keeps the initialized part and recovers native content on undo; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({ document: documentBytes(), clientIds: [41, 73] });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = contentControlsIn(alice.surface!.session.part().root)[0]!.node,
        tailId = alice
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Tail')!,
        initialPart = store(alice).part;
      pair.pause();
      expect(
        alice.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId: owner.id, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).rejected
      ).toBe(false);
      mutate(bob, (pkg) => withReviewGroupRecord(pkg, record('rg_peer', 'Bob')));
      type(bob, tailId, 'Peer: ');
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(contentControlsIn(editor.surface!.session.part().root)).toHaveLength(0);
        expect(ids(editor)).toEqual(['rg_peer']);
        expect(store(editor).part).toEqual(initialPart);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(contentControlsIn(editor.surface!.session.part().root)).toHaveLength(1);
        expect(ids(editor)).toEqual(['rg_peer']);
        expect(store(editor).part).toEqual(initialPart);
        expect(paragraphTextOf(editor.surface!.session.part(), tailId)).toBe('Peer: Tail');
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);
