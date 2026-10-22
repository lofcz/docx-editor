import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, writeOoxmlPackage, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  propertyActivity,
  propertyRecord,
  propertyRecordFixture,
  propertyRecordRevision,
} from '../../test/support/property-record';
import { readEditorRevision } from '../projection/revision-cache';
import type { PropertyReviewRecord } from './property-record-schema';
import {
  readPropertyReviewStore,
  withPropertyReviewRecord,
  withoutPropertyReviewRecord,
  type PropertyStoreWrite,
} from './property-record-store';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

// Carrier-only transactions exercise existing native package edits. R12 owns tracked fallback writes.
function mutate(editor: DocxEditorInstance, change: (pkg: OoxmlPackage) => PropertyStoreWrite) {
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
function records(editor: DocxEditorInstance) {
  const source = readPropertyReviewStore(editor.surface!.session.currentPackage());
  if (source.status === 'invalid') throw new Error(source.reason);
  return source.status === 'ready' ? [...source.records.keys()].sort() : [];
}
async function views(editor: DocxEditorInstance) {
  const token = readEditorRevision(
      editor,
      propertyRecordRevision.documentId,
      propertyRecordRevision.epoch
    )!,
    source = editor.surface!.session,
    [committed, working] = await Promise.all([
      readCommittedView(source, token),
      readWorkingView(source, token),
    ]);
  if (!committed.success) throw new Error(committed.error);
  if (!working.success) throw new Error(working.error);
  return { committed: committed.view.projection.text, working: working.view.projection.text };
}
function type(editor: DocxEditorInstance, id: string, text: string) {
  const surface = editor.surface!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`independent property records retain offline peer proposals and text through history, removal, and save; clients ${clientIds.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(propertyRecordFixture()),
        clientIds,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice
            .surface!.session.paragraphIds()
            .find(
              (paragraphId) =>
                paragraphTextOf(alice.surface!.session.part(), paragraphId) === 'Body'
            )!,
          first = propertyRecord('alice'),
          second: PropertyReviewRecord = {
            ...propertyRecord('bob'),
            author: 'Bob',
            blocks: [
              {
                blockId: 'activity',
                type: 'lesson_activity',
                changes: [
                  {
                    name: 'name',
                    before: { present: true, value: 'Old activity' },
                    after: { present: true, value: 'New activity' },
                  },
                ],
              },
            ],
          };
        bob.surface!.setSelection({
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 0 },
        });
        await pair.settle();
        const geometry = caretAt(alice.surface!.layout(), { paragraphId: id, offset: 0 });
        expect(geometry).not.toBeNull();
        pair.pause();
        mutate(alice, (pkg) => withPropertyReviewRecord(pkg, first));
        expect(caretAt(alice.surface!.layout(), { paragraphId: id, offset: 0 })).toEqual(geometry);
        mutate(bob, (pkg) => withPropertyReviewRecord(pkg, second));
        type(bob, id, 'Bob: ');
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(records(editor)).toEqual(['alice', 'bob']);
          const read = await views(editor);
          expect(read.committed).toContain('name="Old activity" duration="10"');
          expect(read.working).toContain('name="New activity" duration="15"');
          expect(read.working).toContain('Bob: Body');
          expect(
            propertyActivity(editor.surface!.session.currentPackage()).payload.attributes.duration
          ).toBe(10);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Bob: Body');
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote).toBeDefined();
        expect(
          caretAt(alice.surface!.layout(), {
            paragraphId: remote!.head.nodeId,
            offset: remote!.head.offset,
          })
        ).not.toBeNull();
        const pending = await pair.savedDigests();
        expect(pending[0]).toEqual(pending[1]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(records(editor)).toEqual(['bob']);
          expect((await views(editor)).working).toContain('name="New activity" duration="10"');
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Bob: Body');
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(records(editor)).toEqual(['alice', 'bob']);
        pair.pause();
        mutate(alice, (pkg) => withoutPropertyReviewRecord(pkg, first.id));
        type(bob, id, 'Still ');
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(records(editor)).toEqual(['bob']);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Still Bob: Body');
        }
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(records(editor)).toEqual(['alice', 'bob']);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Still Bob: Body');
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(records(editor)).toEqual(['bob']);
        const final = await pair.savedDigests();
        expect(final[0]).toEqual(final[1]);
      } finally {
        pair.destroy();
      }
    }, 30_000);

for (const receiver of [0, 1] as const)
  test(`owner deletion does not lose peer proposal metadata or recreate native content; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(propertyRecordFixture()),
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = propertyActivity(alice.surface!.session.currentPackage()),
        record = propertyRecord();
      mutate(alice, (pkg) => withPropertyReviewRecord(pkg, record));
      await pair.settle();
      pair.pause();
      const deleted = alice.surface!.applyAutomationOps(
        () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
        { kind: 'body' },
        undefined,
        false,
        true
      );
      expect(deleted.rejected).toBe(false);
      mutate(bob, (pkg) =>
        withPropertyReviewRecord(pkg, { ...record, extension: { peer: 'preserve' } })
      );
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage(),
          data = readPropertyReviewStore(pkg);
        expect(propertyActivity(pkg)).toBeUndefined();
        expect(data.status).toBe('ready');
        if (data.status !== 'ready') throw new Error('Missing carrier');
        expect(data.records.get(record.id)!.decoded).toMatchObject({
          status: 'supported',
          value: { extension: { peer: 'preserve' } },
        });
        const token = readEditorRevision(
          editor,
          propertyRecordRevision.documentId,
          propertyRecordRevision.epoch
        )!;
        expect(readStoredPropertyProposals(pkg, token)).toMatchObject({
          success: false,
          code: 'invalid-document',
        });
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(editor.surface!.session.currentPackage())).toBeDefined();
        expect((await views(editor)).working).toContain('duration="15"');
      }
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);
