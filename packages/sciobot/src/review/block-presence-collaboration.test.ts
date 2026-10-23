import { docxFixture } from '../../test/support/document';
import { sourceControl, sourceParagraph } from '../../test/support/target-source-fixture';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { presenceAttribution, presencePackage } from '../../test/support/block-presence';
import { createBlockPresenceRecord, bindBlockPresenceRecord } from './block-presence-plan';
import { projectBlockPresence } from './block-presence-projection';
import { withBlockPresenceRecord } from './block-presence-store';
import { readPropertyReviewStore } from './property-record-store';
import { decodeBlockPresenceRecord, encodeBlockPresenceRecord } from './block-presence-schema';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`presence metadata and pure projection retain real room history and reopen; clients=${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(presencePackage()),
      documentId: 'presence',
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      clientIds: clients,
    });
    try {
      const [alice, bob] = pair.editors;
      const before = alice!.surface!.session.currentPackage();
      const prepared = createBlockPresenceRecord(
        presencePackage(false),
        before,
        [{ blockId: 'outer', before: { present: false }, after: { present: true } }],
        presenceAttribution
      );
      if (!prepared.success) throw Error(prepared.error);
      const record = prepared.record,
        raw = encodeBlockPresenceRecord(record),
        native = canonicalOoxmlFingerprint(alice!.surface!.session.part());
      const write = alice!.surface!.applyAutomationOps(
        () => [],
        { kind: 'body' },
        [
          (working) => {
            const stored = withBlockPresenceRecord(working, record);
            if (!stored.success) throw Error(stored.error);
            return stored.pkg;
          },
        ],
        true,
        true,
        'after'
      );
      if (write.rejected) throw Error(JSON.stringify(write));
      expect(write.rejected).toBe(false);
      expect(write.committed).toBe(true);
      await pair.settle();
      const check = () => {
        for (const editor of pair.editors) {
          const pkg = editor.surface!.session.currentPackage(),
            stored = readPropertyReviewStore(pkg);
          if (stored.status !== 'ready') throw Error('Missing store');
          expect(stored.records.get(record.id)?.raw).toBe(raw);
          expect(decodeBlockPresenceRecord(stored.records.get(record.id)!.raw).status).toBe(
            'supported'
          );
          expect(bindBlockPresenceRecord(pkg, record).success).toBe(true);
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(native);
        }
      };
      check();
      const roomBytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      const inspected = alice!.surface!.applyAutomationOps(
        () => [],
        { kind: 'body' },
        [
          (working) => {
            const projected = projectBlockPresence(working, record, 'committed');
            if (!projected.success) throw Error(projected.error);
            expect(projected.pkg.parts.get(projected.pkg.mainDocumentPart)).not.toBe(
              working.parts.get(working.mainDocumentPart)
            );
            return working;
          },
        ],
        true,
        true,
        'after'
      );
      expect(inspected.rejected).toBe(false);
      await pair.settle();
      pair.documents.forEach((doc, index) =>
        expect(Y.encodeStateAsUpdate(doc)).toEqual(roomBytes[index]!)
      );
      check();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const store = readPropertyReviewStore(editor.surface!.session.currentPackage());
        expect(store.status === 'ready' && store.records.has(record.id)).toBe(false);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(native);
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check();
      const room = await createTwoEditorFixture({
        document: new Uint8Array(await bob!.save()),
        documentId: 'presence',
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
      });
      try {
        await room.settle();
        for (const editor of room.editors)
          expect(
            bindBlockPresenceRecord(editor.surface!.session.currentPackage(), record).success
          ).toBe(true);
        room.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        room.destroy();
      }
      pair.ownership.forEach((handle) => handle.assertHealthy());
    } finally {
      pair.destroy();
    }
  }, 30000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`independent presence records retain outside peer text and independent history; ${clients}/${receiver}`, async () => {
      const read = readOoxmlPackage(
        docxFixture(
          sourceControl('alpha', sourceParagraph('A')) +
            sourceControl('beta', sourceParagraph('B')) +
            sourceParagraph('Outside')
        )
      );
      if (!read.ok) throw Error(read.reason);
      let pkg = read.package;
      for (const id of ['alpha', 'beta']) {
        const stored = withPayloadRecord(
          pkg,
          { id, type: 'custom_atomic', version: 1, attributes: { keep: id } },
          id
        );
        if (!stored.ok) throw Error(stored.reason);
        pkg = stored.pkg;
      }
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(pkg),
        documentId: 'presence',
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const outside = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Outside')!;
        const expected = applyTreeOp(alice!.surface!.session.part(), {
          op: 'insertText',
          paragraphId: outside,
          offset: 7,
          text: ' Peer',
        });
        if (!expected.ok) throw Error(expected.reason);
        pair.pause();
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: outside, offset: 7, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        const records = pair.editors.map((editor, index) => {
          const prepared = createBlockPresenceRecord(
            presencePackage(false),
            editor.surface!.session.currentPackage(),
            [
              {
                blockId: index ? 'beta' : 'alpha',
                before: { present: false },
                after: { present: true },
              },
            ],
            { ...presenceAttribution, id: index ? 'presence-bob' : 'presence-alice' }
          );
          if (!prepared.success) throw Error(prepared.error);
          return prepared.record;
        });
        pair.editors.forEach((editor, index) => {
          const write = editor.surface!.applyAutomationOps(
            () => [],
            { kind: 'body' },
            [
              (working) => {
                const stored = withBlockPresenceRecord(working, records[index]!);
                if (!stored.success) throw Error(stored.error);
                return stored.pkg;
              },
            ],
            true,
            true,
            'after'
          );
          if (write.rejected) throw Error(JSON.stringify(write));
          expect(write.committed).toBe(true);
        });
        pair.reconnect(receiver);
        await pair.settle();
        const check = (ids: readonly string[]) => {
          for (const editor of pair.editors) {
            const current = editor.surface!.session.currentPackage(),
              store = readPropertyReviewStore(current);
            if (store.status !== 'ready') throw Error('Missing store');
            expect([...store.records.keys()].sort()).toEqual([...ids].sort());
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(expected.part)
            );
            for (const record of records.filter((record) => ids.includes(record.id))) {
              expect(store.records.get(record.id)?.raw).toBe(encodeBlockPresenceRecord(record));
              expect(bindBlockPresenceRecord(current, record).success).toBe(true);
            }
          }
        };
        check(records.map((record) => record.id));
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(['presence-bob']);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(records.map((record) => record.id));
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(['presence-alice']);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(records.map((record) => record.id));
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
