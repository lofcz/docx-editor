import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { propertyActivity, propertyRecordFixture } from '../../test/support/property-record';
import {
  labelEntry,
  labelPayload,
  labelSource,
  legacyLabelNode,
  writeLabel,
} from '../../test/support/payload-label';
import { withPayloadRecord } from './payload-opc-codec';

function paragraph(editor: DocxEditorInstance, text: string) {
  const surface = editor.surface!,
    id = surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id) === text);
  if (!id) throw new Error('Missing ' + text);
  return id;
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

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const equal of [false, true])
      for (const modern of [false, true])
        test(`scalar labels preserve whole values, peer fields/text, both actors' history, geometry, and reopening; ${clients}/${receiver}/equal=${equal}/modern=${modern}`, async () => {
          let pkg = propertyRecordFixture();
          if (modern) {
            const prepared = withPayloadRecord(pkg, labelPayload(pkg), 'Prepared label');
            if (!prepared.ok) throw new Error(prepared.reason);
            pkg = prepared.pkg;
          }
          const pair = await createTwoEditorFixture({
            document: writeOoxmlPackage(pkg),
            clientIds: clients,
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
              id = paragraph(bob, 'Body'),
              created = legacyLabelNode(labelSource(alice)),
              aliceLabel = '  Žák\tAlice\n😀  ',
              bobLabel = equal ? aliceLabel : '"Bob"\r\n尾 ',
              geometry = caretAt(alice.surface!.layout(), { paragraphId: id, offset: 0 });
            expect(geometry).not.toBeNull();
            pair.pause();
            writeLabel(alice, aliceLabel, 'duration', 15);
            expect(caretAt(alice.surface!.layout(), { paragraphId: id, offset: 0 })).toEqual(
              geometry
            );
            type(bob, id, 'Peer: ');
            writeLabel(bob, bobLabel, 'name', 'Peer activity');
            pair.reconnect(receiver);
            await pair.settle();
            const merged = labelEntry(labelSource(alice)).label;
            expect([aliceLabel, bobLabel]).toContain(merged);
            if (equal) expect(merged).toBe(aliceLabel);
            for (const editor of pair.editors) {
              expect(labelEntry(labelSource(editor)).label).toBe(merged);
              expect(labelPayload(labelSource(editor)).attributes).toMatchObject({
                duration: 15,
                name: 'Peer activity',
              });
              expect(legacyLabelNode(labelSource(editor))).toMatchObject({
                data: created.data,
                label: created.label,
              });
              expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
            }
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (row) => row.actorId === 'bob'
            );
            expect(remote).toBeDefined();
            expect(
              caretAt(alice.surface!.layout(), {
                paragraphId: remote!.head.nodeId,
                offset: remote!.head.offset,
              })
            ).not.toBeNull();
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(labelEntry(labelSource(editor)).label).toBe(bobLabel);
              expect(labelPayload(labelSource(editor)).attributes).toMatchObject({
                duration: 10,
                name: 'Peer activity',
              });
              expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
            }
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(labelEntry(labelSource(editor)).label).toBe(aliceLabel);
              expect(labelPayload(labelSource(editor)).attributes).toMatchObject({
                duration: 15,
                name: 'Old activity',
              });
              expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
            }
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            const final = labelEntry(labelSource(alice)).label;
            expect([aliceLabel, bobLabel]).toContain(final);
            for (const editor of pair.editors) {
              expect(labelEntry(labelSource(editor)).label).toBe(final);
              const read = readOoxmlPackage(new Uint8Array(await editor.save()));
              if (!read.ok) throw new Error(read.reason);
              expect(labelEntry(read.package).label).toBe(final);
              expect(labelPayload(read.package).attributes).toMatchObject({
                duration: 15,
                name: 'Peer activity',
              });
            }
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
            const snapshot = Y.encodeStateAsUpdate(pair.documents[receiver]!);
            const reopened = await createTwoEditorFixture({
              document: new Uint8Array(),
              roomSnapshot: snapshot,
            });
            try {
              await reopened.settle();
              for (const editor of reopened.editors) {
                expect(labelEntry(labelSource(editor)).label).toBe(final);
                expect(labelPayload(labelSource(editor)).attributes).toMatchObject({
                  duration: 15,
                  name: 'Peer activity',
                });
                expect(
                  paragraphTextOf(editor.surface!.session.part(), paragraph(editor, 'Peer: Body'))
                ).toBe('Peer: Body');
              }
            } finally {
              reopened.destroy();
            }
          } finally {
            pair.destroy();
          }
        }, 30000);

for (const receiver of [0, 1] as const)
  test(`owner deletion preserves peer scalar label and fields and restores text inside the deleted owner; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(propertyRecordFixture()),
      clientIds: [41, 73],
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = propertyActivity(labelSource(alice)),
        id = paragraph(bob, 'Body');
      pair.pause();
      expect(
        alice.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).committed
      ).toBe(true);
      type(bob, id, 'Peer: ');
      writeLabel(bob, '  Peer label\n尾  ', 'name', 'Peer activity');
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(labelSource(editor))).toBeUndefined();
        expect(labelEntry(labelSource(editor)).label).toBe('  Peer label\n尾  ');
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(labelSource(editor))).toBeDefined();
        expect(labelEntry(labelSource(editor)).label).toBe('  Peer label\n尾  ');
        expect(labelPayload(labelSource(editor)).attributes.name).toBe('Peer activity');
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body');
        expect(caretAt(editor.surface!.layout(), { paragraphId: id, offset: 0 })).not.toBeNull();
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(labelEntry(labelSource(editor)).label).toBe('  Peer label\n尾  ');
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  }, 30000);
