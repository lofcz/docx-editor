import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { seedReviewFieldCapacity } from '../../test/support/review-field-capacity';
import { groups } from '../../test/support/review-groups';
import { CELL_CONTEXT } from '../../test/support/table-review';
import { REVIEW_FIELD_TARGETS, REVIEW_FIELD_INTENTS } from '../collaboration/review-field-codec';
import { readEditorRevision } from '../projection/revision-cache';
import { acceptReviewGroup } from './accept-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const p = (text: string) => '<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>';
const document = () =>
  docxFixture(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr><w:tc>' +
      p('A') +
      '</w:tc><w:tc><w:tcPr><w:cellDel w:id="9" w:author="Teacher"/>' +
      '</w:tcPr>' +
      p('B') +
      '</w:tc></w:tr></w:tbl>' +
      p('Tail')
  );
for (const kind of ['targets', 'intents'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(
        'native field ' +
          kind +
          ' capacity refuses complete geometry writes before mutation; ' +
          clients +
          '/' +
          receiver,
        async () => {
          const pair = await createTwoEditorFixture({
            document: document(),
            ownedAttributes: true,
            clientIds: clients,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          try {
            const alice = pair.editors[0]!,
              bob = pair.editors[1]!,
              surface = alice.surface!,
              tail = surface.session
                .paragraphIds()
                .find((id) => paragraphTextOf(surface.session.part(), id) === 'Tail')!;
            surface.setSelection({
              anchor: { paragraphId: tail, offset: 4 },
              head: { paragraphId: tail, offset: 4 },
            });
            surface.insertPlainText(' retained');
            surface.flushPendingInput();
            await pair.settle();
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            seedReviewFieldCapacity(pair.documents[0]!, kind, 1);
            await pair.settle();
            const id = groups(alice)[0]!.id,
              before = pair.editors.map((editor) => ({
                bytes: writeOoxmlPackage(editor.surface!.session.currentPackage()),
                revision: editor.surface!.session.packageRevision(),
              })),
              vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
              counts = pair.documents.map((doc) => [
                doc.getMap(REVIEW_FIELD_TARGETS).size,
                doc.getMap(REVIEW_FIELD_INTENTS).size,
              ]),
              history = pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()]);
            expect(history[0]![1]).toBe(true);
            pair.pause();
            for (const editor of [alice, bob])
              expect(
                acceptReviewGroup(
                  editor,
                  id,
                  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
                  CELL_CONTEXT
                )
              ).toMatchObject({
                success: false,
                error: expect.stringContaining('record limit'),
              });
            pair.reconnect(receiver);
            await pair.settle();
            pair.editors.forEach((editor, index) => {
              expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(
                before[index]!.bytes
              );
              expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
              expect(groups(editor).map((group) => group.id)).toEqual([id]);
              expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
              expect([
                pair.documents[index]!.getMap(REVIEW_FIELD_TARGETS).size,
                pair.documents[index]!.getMap(REVIEW_FIELD_INTENTS).size,
              ]).toEqual(counts[index]);
            });
            expect(
              pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()])
            ).toEqual(history);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail retained');
          } finally {
            pair.destroy();
          }
        },
        30000
      );
