import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { propertyRecordFixture, propertyActivity } from '../../test/support/property-record';
import {
  executeProperties,
  propertyContext,
  propertyLabel,
  propertyToken,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { GENERATED_LABEL_TAG } from '../projection/block-tags';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { readPropertyReviewStore } from './property-record-store';
import { decidePropertyProposal } from './decide-properties';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

function documentBytes() {
  const zip = unzipSync(writeOoxmlPackage(propertyRecordFixture())),
    other = `<w:sdt><w:sdtPr><w:tag w:val="${blockTag('other')}"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:tag w:val="${GENERATED_LABEL_TAG}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Other — 20 minutes</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Other body</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
  zip['word/document.xml'] = strToU8(
    strFromU8(zip['word/document.xml']!).replace('<w:sectPr>', other + '<w:sectPr>')
  );
  const loaded = readOoxmlPackage(zipSync(zip));
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      id: 'other',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Other', duration: 20 },
    },
    'Other label'
  );
  if (!stored.ok) throw new Error(stored.reason);
  return writeOoxmlPackage(stored.pkg);
}
function records(editor: DocxEditorInstance) {
  const read = readPropertyReviewStore(editor.surface!.session.currentPackage());
  if (read.status !== 'ready') throw new Error('Missing property records');
  return [...read.records.keys()].sort();
}
async function views(editor: DocxEditorInstance) {
  const token = propertyToken(editor),
    session = editor.surface!.session,
    [committed, working] = await Promise.all([
      readCommittedView(session, token),
      readWorkingView(session, token),
    ]);
  if (!committed.success) throw new Error(committed.error);
  if (!working.success) throw new Error(working.error);
  return [committed.view.projection.text, working.view.projection.text];
}
function insert(editor: DocxEditorInstance, id: string, offset: number, text: string) {
  const surface = editor.surface!;
  surface.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`property proposals and ${decision} preserve independent peer proposals, text, history, geometry, and save; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: documentBytes(),
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            bodyId = alice
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Body')!,
            phaseId = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get(
              'phase'
            )!.controlId;
          bob.surface!.setSelection({
            anchor: { paragraphId: bodyId, offset: 0 },
            head: { paragraphId: bodyId, offset: 0 },
          });
          await pair.settle();
          pair.pause();
          const first = executeProperties(
              alice,
              { set: { name: 'New activity', duration: 15 } },
              'New activity — 15 minutes',
              'alice'
            ),
            second = executeProperties(
              bob,
              { set: { name: 'Other new', duration: 25 } },
              'Other new — 25 minutes',
              'bob',
              'other'
            );
          if (!first.success) throw new Error(first.error);
          if (!second.success) throw new Error(second.error);
          insert(bob, bodyId, 0, 'Bob: ');
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['alice', 'bob']);
            const [committed, working] = await views(editor);
            expect(committed).toContain('name="Old activity" duration="10"');
            expect(committed).toContain('name="Other" duration="20"');
            expect(working).toContain('name="New activity" duration="15"');
            expect(working).toContain('name="Other new" duration="25"');
            expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe('Bob: Body');
            expect(
              propertyActivity(editor.surface!.session.currentPackage()).payload.attributes.duration
            ).toBe(10);
          }
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote).toBeDefined();
          expect(remote!.head.offset).toBe(bob.surface!.state().selection!.head.offset);
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
            expect(paragraphTextOf(editor.surface!.session.part(), propertyLabel(editor).id)).toBe(
              'Old activity — 10 minutes'
            );
            expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe('Bob: Body');
          }
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(records(editor)).toEqual(['alice', 'bob']);
          const decided = decidePropertyProposal(
            bob,
            'alice',
            decision,
            propertyToken(bob),
            propertyContext
          );
          if (!decided.success) throw new Error(decided.error);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['bob']);
            expect(
              propertyActivity(editor.surface!.session.currentPackage()).payload.attributes.duration
            ).toBe(decision === 'accept' ? 15 : 10);
            expect(paragraphTextOf(editor.surface!.session.part(), propertyLabel(editor).id)).toBe(
              decision === 'accept' ? 'New activity — 15 minutes' : 'Old activity — 10 minutes'
            );
            expect(
              buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('phase')!.controlId
            ).toBe(phaseId);
          }
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual(['alice', 'bob']);
            expect(
              propertyActivity(editor.surface!.session.currentPackage()).payload.attributes.duration
            ).toBe(10);
            expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe('Bob: Body');
          }
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(
            decidePropertyProposal(alice, 'bob', 'accept', propertyToken(alice), propertyContext)
              .success
          ).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(records(editor)).toEqual([]);
            expect(
              buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('other')!.payload
                .attributes.duration
            ).toBe(25);
            expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe('Bob: Body');
          }
          const final = await pair.savedDigests();
          expect(final[0]).toEqual(final[1]);
        } finally {
          pair.destroy();
        }
      }, 30_000);

for (const receiver of [0, 1] as const)
  test(`owner deletion during a property proposal preserves record data and cannot recreate the deleted block; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: documentBytes(),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        owner = propertyActivity(alice.surface!.session.currentPackage());
      pair.pause();
      expect(
        alice.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).rejected
      ).toBe(false);
      expect(executeProperties(bob).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(editor.surface!.session.currentPackage())).toBeUndefined();
        expect(records(editor)).toEqual([propertyContext.proposalId]);
        expect(
          decidePropertyProposal(
            editor,
            propertyContext.proposalId,
            'accept',
            propertyToken(editor),
            propertyContext
          ).success
        ).toBe(false);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(propertyActivity(editor.surface!.session.currentPackage())).toBeDefined();
        expect((await views(editor))[1]).toContain('duration="15"');
      }
      expect(
        decidePropertyProposal(
          bob,
          propertyContext.proposalId,
          'reject',
          propertyToken(bob),
          propertyContext
        ).success
      ).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);

for (const decision of ['accept', 'reject'] as const)
  test(`property ${decision} retains a peer's later text inside the generated fallback`, async () => {
    const pair = await createTwoEditorFixture({
      document: documentBytes(),
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      expect(executeProperties(alice).success).toBe(true);
      await pair.settle();
      const paragraph = propertyLabel(bob),
        text = paragraphTextOf(bob.surface!.session.part(), paragraph.id)!;
      insert(bob, paragraph.id, text.length, ' Bob');
      await pair.settle();
      expect(
        decidePropertyProposal(
          alice,
          propertyContext.proposalId,
          decision,
          propertyToken(alice),
          propertyContext
        ).success
      ).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(paragraphTextOf(editor.surface!.session.part(), propertyLabel(editor).id)).toBe(
          (decision === 'accept' ? 'New activity — 15 minutes' : 'Old activity — 10 minutes') +
            ' Bob'
        );
        expect(
          propertyActivity(editor.surface!.session.currentPackage()).payload.attributes.duration
        ).toBe(decision === 'accept' ? 15 : 10);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(records(editor)).toEqual([propertyContext.proposalId]);
        expect(paragraphTextOf(editor.surface!.session.part(), propertyLabel(editor).id)).toContain(
          ' Bob'
        );
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);
