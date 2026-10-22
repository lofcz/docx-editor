import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { groups, groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { cellReviewDocument, cellScopeFor, CELL_CONTEXT } from '../../test/support/table-review';
import { planTableFormat } from '../diff/table-format';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptReviewGroup } from './accept-group';
import { trackTableFormatPlan } from './track-table-format';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
function type(editor: DocxEditorInstance, id: string, text: string) {
  const surface = editor.surface!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`complete body/header group retains peer acceptance through both undo steps and native save; ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(crossStoryGroupFixture()),
        clientIds: clients,
        ownedAttributes: true,
        editor: editorOptions,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = 'g1:stored:rg_cross_story',
          before = groups(alice).map((group) => group.id),
          bodyId = groups(alice).find((group) => group.id === id)!.anchor!.paragraphId!;
        expect(groups(alice).find((group) => group.id === id)!.native).toHaveLength(2);
        pair.pause();
        expect(acceptReviewGroup(alice, id, groupToken(alice), propertyContext).success).toBe(true);
        type(bob, bodyId, 'Peer: ');
        expect(acceptReviewGroup(bob, id, groupToken(bob), propertyContext).success).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        const accepted = () => {
          for (const editor of pair.editors) {
            expect(groups(editor).map((group) => group.id)).toEqual(
              before.filter((value) => value !== id)
            );
            expect(collectRevisionSites(source(editor).parts.get('/word/header1.xml')!)).toEqual(
              []
            );
            expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(1);
            expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe(
              'Peer: Body Pending'
            );
          }
        };
        accepted();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        accepted();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(groups(editor).map((group) => group.id)).toEqual(before);
          expect(groups(editor).find((group) => group.id === id)!.native).toHaveLength(2);
          expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe(
            'Peer: Body Pending'
          );
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        accepted();
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
        for (const editor of pair.editors) {
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(collectRevisionSites(saved.package.parts.get('/word/header1.xml')!)).toEqual([]);
          expect(collectRevisionSites(saved.package.parts.get('/word/document.xml')!)).toHaveLength(
            1
          );
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
      } finally {
        pair.destroy();
      }
    }, 30000);

function widths(part: OoxmlPart) {
  const body = part.root.children.find((n) => n.kind === 'body')!,
    table = body.kind !== 'textValue' && body.children.find((n) => n.kind === 'table');
  if (!table || table.kind === 'textValue') throw new Error('Missing table');
  const grid = table.children.find((n) => n.kind !== 'textValue' && n.localName === 'tblGrid');
  if (!grid || grid.kind === 'textValue') throw new Error('Missing grid');
  return grid.children
    .filter((n) => n.kind !== 'textValue' && n.localName === 'gridCol')
    .map((n) => n.kind !== 'textValue' && n.attributes.find((a) => a.localName === 'w')?.value);
}

for (const receiver of [0, 1] as const)
  for (const concurrentAcceptance of [false, true])
    test(`one complete table format group accepts all native members and retains peer cell text through history; receiver=${receiver}/both=${concurrentAcceptance}`, async () => {
      const pair = await createTwoEditorFixture({
        document: cellReviewDocument(),
        clientIds: [41, 73],
        ownedAttributes: true,
        editor: editorOptions,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          scope = cellScopeFor(alice),
          plan = planTableFormat(
            {
              part: alice.surface!.session.part(),
              revision: readEditorRevision(alice, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
              tableId: scope.tableId,
            },
            { op: 'setTableProperties', columnWidthsTwips: [3500, 2500] }
          );
        if (!plan.success) throw new Error(plan.error);
        expect(trackTableFormatPlan(alice, plan.plan, CELL_CONTEXT).success).toBe(true);
        await pair.settle();
        expect(groups(alice)).toHaveLength(1);
        const id = groups(alice)[0]!.id;
        pair.pause();
        expect(
          acceptReviewGroup(
            alice,
            id,
            readEditorRevision(alice, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
            CELL_CONTEXT
          ).success
        ).toBe(true);
        type(bob, scope.paragraphId, 'Peer: ');
        if (concurrentAcceptance)
          expect(
            acceptReviewGroup(
              bob,
              id,
              readEditorRevision(bob, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
              CELL_CONTEXT
            ).success
          ).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(groups(editor)).toEqual([]);
          expect(collectRevisionSites(editor.surface!.session.part())).toEqual([]);
          expect(widths(editor.surface!.session.part())).toEqual(['3500', '2500']);
          expect(paragraphTextOf(editor.surface!.session.part(), scope.paragraphId)).toBe(
            'Peer: Two'
          );
        }
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(groups(editor).map((group) => group.id)).toEqual(concurrentAcceptance ? [] : [id]);
          expect(paragraphTextOf(editor.surface!.session.part(), scope.paragraphId)).toBe(
            'Peer: Two'
          );
        }
        if (concurrentAcceptance) {
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(groups(editor).map((group) => group.id)).toEqual([id]);
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(widths(saved.package.parts.get('/word/document.xml')!)).toEqual(['3500', '2500']);
          expect(collectRevisionSites(saved.package.parts.get('/word/document.xml')!)).toEqual([]);
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
      } finally {
        pair.destroy();
      }
    }, 30000);
