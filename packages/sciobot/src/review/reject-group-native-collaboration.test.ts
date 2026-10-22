import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  cellReviewDocument,
  cellScopeFor,
  childNamed,
  CELL_CONTEXT,
} from '../../test/support/table-review';
import { groups } from '../../test/support/review-groups';
import { planTableFormat } from '../diff/table-format';
import { planTableMerge } from '../diff/table-merges';
import { trackTableMergePlan } from './track-table-merges';
import { readEditorRevision } from '../projection/revision-cache';
import { trackTableFormatPlan } from './track-table-format';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { rejectReviewGroup } from './reject-group';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const revision = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
const reject = (editor: DocxEditorInstance, id: string) =>
  rejectReviewGroup(editor, id, revision(editor), CELL_CONTEXT);
// R36 retains empty structural identities in actor-bound history. Compare exact Word output,
// and verify that native export leaves the live package unchanged.
function wordFingerprint(pkg: OoxmlPackage) {
  const before = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const saved = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!saved.ok) throw new Error(saved.reason);
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toBe(before);
  return canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!);
}
function widths(part: OoxmlPart) {
  const body = childNamed(part.root, 'body')!;
  const table = body.kind !== 'textValue' && body.children.find((node) => node.kind === 'table');
  if (!table) throw new Error('Missing table');
  const grid = childNamed(table, 'tblGrid')!;
  return grid.kind === 'textValue'
    ? []
    : grid.children.map(
        (node) =>
          node.kind !== 'textValue' &&
          node.attributes.find((attribute) => attribute.localName === 'w')?.value
      );
}
for (const operation of ['widths', 'merge'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`concurrent rejection of complete native ${operation} retains original Word state until both actors undo; ${clients}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: cellReviewDocument(),
          clientIds: clients,
          ownedAttributes: true,
          editor: editorOptions,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const scope = cellScopeFor(alice),
            original = widths(alice.surface!.session.part());
          expect(original).toEqual(['3000', '3000']);
          const before = alice.surface!.session.part();
          const originalParagraph = alice
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(before, id) === 'Two')!;
          const direct = applyTreeOp(before, {
            op: 'insertText',
            paragraphId: originalParagraph,
            offset: 3,
            text: ' peer',
          });
          if (!direct.ok) throw new Error(direct.reason);
          const originalWithPeer = canonicalOoxmlFingerprint(direct.part);
          const originalWord = wordFingerprint(
            withPart(alice.surface!.session.currentPackage(), direct.part)
          );
          const input = {
            revision: revision(alice),
            part: alice.surface!.session.part(),
            tableId: scope.tableId,
          };
          if (operation === 'widths') {
            const plan = planTableFormat(input, {
              op: 'setTableProperties',
              columnWidthsTwips: [3500, 2500],
            });
            if (!plan.success) throw new Error(plan.error);
            expect(trackTableFormatPlan(alice, plan.plan, CELL_CONTEXT).success).toBe(true);
          } else {
            const plan = planTableMerge(input, {
              kind: 'merge',
              cellIds: [cellScopeFor(alice, 'A').cellId, cellScopeFor(alice, 'B').cellId],
            });
            if (!plan.success) throw new Error(plan.error);
            expect(trackTableMergePlan(alice, plan.plan, CELL_CONTEXT).success).toBe(true);
          }
          await pair.settle();
          const paragraph = bob
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Two')!;
          bob.surface!.setSelection({
            anchor: { paragraphId: paragraph, offset: 3 },
            head: { paragraphId: paragraph, offset: 3 },
          });
          bob.surface!.insertPlainText(' peer');
          bob.surface!.flushPendingInput();
          await pair.settle();
          const pending = groups(alice);
          expect(pending).toHaveLength(1);
          if (operation === 'widths') expect(pending[0]!.native).toHaveLength(6);
          else expect(pending[0]!.native.length).toBeGreaterThan(1);
          const nativeCount = collectRevisionSites(alice.surface!.session.part()).length;
          const id = pending[0]!.id;
          const pendingWord = canonicalOoxmlFingerprint(alice.surface!.session.part());
          const pendingExport = wordFingerprint(alice.surface!.session.currentPackage());
          pair.pause();
          expect(reject(alice, id).success).toBe(true);
          expect(reject(bob, id).success).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          const assertState = (pending: boolean) => {
            for (const editor of pair.editors) {
              expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(
                pending ? nativeCount : 0
              );
              expect(groups(editor).map((group) => group.id)).toEqual(pending ? [id] : []);
              expect(widths(editor.surface!.session.part())).toEqual(
                pending && operation === 'widths' ? ['3500', '2500'] : original
              );
              if (operation === 'widths')
                expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                  pending ? pendingWord : originalWithPeer
                );
              expect(wordFingerprint(editor.surface!.session.currentPackage())).toBe(
                pending ? pendingExport : originalWord
              );
              expect(paragraphTextOf(editor.surface!.session.part(), paragraph)).toBe('Two peer');
            }
          };
          assertState(false);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          assertState(false);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          assertState(true);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          assertState(false);
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (row) => row.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(paragraph);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: paragraph, offset: 8 })
          );
          expect(carets[0]).not.toBeNull();
          expect(carets[1]).not.toBeNull();
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          const bytes = new Uint8Array(await alice.save()),
            saved = readOoxmlPackage(bytes);
          if (!saved.ok) throw new Error(saved.reason);
          expect(widths(saved.package.parts.get(saved.package.mainDocumentPart)!)).toEqual(
            original
          );
          expect(
            collectRevisionSites(saved.package.parts.get(saved.package.mainDocumentPart)!)
          ).toEqual([]);
          for (const mode of [
            { document: bytes },
            {
              document: new Uint8Array(),
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[receiver]!),
            },
          ]) {
            const reopened = await createTwoEditorFixture({
              ...mode,
              ownedAttributes: true,
              editor: editorOptions,
            });
            try {
              await reopened.settle();
              for (const editor of reopened.editors) {
                expect(groups(editor)).toEqual([]);
                expect(widths(editor.surface!.session.part())).toEqual(original);
                expect(editor.surface!.session.bodyText()).toContain('Two peer');
              }
            } finally {
              reopened.destroy();
            }
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60000);
