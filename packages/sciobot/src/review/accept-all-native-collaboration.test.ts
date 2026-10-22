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
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { cellReviewDocument, cellScopeFor, CELL_CONTEXT } from '../../test/support/table-review';
import { groups, proposeGroupText } from '../../test/support/review-groups';
import { planTableFormat } from '../diff/table-format';
import { planTableMerge } from '../diff/table-merges';
import { readEditorRevision } from '../projection/revision-cache';
import { trackTableFormatPlan } from './track-table-format';
import { trackTableMergePlan } from './track-table-merges';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptAllReviewGroups } from './accept-all';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const revision = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
function word(pkg: OoxmlPackage) {
  const before = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toBe(before);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
for (const kind of ['widths', 'merge', 'split'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(
        'bulk native ' +
          kind +
          ' and neighboring text retain exact Word through concurrent decisions and both actors history; ' +
          clients +
          '/' +
          receiver,
        async () => {
          const pair = await createTwoEditorFixture({
            document: cellReviewDocument(),
            ownedAttributes: true,
            clientIds: clients,
            editor: options,
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
              scope = cellScopeFor(alice),
              part = alice.surface!.session.part(),
              input = { part, revision: revision(alice), tableId: scope.tableId };
            if (kind === 'widths') {
              const plan = planTableFormat(input, {
                op: 'setTableProperties',
                columnWidthsTwips: [3500, 2500],
              });
              if (!plan.success) throw new Error(plan.error);
              expect(trackTableFormatPlan(alice, plan.plan, CELL_CONTEXT).success).toBe(true);
            } else {
              const plan = planTableMerge(
                input,
                kind === 'merge'
                  ? {
                      kind: 'merge',
                      cellIds: [cellScopeFor(alice, 'A').cellId, cellScopeFor(alice, 'B').cellId],
                    }
                  : { kind: 'split', cellId: scope.cellId, rows: 2, columns: 2 }
              );
              if (!plan.success) throw new Error(plan.error);
              expect(trackTableMergePlan(alice, plan.plan, CELL_CONTEXT).success).toBe(true);
            }
            proposeGroupText(alice, 'Tail', 'Tail changed');
            await pair.settle();
            const peer = bob
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Two')!;
            bob.surface!.setSelection({
              anchor: { paragraphId: peer, offset: 3 },
              head: { paragraphId: peer, offset: 3 },
            });
            bob.surface!.insertPlainText(' peer');
            bob.surface!.flushPendingInput();
            await pair.settle();
            const ids = groups(alice).map((group) => group.id),
              pending = alice.surface!.session.currentPackage(),
              sites = collectRevisionSites(alice.surface!.session.part()).length;
            expect(ids.length).toBeGreaterThan(1);
            const accepted = applyTreeOp(alice.surface!.session.part(), {
              op: 'acceptAllRevisions',
            });
            if (!accepted.ok) throw new Error(accepted.reason);
            const expected = word(withPart(pending, accepted.part)),
              original = word(pending);
            pair.pause();
            for (const editor of pair.editors)
              expect(acceptAllReviewGroups(editor, revision(editor), CELL_CONTEXT).success).toBe(
                true
              );
            pair.reconnect(receiver);
            await pair.settle();
            const check = (resolved: boolean) => {
              for (const editor of pair.editors) {
                expect(groups(editor).map((group) => group.id)).toEqual(resolved ? [] : ids);
                expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(
                  resolved ? 0 : sites
                );
                expect(word(editor.surface!.session.currentPackage())).toBe(
                  resolved ? expected : original
                );
                expect(
                  caretAt(editor.surface!.layout(), { paragraphId: peer, offset: 3 })
                ).not.toBeNull();
              }
            };
            check(true);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            check(true);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            check(false);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            check(true);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            check(true);
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
            const reopened = await createTwoEditorFixture({
              document: new Uint8Array(await alice.save()),
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
              ownedAttributes: true,
              editor: options,
            });
            try {
              for (const editor of reopened.editors) {
                expect(groups(editor)).toEqual([]);
                expect(word(editor.surface!.session.currentPackage())).toBe(expected);
              }
            } finally {
              reopened.destroy();
            }
          } finally {
            pair.destroy();
          }
        },
        30000
      );
