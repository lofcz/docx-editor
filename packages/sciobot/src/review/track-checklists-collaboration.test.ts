// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { listDocument, listParagraph, checklistParagraph } from '../../test/support/list-document';
import { propertyEditor } from '../../test/support/property-execution';
import { readEditorRevision } from '../projection/revision-cache';
import { planListMembership } from '../diff/lists';
import type { ListTarget } from '../diff/list-types';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { trackChecklistPlan } from './track-checklists';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const context = {
  documentId: 'checklist-review',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-22T09:00:00Z',
};
const cases: [string, string, ListTarget, string, string][] = [
  [
    'create',
    listParagraph('Task'),
    { kind: 'checklist', level: 0, checked: false },
    'Task Peer',
    '- [ ] Task Peer',
  ],
  [
    'check',
    checklistParagraph(false),
    { kind: 'checklist', level: 0, checked: true },
    '- [ ] Task Peer',
    '- [x] Task Peer',
  ],
  [
    'nest',
    checklistParagraph(false),
    { kind: 'checklist', level: 2, checked: false },
    '- [ ] Task Peer',
    '    - [ ] Task Peer',
  ],
  ['remove', checklistParagraph(false), { kind: 'none' }, '- [ ] Task Peer', 'Task Peer'],
];
async function views(editor: DocxEditorInstance) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const results = await Promise.all([
    readCommittedView(editor.surface!.session, revision),
    readWorkingView(editor.surface!.session, revision),
  ]);
  return results.map((result) => {
    if (!result.success) throw Error(result.error);
    return result.view.projection.text.trimEnd();
  });
}
function propose(editor: DocxEditorInstance, target: ListTarget) {
  const session = editor.surface!.session;
  const plan = planListMembership(
    {
      pkg: session.currentPackage(),
      partName: session.part().name,
      revision: readEditorRevision(editor, context.documentId, context.epoch)!,
    },
    [{ paragraphId: session.paragraphIds()[0]!, target }]
  );
  if (!plan.success) throw Error(plan.error);
  const result = trackChecklistPlan(editor, plan.plan, context);
  if (!result.success) throw Error(result.error);
}
for (const [name, body, target, original, proposed] of cases)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`tracked checklist ${name}, peer suffix, clients ${clientIds}, receiver ${receiver}`, async () => {
        const options = {
          document: listDocument(body),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          clientIds,
        };
        const pair = await createTwoEditorFixture(options);
        const [alice, bob] = pair.editors;
        const same = async (expected: readonly string[]) => {
          for (const editor of pair.editors) expect(await views(editor)).toEqual([...expected]);
        };
        try {
          const id = bob!.surface!.session.paragraphIds()[0]!;
          pair.pause();
          propose(alice!, target);
          const offset = paragraphTextOf(bob!.surface!.session.part(), id)!.length;
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: id, offset, text: ' Peer' },
            ]).rejected
          ).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          await same([original, proposed]);
          const end = paragraphTextOf(bob!.surface!.session.part(), id)!.length;
          bob!.surface!.setSelection({
            anchor: { paragraphId: id, offset: end },
            head: { paragraphId: id, offset: end },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(id);
          expect(remote.head.offset).toBe(end);
          const caret = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: id, offset: end })
          );
          expect(caret[0]).not.toBeNull();
          expect({ x: caret[0]!.x, y: caret[0]!.y }).toEqual({ x: caret[1]!.x, y: caret[1]!.y });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await same([original, original]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await same([original, proposed]);
          for (const action of ['accept', 'reject'] as const) {
            expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            await same([
              action === 'accept' ? proposed : original,
              action === 'accept' ? proposed : original,
            ]);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await same([original, proposed]);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
          }
          expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
          const read = readOoxmlPackage(new Uint8Array(await alice!.save()));
          if (!read.ok) throw Error(read.reason);
          const saved = propertyEditor(read.package);
          try {
            expect(await views(saved)).toEqual([original, proposed]);
          } finally {
            saved.destroy();
          }
          const resumed = await createTwoEditorFixture({
            ...options,
            clientIds: [181, 211],
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          });
          try {
            for (const editor of resumed.editors)
              expect(await views(editor)).toEqual([original, proposed]);
          } finally {
            resumed.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const [name, body, target] of cases)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`tracked checklist ${name} preserves peer deletion; clients ${clientIds}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: listDocument(body),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          clientIds,
        });
        const [alice, bob] = pair.editors;
        const original = name === 'create' ? '' : '- [ ]';
        const proposed =
          name === 'remove'
            ? ''
            : name === 'check'
              ? '- [x]'
              : name === 'nest'
                ? '    - [ ]'
                : '- [ ]';
        const same = async (expected: string[]) => {
          for (const editor of pair.editors) expect(await views(editor)).toEqual(expected);
        };
        try {
          const id = bob!.surface!.session.paragraphIds()[0]!;
          const end = paragraphTextOf(bob!.surface!.session.part(), id)!.length;
          pair.pause();
          propose(alice!, target);
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'deleteText', paragraphId: id, start: end - 4, end },
            ]).rejected
          ).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          await same([original, proposed]);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await same([original, original]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await same([original, proposed]);
          for (const action of ['accept', 'reject'] as const) {
            expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            await same([
              action === 'accept' ? proposed : original,
              action === 'accept' ? proposed : original,
            ]);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await same([original, proposed]);
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);
