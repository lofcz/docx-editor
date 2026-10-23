// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { bornWordBody, paraIds, wordPart } from '../../test/support/born-word-body';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupToken } from '../../test/support/review-groups';
import { planTargetDocument } from '../diff/plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';

const context = { ...propertyContext, date: '2026-10-23T00:00:00Z' };
const documentBytes = () =>
  docxFixture('<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>');
async function views(editor: DocxEditorInstance) {
  const original = await readCommittedView(editor.surface!.session, groupToken(editor));
  const working = await readWorkingView(editor.surface!.session, groupToken(editor));
  if (!original.success || !working.success) throw Error('Missing review views');
  return [original.view.projection.text, working.view.projection.text];
}
async function directTarget(target: string) {
  const loaded = readOoxmlPackage(documentBytes());
  if (!loaded.ok) throw Error(loaded.reason);
  const editor = propertyEditor(loaded.package);
  const projection = await readWorkingView(editor.surface!.session, groupToken(editor));
  if (!projection.success) throw Error(projection.error);
  const plan = planTargetDocument(
    {
      pkg: editor.surface!.session.currentPackage(),
      revision: groupToken(editor),
      projection: projection.view.projection,
    },
    target
  );
  if (!plan.success) throw Error(plan.error);
  const applied = applyDirectPlan(editor, plan.plan, context);
  if (!applied.success) throw Error(applied.error);
  return editor;
}
for (const clientIds of [
  [101, 202],
  [202, 101],
] as const)
  for (const receiver of [0, 1] as const)
    test(`born table later waves retain peer text, native decisions, history, carets, and reopen; ids=${clientIds}, receiver=${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: documentBytes(),
        documentId: context.documentId,
        ownedAttributes: true,
        clientIds,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      const table = '| A | **B** |\n| :--- | ---: |\n| [C](https://example.com/) | `D` |';
      const firstTarget = `Before\n\n- One\n\n${table}\n\n- Two\n\nAfter`;
      const nextTarget = firstTarget.replace('| A |', '| Revised |');
      const direct = await directTarget(nextTarget);
      const baseline = await directTarget('Before\n\nAfter');
      for (const editor of [direct, baseline]) {
        const paragraphId = editor
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === 'After')!;
        expect(
          editor.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 5, text: ' Peer' },
          ]).rejected
        ).toBe(false);
      }
      try {
        const [alice, bob] = pair.editors;
        const imported = new Set(paraIds(wordPart(alice!.surface!.session.currentPackage()).root));
        const first = await recomposeOwnedProposals(
          alice!,
          {
            revision: groupToken(alice!),
            targetMarkdown: firstTarget,
            ownership: [],
            contentProofs: [],
          },
          context
        );
        if (!first.success) throw Error(first.error);
        await pair.settle();
        const firstViews = await views(alice!);
        pair.pause();
        const after = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'After')!;
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: after, offset: 5, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        const second = await recomposeOwnedProposals(
          alice!,
          {
            revision: groupToken(alice!),
            targetMarkdown: nextTarget,
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          context
        );
        if (!second.success) throw Error(second.error);
        pair.reconnect(receiver);
        await pair.settle();
        const expected = [(await views(baseline))[0]!, (await views(direct))[1]!];
        for (const editor of pair.editors) expect(await views(editor)).toEqual(expected);
        const cell = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Revised')!;
        expect(cell).toBeDefined();
        bob!.surface!.setSelection({
          anchor: { paragraphId: cell, offset: 3 },
          head: { paragraphId: cell, offset: 3 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (value) => value.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(cell);
        expect(remote?.head.offset).toBe(3);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId: cell, offset: 3 })
        );
        expect(carets[0]).not.toBeNull();
        expect(carets[1]).not.toBeNull();
        expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({ x: carets[1]!.x, y: carets[1]!.y });
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual(
            firstViews.map((text) => text.replace('After', 'After Peer'))
          );
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const pending = canonicalOoxmlFingerprint(
          wordPart(alice!.surface!.session.currentPackage())
        );
        for (const action of ['accept', 'reject'] as const) {
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage();
            if (action === 'accept')
              expect(bornWordBody(pkg, imported)).toBe(
                bornWordBody(direct.surface!.session.currentPackage(), imported)
              );
            else
              expect(canonicalOoxmlFingerprint(wordPart(pkg))).toBe(
                canonicalOoxmlFingerprint(wordPart(baseline.surface!.session.currentPackage()))
              );
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(
              canonicalOoxmlFingerprint(wordPart(editor.surface!.session.currentPackage()))
            ).toBe(pending);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
        }
        const saved = new Uint8Array(await alice!.save());
        const loaded = readOoxmlPackage(saved);
        if (!loaded.ok) throw Error(loaded.reason);
        const reopened = propertyEditor(loaded.package);
        try {
          expect(await views(reopened)).toEqual(expected);
        } finally {
          reopened.destroy();
        }
        const room = await createTwoEditorFixture({
          document: saved,
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          await room.settle();
          for (const editor of room.editors) expect(await views(editor)).toEqual(expected);
          room.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          room.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        baseline.destroy();
        direct.destroy();
        pair.destroy();
      }
    }, 30000);
