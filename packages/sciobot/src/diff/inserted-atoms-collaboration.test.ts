// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { mathDocument, importedMath } from '../../test/support/math-document';
import { prepareNativeDocument } from '../native/prepare-document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { planTargetDocument } from '../diff/plan-target';
import { decodeEditPlan, serializeEditPlan } from '../diff/plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
const context = { documentId: 'math', epoch: 'shared' };
const key = (markdown: string) => {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
};
const expression = 'x^2';
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`AI inline births preserve peer edits, both histories, geometry and DOCX/room reopen ${clients}/${receiver}`, async () => {
      const prepared = prepareNativeDocument(mathDocument());
      if (!prepared.success) throw Error(prepared.error);
      const fixture = { document: prepared.document, clientIds: clients };
      const pair = await createTwoEditorFixture(fixture);
      try {
        const [alice, bob] = pair.editors;
        const token = readEditorRevision(alice!, context.documentId, context.epoch)!;
        const view = await readWorkingView(alice!.surface!.session, token);
        if (!view.success) throw Error(view.error);
        const original = view.view.projection.text;
        const target =
          original +
          '\n\nA $x^2$ [Ada](mention:user-7) <source url="https://example.test" title="Paper" /> Z';
        const omml = importedMath(alice!.surface!.session.currentPackage());
        const neighbor = alice!.surface!.session.paragraphIds().at(-1)!;
        let next = 0;
        const planned = planTargetDocument(
          {
            revision: token,
            pkg: alice!.surface!.session.currentPackage(),
            projection: view.view.projection,
          },
          target,
          { allocateBlockId: () => `shared-inline-${++next}` }
        );
        if (!planned.success) throw Error(planned.error);
        expect(decodeEditPlan(serializeEditPlan(planned.plan))).toEqual({
          success: true,
          plan: planned.plan,
        });
        const type = (text: string) => {
          const result = bob!.surface!.applyAutomationOps(
            () => [
              {
                op: 'insertText',
                paragraphId: neighbor,
                offset: paragraphTextOf(bob!.surface!.session.part(), neighbor)!.length,
                text,
              },
            ],
            { kind: 'body' },
            [],
            false,
            true
          );
          expect(result.rejected).toBe(false);
        };
        const check = async (present = true, suffix = ' Peer') => {
          const bytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage();
            expect(importedMath(pkg)).toEqual(omml);
            const index = buildBlockIndex(pkg);
            expect(index.diagnostics).toEqual([]);
            const block = index.byId.get('shared-inline-1');
            expect(block?.payload.attributes.texExpression).toBe(present ? expression : undefined);
            expect(index.byId.get('shared-inline-2')?.payload.type).toBe(
              present ? 'mention' : undefined
            );
            expect(index.byId.get('shared-inline-3')?.payload.type).toBe(
              present ? 'source' : undefined
            );
            expect(paragraphTextOf(editor.surface!.session.part(), neighbor)).toBe(
              'Neighbor' + suffix
            );
            const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
            const projected = await readWorkingView(editor.surface!.session, revision);
            if (!projected.success) throw Error(projected.error);
            expect(editor.surface!.session.currentPackage()).toBe(pkg);
            expect(key(projected.view.projection.text)).toBe(
              key((present ? target : original).replace('Neighbor', 'Neighbor' + suffix))
            );
          }
          pair.documents.forEach((doc, index) =>
            expect(Y.encodeStateAsUpdate(doc)).toEqual(bytes[index]!)
          );
          expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
            canonicalOoxmlFingerprint(bob!.surface!.session.part())
          );
          const [a, b] = await pair.savedDigests();
          expect(a).toEqual(b);
        };
        pair.pause();
        const applied = applyDirectPlan(alice!, planned.plan, context);
        if (!applied.success) throw Error(applied.error);
        type(' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        await check();
        pair.pause();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        type(' More');
        pair.reconnect(receiver);
        await pair.settle();
        await check(false, ' Peer More');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, ' Peer More');
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(true, '');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, ' Peer More');
        const atomParagraph = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'A x^2 Ada Paper Z')!;
        expect(atomParagraph).toBeDefined();
        const offset = paragraphTextOf(bob!.surface!.session.part(), atomParagraph)!.length;
        bob!.surface!.setSelection({
          anchor: { paragraphId: atomParagraph, offset },
          head: { paragraphId: atomParagraph, offset },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob');
        expect(remote?.head.nodeId).toBe(atomParagraph);
        expect(remote?.head.offset).toBe(offset);
        const a = caretAt(alice!.surface!.layout(), { paragraphId: atomParagraph, offset }),
          b = caretAt(bob!.surface!.layout(), { paragraphId: atomParagraph, offset });
        expect(a).not.toBeNull();
        expect(b).not.toBeNull();
        expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
        const expected = canonicalOoxmlFingerprint(alice!.surface!.session.part());
        for (const shared of [false, true]) {
          const reopened = await createTwoEditorFixture({
            ...fixture,
            clientIds: undefined,
            document: new Uint8Array(await alice!.save()),
            ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) {
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(expected);
              expect(importedMath(editor.surface!.session.currentPackage())).toEqual(omml);
              expect(
                buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(
                  'shared-inline-1'
                )?.payload.attributes.texExpression
              ).toBe(expression);
              const projected = await readWorkingView(
                editor.surface!.session,
                readEditorRevision(editor, context.documentId, context.epoch)!
              );
              if (!projected.success) throw Error(projected.error);
              expect(key(projected.view.projection.text)).toBe(
                key(target.replace('Neighbor', 'Neighbor Peer More'))
              );
            }
            const [a, b] = await reopened.savedDigests();
            expect(a).toEqual(b);
          } finally {
            reopened.destroy();
          }
        }
      } finally {
        pair.destroy();
      }
    });

test('unprepared shared atom births refuse without shared writes or history', async () => {
  const pair = await createTwoEditorFixture({ document: mathDocument() });
  try {
    const editor = pair.editors[0]!;
    const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
    const projected = await readWorkingView(editor.surface!.session, revision);
    if (!projected.success) throw Error(projected.error);
    const planned = planTargetDocument(
      {
        revision,
        pkg: editor.surface!.session.currentPackage(),
        projection: projected.view.projection,
      },
      projected.view.projection.text + '\n\n$x^2$',
      { allocateBlockId: () => 'unprepared' }
    );
    if (!planned.success) throw Error(planned.error);
    const before = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    const pkg = editor.surface!.session.currentPackage();
    expect(applyDirectPlan(editor, planned.plan, context)).toEqual({
      success: false,
      code: 'unsupported-operation',
      error: 'Prepare native payload storage before sharing inline atom births.',
    });
    expect(editor.surface!.session.currentPackage()).toBe(pkg);
    expect(editor.surface!.session.canUndo()).toBe(false);
    pair.documents.forEach((doc, index) =>
      expect(Y.encodeStateAsUpdate(doc)).toEqual(before[index]!)
    );
  } finally {
    pair.destroy();
  }
});
