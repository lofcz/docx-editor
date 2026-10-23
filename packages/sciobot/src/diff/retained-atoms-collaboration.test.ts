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
    test(`retained paragraph atom births preserve same-paragraph peer edits, both histories, geometry and DOCX/room reopen ${clients}/${receiver}`, async () => {
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
        const target = original.replace(
          'Neighbor',
          'Neighbor A $x^2$ [Ada](mention:user-7) <source url="https://example.test" title="Paper" /> Z'
        );
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
          { allocateBlockId: () => `retained-${++next}` }
        );
        if (!planned.success) throw Error(planned.error);
        expect(decodeEditPlan(serializeEditPlan(planned.plan))).toEqual({
          success: true,
          plan: planned.plan,
        });
        let peerLength = 0;
        const type = (text: string) => {
          const result = bob!.surface!.applyAutomationOps(
            () => [
              {
                op: 'insertText',
                paragraphId: neighbor,
                offset: 5 + peerLength,
                text,
              },
            ],
            { kind: 'body' },
            [],
            false,
            true
          );
          expect(result.rejected).toBe(false);
          peerLength += text.length;
        };
        const check = async (present = true, suffix = ' Peer', tail = '') => {
          const bytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage();
            expect(importedMath(pkg)).toEqual(omml);
            const index = buildBlockIndex(pkg);
            expect(index.diagnostics).toEqual([]);
            const block = index.byId.get('retained-1');
            expect(block?.payload.attributes.texExpression).toBe(present ? expression : undefined);
            expect(index.byId.get('retained-2')?.payload.type).toBe(
              present ? 'mention' : undefined
            );
            expect(index.byId.get('retained-3')?.payload.type).toBe(present ? 'source' : undefined);
            expect(paragraphTextOf(editor.surface!.session.part(), neighbor)).toBe(
              'Neigh' + suffix + 'bor' + (present ? ' A x^2 Ada Paper Z' : '') + tail
            );
            const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
            const projected = await readWorkingView(editor.surface!.session, revision);
            if (!projected.success) throw Error(projected.error);
            expect(editor.surface!.session.currentPackage()).toBe(pkg);
            expect(key(projected.view.projection.text)).toBe(
              key(
                (present ? target : original).replace('Neighbor', 'Neigh' + suffix + 'bor') + tail
              )
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
        const peerTail = bob!.surface!.applyAutomationOps(
          () => [
            {
              op: 'insertText',
              paragraphId: neighbor,
              offset: paragraphTextOf(bob!.surface!.session.part(), neighbor)!.length,
              text: ' More',
            },
          ],
          { kind: 'body' },
          [],
          false,
          true
        );
        expect(peerTail.rejected).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        await check(false, ' Peer', ' More');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, ' Peer', ' More');
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(true, '');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, ' Peer', ' More');
        const offset = paragraphTextOf(bob!.surface!.session.part(), neighbor)!.length;
        bob!.surface!.setSelection({
          anchor: { paragraphId: neighbor, offset },
          head: { paragraphId: neighbor, offset },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob');
        expect(remote?.head.nodeId).toBe(neighbor);
        expect(remote?.head.offset).toBe(offset);
        const a = caretAt(alice!.surface!.layout(), { paragraphId: neighbor, offset }),
          b = caretAt(bob!.surface!.layout(), { paragraphId: neighbor, offset });
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
                buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('retained-1')
                  ?.payload.attributes.texExpression
              ).toBe(expression);
              const projected = await readWorkingView(
                editor.surface!.session,
                readEditorRevision(editor, context.documentId, context.epoch)!
              );
              if (!projected.success) throw Error(projected.error);
              expect(key(projected.view.projection.text)).toBe(
                key(target.replace('Neighbor', 'Neigh Peerbor') + ' More')
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
