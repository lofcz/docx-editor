// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { importedMath, mathDocument } from '../../test/support/math-document';
import { prepareNativeDocument } from '../native/prepare-document';
import { buildBlockIndex } from './block-index';
import { createMathPayload, mathWordFallback } from './math-codec';
import { insertMath } from './math-insert';

const inline = createMathPayload('inline-math', 'inline', '  \\frac{𝛼}{β} < 1  ');
const block = createMathPayload(
  'block-math',
  'block',
  '\\begin{aligned}\r\n x &= y \\\\\n z &= 𝛼\r\\end{aligned}\n'
);
const word = (pkg: OoxmlPackage) => canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`math sources survive concurrent insertion, peer editing during deletion, both histories, geometry and reopen ${clients}/${receiver}`, async () => {
      const prepared = prepareNativeDocument(mathDocument());
      if (!prepared.success) throw Error(prepared.error);
      const fixture = { document: prepared.document, clientIds: clients };
      const pair = await createTwoEditorFixture(fixture);
      try {
        const [alice, bob] = pair.editors;
        const first = alice!.surface!.session.paragraphIds()[0]!;
        const originalMath = importedMath(alice!.surface!.session.currentPackage());
        expect(originalMath).toHaveLength(2);
        const check = async (hasInline = true, hasBlock = true, suffix = '') => {
          for (const editor of pair.editors) {
            const live = editor.surface!.session.currentPackage();
            const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!loaded.ok) throw Error(loaded.reason);
            for (const pkg of [live, loaded.package]) {
              const index = buildBlockIndex(pkg);
              expect(index.diagnostics).toEqual([]);
              expect(index.byId.get(inline.id)?.payload).toEqual(hasInline ? inline : undefined);
              expect(index.byId.get(block.id)?.payload).toEqual(hasBlock ? block : undefined);
              expect(importedMath(pkg)).toEqual(originalMath);
              const part = pkg.parts.get(pkg.mainDocumentPart)!;
              // Reopen can assign fresh native IDs; paragraph order is stable.
              const body =
                part.root.kind !== 'textValue' &&
                part.root.children.find((n) => n.kind !== 'textValue' && n.localName === 'body');
              if (!body) throw Error('Missing body');
              const paragraph = body.children.find((n) => n.kind === 'paragraph')!;
              expect(paragraphTextOf(part, paragraph.id)).toBe(
                'Before' +
                  (hasInline ? mathWordFallback(inline).inlineText : '') +
                  '\uFFFCAfter' +
                  suffix
              );
            }
          }
          expect(word(pair.editors[0]!.surface!.session.currentPackage())).toBe(
            word(pair.editors[1]!.surface!.session.currentPackage())
          );
          const [a, b] = await pair.savedDigests();
          expect(a).toEqual(b);
        };
        pair.pause();
        const a = insertMath(alice!, {
          payload: inline,
          position: { kind: 'inline', paragraphId: first, offset: 6 },
        });
        if (!a.success) throw Error(a.error);
        const b = insertMath(bob!, {
          payload: block,
          position: { kind: 'block', parentControlId: null, beforeId: null },
        });
        if (!b.success) throw Error(b.error);
        pair.reconnect(receiver);
        await pair.settle();
        await check();
        // Delete Alice's equation while Bob edits the same paragraph after imported OMML.
        pair.pause();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        const end = paragraphTextOf(bob!.surface!.session.part(), first)!.length;
        bob!.surface!.setSelection({
          anchor: { paragraphId: first, offset: end },
          head: { paragraphId: first, offset: end },
        });
        bob!.surface!.insertPlainText(' peer');
        bob!.surface!.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
        await check(false, true, ' peer');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, true, ' peer');
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(true, false);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, true, ' peer');
        const offset = paragraphTextOf(bob!.surface!.session.part(), first)!.length;
        bob!.surface!.setSelection({
          anchor: { paragraphId: first, offset },
          head: { paragraphId: first, offset },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob');
        expect(remote?.head.nodeId).toBe(first);
        expect(remote?.head.offset).toBe(offset);
        const left = caretAt(alice!.surface!.layout(), { paragraphId: first, offset });
        const right = caretAt(bob!.surface!.layout(), { paragraphId: first, offset });
        expect(left).not.toBeNull();
        expect(right).not.toBeNull();
        expect({ x: left!.x, y: left!.y }).toEqual({ x: right!.x, y: right!.y });
        const expectedWord = word(alice!.surface!.session.currentPackage());
        for (const shared of [false, true]) {
          const reopened = await createTwoEditorFixture({
            ...fixture,
            document: new Uint8Array(await alice!.save()),
            ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) {
              const pkg = editor.surface!.session.currentPackage();
              expect(word(pkg)).toBe(expectedWord);
              expect(importedMath(pkg)).toEqual(originalMath);
              expect(buildBlockIndex(pkg).byId.get(inline.id)?.payload).toEqual(inline);
              expect(buildBlockIndex(pkg).byId.get(block.id)?.payload).toEqual(block);
            }
            const [savedA, savedB] = await reopened.savedDigests();
            expect(savedA).toEqual(savedB);
          } finally {
            reopened.destroy();
          }
        }
      } finally {
        pair.destroy();
      }
    });

test('unprepared shared equation insertion refuses before publication, metadata, or history changes', async () => {
  const pair = await createTwoEditorFixture({ document: mathDocument() });
  try {
    const editor = pair.editors[0]!;
    const paragraphId = editor.surface!.session.paragraphIds()[0]!;
    for (const [payload, position] of [
      [inline, { kind: 'inline' as const, paragraphId, offset: 0 }],
      [block, { kind: 'block' as const, parentControlId: null, beforeId: null }],
    ] as const) {
      const before = editor.surface!.session.currentPackage();
      const bytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      expect(insertMath(editor, { payload, position })).toMatchObject({
        success: false,
        code: 'unsupported-operation',
        error: 'Prepare native payload storage before sharing application equations.',
      });
      expect(editor.surface!.session.currentPackage()).toBe(before);
      pair.documents.forEach((doc, index) =>
        expect(Y.encodeStateAsUpdate(doc)).toEqual(bytes[index]!)
      );
    }
  } finally {
    pair.destroy();
  }
});
