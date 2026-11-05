import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlTextOf,
  findNode,
  paragraphTextOf,
  parentNodeOf,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { importedMath, mathDocument } from '../../test/support/math-document';
import { prepareNativeDocument } from '../native/prepare-document';
import { removeBlock } from './remove';
import { mathWordFallback } from './math-codec';
import { buildBlockIndex } from './block-index';
import { insertMathAtSelection, updateMath } from './math-edit';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`math conversion retains peer text, identity, undo, cursors and save/reopen (${clients})`, async () => {
    const prepared = prepareNativeDocument(mathDocument());
    if (!prepared.success) throw Error(prepared.error);
    const pair = await createTwoEditorFixture({ document: prepared.document, clientIds: clients });
    try {
      const [alice, bob] = pair.editors;
      const a = alice!.surface!;
      const first = a.session.paragraphIds()[0]!;
      const point = { paragraphId: first, offset: 3 };
      const sourceOmml = importedMath(a.session.currentPackage());
      expect(
        insertMathAtSelection(alice!, {
          id: 'math',
          source: 'x',
          kind: 'inline',
          selection: { anchor: point, head: point },
          expectedRevision: a.session.packageRevision(),
        }).success
      ).toBe(true);
      await pair.settle();
      const block = (index = 0) =>
        buildBlockIndex(pair.editors[index]!.surface!.session.currentPackage()).byId.get('math')!;
      const original = block().payload;
      const controlId = block().controlId;
      pair.pause();
      const conversion = updateMath(alice!, { original, source: 'y', kind: 'block' });
      expect(conversion).toMatchObject({ success: true });
      const b = bob!.surface!;
      const last = first;
      const end = { paragraphId: last, offset: paragraphTextOf(b.session.part(), last)!.length };
      b.setSelection({ anchor: end, head: end });
      b.insertPlainText(' peer');
      b.flushPendingInput();
      pair.reconnect();
      await pair.settle();
      const verify = async (kind: 'inline' | 'block', source: string) => {
        for (const [index, editor] of pair.editors.entries()) {
          const surface = editor.surface!;
          expect(block(index).payload.attributes.texExpression).toBe(source);
          expect(
            contentControlTextOf(findNode(surface.session.part(), block(index).controlId)!)
          ).toBe(source.split(/\r\n|\r|\n/).join(' '));
          expect(block(index).controlId).toBe(controlId);
          expect(parentNodeOf(surface.session.part(), block(index).controlId)?.kind).toBe(
            kind === 'inline' ? 'paragraph' : 'body'
          );
          expect(importedMath(surface.session.currentPackage())).toEqual(sourceOmml);
          expect(
            surface.session
              .paragraphIds()
              .map((id) => paragraphTextOf(surface.session.part(), id))
              .some((text) => text?.endsWith(' peer'))
          ).toBe(true);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      };
      await verify('block', 'y');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      await verify('inline', 'x');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await verify('block', 'y');
      const demotion = updateMath(alice!, {
        original: block().payload,
        source: 'z',
        kind: 'inline',
      });
      expect(demotion).toMatchObject({ success: true });
      await pair.settle();
      await verify('inline', 'z');
      const stale = block().payload;
      expect(
        updateMath(bob!, {
          original: block(1).payload,
          source: 'peer\r\n\r\nequation',
          kind: 'inline',
        }).success
      ).toBe(true);
      await pair.settle();
      const unchanged = a.session.currentPackage();
      expect(updateMath(alice!, { original: stale, source: 'draft', kind: 'block' })).toMatchObject(
        { success: false, code: 'stale-revision' }
      );
      expect(a.session.currentPackage()).toBe(unchanged);
      const caretId = b.session.paragraphIds().at(-1)!;
      const caret = { paragraphId: caretId, offset: 2 };
      b.setSelection({ anchor: caret, head: caret });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (value) => value.actorId === 'bob'
      );
      expect(remote?.head.nodeId).toBe(caretId);
      const geometries = pair.editors.map((editor) => caretAt(editor.surface!.layout(), caret));
      expect(geometries[0]).not.toBeNull();
      expect(geometries[0]!.x).toBe(geometries[1]!.x);
      expect(geometries[0]!.y).toBe(geometries[1]!.y);
      for (const shared of [false, true]) {
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice!.save()),
          ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
        });
        try {
          await reopened.settle();
          const digests = await reopened.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          for (const editor of reopened.editors) {
            expect(
              buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('math')!.payload
            ).toEqual(block().payload);
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(a.session.part())
            );
          }
        } finally {
          reopened.destroy();
        }
      }
      // Each conversion direction owns exactly one history unit.
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(block().payload.type).toBe('equation');
      expect(block().payload.attributes.texExpression).toBe('peer\r\n\r\nequation');
      await verify('block', 'peer\r\n\r\nequation');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(block().payload.type).toBe('inline_equation');
      await verify('inline', 'peer\r\n\r\nequation');
      const draft = block().payload;
      const deleted = removeBlock(bob!, { controlId: block(1).controlId, mode: 'delete' });
      expect(deleted.success).toBe(true);
      await pair.settle();
      expect(
        updateMath(alice!, { original: draft, source: 'draft', kind: 'inline' })
      ).toMatchObject({ success: false, code: 'stale-revision' });
    } finally {
      pair.destroy();
    }
  });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`equation creation undo retains peer source through redo (${clients})`, async () => {
    const prepared = prepareNativeDocument(mathDocument());
    if (!prepared.success) throw Error(prepared.error);
    const pair = await createTwoEditorFixture({ document: prepared.document, clientIds: clients });
    try {
      const [alice, bob] = pair.editors;
      const a = alice!.surface!;
      const point = { paragraphId: a.session.paragraphIds()[0]!, offset: 3 };
      expect(
        insertMathAtSelection(alice!, {
          id: 'created',
          source: 'original',
          kind: 'block',
          selection: { anchor: point, head: point },
          expectedRevision: a.session.packageRevision(),
        }).success
      ).toBe(true);
      await pair.settle();
      const block = (index: number) =>
        buildBlockIndex(pair.editors[index]!.surface!.session.currentPackage()).byId.get('created');
      const controlId = block(0)!.controlId;
      expect(
        updateMath(bob!, { original: block(1)!.payload, source: 'peer\nsource', kind: 'block' })
          .success
      ).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(block(0)).toEqual(block(1));
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const [index, editor] of pair.editors.entries()) {
        expect(block(index)?.controlId).toBe(controlId);
        expect(block(index)?.payload.attributes.texExpression).toBe('peer\nsource');
        expect(contentControlTextOf(findNode(editor.surface!.session.part(), controlId)!)).toBe(
          'peer source'
        );
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });

// This exact Word fallback oracle remains active until conversion history preserves peer content.
test('conversion undo preserves a later peer source in the canonical Word fallback', async () => {
  const prepared = prepareNativeDocument(mathDocument());
  if (!prepared.success) throw Error(prepared.error);
  const pair = await createTwoEditorFixture({ document: prepared.document });
  try {
    const [alice, bob] = pair.editors;
    const a = alice!.surface!;
    const point = { paragraphId: a.session.paragraphIds()[0]!, offset: 3 };
    expect(
      insertMathAtSelection(alice!, {
        id: 'math',
        source: 'x',
        kind: 'inline',
        selection: { anchor: point, head: point },
        expectedRevision: a.session.packageRevision(),
      }).success
    ).toBe(true);
    await pair.settle();
    const block = (index = 0) =>
      buildBlockIndex(pair.editors[index]!.surface!.session.currentPackage()).byId.get('math')!;
    expect(
      updateMath(alice!, { original: block().payload, source: 'y', kind: 'block' })
    ).toMatchObject({ success: true });
    await pair.settle();
    expect(
      updateMath(alice!, { original: block().payload, source: 'z', kind: 'inline' })
    ).toMatchObject({ success: true });
    await pair.settle();
    expect(
      updateMath(bob!, {
        original: block(1).payload,
        source: 'peer\r\n\r\nequation',
        kind: 'inline',
      }).success
    ).toBe(true);
    await pair.settle();
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const [index, editor] of pair.editors.entries()) {
      const current = block(index);
      expect(current.payload.attributes.texExpression).toBe('peer\r\n\r\nequation');
      expect(
        contentControlTextOf(findNode(editor.surface!.session.part(), current.controlId)!)
      ).toBe('peer  equation');
    }
  } finally {
    pair.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`concurrent whole math replacements keep source and fallback together (${clients})`, async () => {
    const prepared = prepareNativeDocument(mathDocument());
    if (!prepared.success) throw Error(prepared.error);
    const pair = await createTwoEditorFixture({ document: prepared.document, clientIds: clients });
    try {
      const [alice, bob] = pair.editors;
      const surface = alice!.surface!;
      const point = { paragraphId: surface.session.paragraphIds()[0]!, offset: 3 };
      expect(
        insertMathAtSelection(alice!, {
          id: 'math',
          source: 'x',
          kind: 'block',
          selection: { anchor: point, head: point },
          expectedRevision: surface.session.packageRevision(),
        }).success
      ).toBe(true);
      await pair.settle();
      const block = (index = 0) =>
        buildBlockIndex(pair.editors[index]!.surface!.session.currentPackage()).byId.get('math')!;
      pair.pause();
      expect(
        updateMath(alice!, { original: block().payload, source: 'alice', kind: 'block' }).success
      ).toBe(true);
      expect(
        updateMath(bob!, { original: block(1).payload, source: 'bob\r\n\r\nsecond', kind: 'block' })
          .success
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const [index, editor] of pair.editors.entries()) {
        const current = block(index);
        expect(current.payload).toEqual(block().payload);
        const expected = mathWordFallback(
          current.payload as Parameters<typeof mathWordFallback>[0]
        ).paragraphs.join('');
        expect(
          contentControlTextOf(findNode(editor.surface!.session.part(), current.controlId)!)
        ).toBe(expected);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const [index, editor] of pair.editors.entries()) {
        expect(block(index).payload.attributes.texExpression).toBe('bob\r\n\r\nsecond');
        expect(
          contentControlTextOf(findNode(editor.surface!.session.part(), block(index).controlId)!)
        ).toBe('bob  second');
      }
      for (const round of [1, 2, 3]) {
        pair.pause();
        expect(
          updateMath(alice!, {
            original: block().payload,
            source: `Alice repeated ${round}`,
            kind: 'block',
          }).success
        ).toBe(true);
        expect(
          updateMath(bob!, {
            original: block(1).payload,
            source: `Bob repeated ${round}`,
            kind: 'block',
          }).success
        ).toBe(true);
        pair.reconnect();
        await pair.settle();
        const winner = block().payload.attributes.texExpression;
        if (typeof winner !== 'string') throw Error('Missing equation source');
        for (const [index, editor] of pair.editors.entries()) {
          expect(block(index).payload.attributes.texExpression).toBe(winner);
          expect(
            contentControlTextOf(findNode(editor.surface!.session.part(), block(index).controlId)!)
          ).toBe(winner);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const [index, editor] of pair.editors.entries()) {
          expect(block(index)?.payload.attributes.texExpression).toBe(`Bob repeated ${round}`);
          expect(
            contentControlTextOf(findNode(editor.surface!.session.part(), block(index).controlId)!)
          ).toBe(`Bob repeated ${round}`);
        }
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const [index, editor] of pair.editors.entries()) {
          expect(block(index)?.payload.attributes.texExpression).toBe(winner);
          expect(
            contentControlTextOf(findNode(editor.surface!.session.part(), block(index).controlId)!)
          ).toBe(winner);
        }
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
