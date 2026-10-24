// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { listDocument, listParagraph } from '../../test/support/list-document';
import {
  setManualList,
  setManualListChecked,
  adjustManualListLevel,
  manualListState,
} from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['lower-roman', 'todo', 'check'] as const)
    test(`manual ${action} retains offline text, independent history, reconnect, geometry, deletion recovery and Word convergence ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: listDocument(
          listParagraph('Alpha') + listParagraph('Imported', '3') + listParagraph('Tail')
        ),
        documentId: context.documentId,
        ownedAttributes: true,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          id = a.session.paragraphIds()[0]!;
        const select = (s: typeof a, offset: number) => {
          const point = { paragraphId: id, offset };
          s.setSelection({ anchor: point, head: point });
        };
        select(a, 2);
        if (action === 'check') {
          expect(setManualList(alice!, 'todo', context).success).toBe(true);
          await pair.settle();
        }
        const before = a.session.part(),
          offset = paragraphTextOf(before, id)!.length;
        select(b, offset);
        pair.pause();
        expect(
          (action === 'check'
            ? setManualListChecked(alice!, true, context)
            : setManualList(alice!, action, context)
          ).success
        ).toBe(true);
        expect(adjustManualListLevel(alice!, 'increase', context).success).toBe(true);
        const authored = a.session.part();
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const expectedOffset = offset + (action === 'todo' ? 1 : 0);
        const peer = applyTreeOp(authored, {
          op: 'insertText',
          paragraphId: id,
          offset: expectedOffset,
          text: '!',
        });
        if (!peer.ok) throw new Error(peer.reason);
        const after = canonicalOoxmlFingerprint(peer.part);
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(after);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(
          manualListState(alice!, context)!.checkbox?.level ??
            manualListState(alice!, context)!.numbering?.level
        ).toBe(0);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const preserved = applyTreeOp(before, {
          op: 'insertText',
          paragraphId: id,
          offset,
          text: '!',
        });
        if (!preserved.ok) throw new Error(preserved.reason);
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(preserved.part)
          );
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (let i = 0; i < 2; i++) {
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
        }
        const selection = {
          anchor: { paragraphId: id, offset: 1 },
          head: { paragraphId: id, offset: 4 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const e of pair.editors) {
          expect(caretAt(e.surface!.layout(), selection.head)).not.toBeNull();
          expect(
            selectionRects(e.surface!.layout(), selection, e.surface!.session.paragraphIds()).length
          ).toBeGreaterThan(0);
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(after);
        }
        const beforeDeletion = a.session.part();
        pair.pause();
        expect(a.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).committed).toBe(
          true
        );
        select(b, expectedOffset);
        b.insertPlainText('?');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const recovered = applyTreeOp(beforeDeletion, {
          op: 'insertText',
          paragraphId: id,
          offset: expectedOffset,
          text: '?',
        });
        if (!recovered.ok) throw new Error(recovered.reason);
        const saved = [];
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(recovered.part)
          );
          const read = readOoxmlPackage(new Uint8Array(await e.save()));
          if (!read.ok) throw new Error(read.reason);
          expect(
            canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
          ).toEqual(canonicalOoxmlFingerprint(recovered.part));
          saved.push(semanticDigest(read.package.parts.values()));
        }
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });

test('shared checklist continuation restores the new paragraph caret on redo through the application session wrapper', async () => {
  const { applyManualListInput, handleManualListKey } = await import('./list-input');
  const pair = await createTwoEditorFixture({
    document: listDocument(listParagraph('[]') + listParagraph('Tail')),
    documentId: context.documentId,
    ownedAttributes: true,
  });
  try {
    const editor = pair.editors[0]!,
      s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      point = { paragraphId: id, offset: 2 };
    s.setSelection({ anchor: point, head: point });
    expect(applyManualListInput(editor, context)).toBe(true);
    s.insertPlainText('Task');
    s.flushPendingInput();
    expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
    await pair.settle();
    const selection = s.state().selection;
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(s.state().selection).toEqual(selection);
  } finally {
    pair.destroy();
  }
});
