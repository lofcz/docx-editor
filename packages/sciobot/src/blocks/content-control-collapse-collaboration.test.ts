import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  parentNodeOf,
} from '@docx-editor.dev/core/store';
import {
  cell,
  docx,
  p,
  row,
  sdt,
  table,
} from '../../../core/src/automation/__tests__/support/protocol';
import { createTwoEditorFixture } from '../../test/support/collaboration';

for (const inCell of [false, true])
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    test(`local folds retain peer content, presence, history, deletion, and reopen (${clientIds}; cell=${inCell})`, async () => {
      const control = sdt(p('Summary'), p('Body'));
      const document = docx((inCell ? table(row(cell(control))) : control) + p('After'));
      const pair = await createTwoEditorFixture({ document, clientIds });
      try {
        const [alice, bob] = pair.editors;
        const surfaces = pair.editors.map((editor) => editor.surface!);
        const a = surfaces[0]!,
          b = surfaces[1]!;
        const idOf = (text: string) =>
          a.session.paragraphIds().find((id) => paragraphTextOf(a.session.part(), id) === text)!;
        const summary = idOf('Summary'),
          body = idOf('Body'),
          after = idOf('After');
        const owner = parentNodeOf(
          a.session.part(),
          parentNodeOf(a.session.part(), summary)!.id
        )!.id;
        const fold = () => {
          expect(a.setCollapsedContentControls(new Map([[owner, summary]]))).toEqual({ ok: true });
          a.flushPendingInput();
        };
        const checkFold = () => {
          expect(a.collapsedContentControls()).toEqual([
            { controlId: owner, summaryParagraphId: summary },
          ]);
          expect(b.collapsedContentControls()).toEqual([]);
          expect(caretAt(a.layout(), { paragraphId: body, offset: 0 })).toBeNull();
          expect(caretAt(b.layout(), { paragraphId: body, offset: 0 })).not.toBeNull();
        };
        const before = surfaces.map((surface) => canonicalOoxmlFingerprint(surface.session.part()));
        const revisions = surfaces.map((surface) => surface.session.packageRevision());
        const point = { paragraphId: body, offset: 2 };
        b.setSelection({ anchor: point, head: point });
        await pair.settle();
        expect(pair.containers[0]!.querySelector('.docx-remote-caret')).not.toBeNull();
        fold();
        await pair.settle();
        checkFold();
        expect(pair.containers[0]!.querySelector('.docx-remote-caret')).toBeNull();
        expect(
          surfaces.map((surface) => canonicalOoxmlFingerprint(surface.session.part()))
        ).toEqual(before);
        expect(surfaces.map((surface) => surface.session.packageRevision())).toEqual(revisions);
        expect(surfaces.map((surface) => surface.state().canUndo)).toEqual([false, false]);

        pair.pause();
        expect(
          a.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: after, offset: 5, text: ' local' },
          ]).committed
        ).toBe(true);
        expect(
          b.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: body, offset: 4, text: ' peer' },
          ]).committed
        ).toBe(true);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        checkFold();
        for (const surface of surfaces) {
          expect(paragraphTextOf(surface.session.part(), body)).toBe('Body peer');
          expect(paragraphTextOf(surface.session.part(), after)).toBe('After local');
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        checkFold();
        expect(paragraphTextOf(a.session.part(), body)).toBe('Body peer');
        expect(paragraphTextOf(b.session.part(), after)).toBe('After');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        checkFold();
        expect(paragraphTextOf(a.session.part(), body)).toBe('Body');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        checkFold();
        a.setCollapsedContentControls(new Map());
        a.flushPendingInput();
        expect(pair.containers[0]!.querySelector('.docx-remote-caret')).not.toBeNull();
        fold();

        // Deleting the exact summary retires its view entry while a peer changes hidden content.
        pair.pause();
        expect(
          a.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: summary }]).committed
        ).toBe(true);
        expect(
          b.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: body, offset: 0, text: 'Edited ' },
          ]).committed
        ).toBe(true);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        expect(a.collapsedContentControls()).toEqual([]);
        expect(caretAt(a.layout(), { paragraphId: body, offset: 0 })).not.toBeNull();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(a.collapsedContentControls()).toEqual([]);
        expect(paragraphTextOf(a.session.part(), summary)).toBe('Summary');
        expect(paragraphTextOf(a.session.part(), body)).toBe('Edited Body peer');
        fold();
        pair.pause();
        expect(
          a.applyAutomationOps(() => [
            { op: 'removeContentControl', controlId: owner, keepContent: false },
          ]).committed
        ).toBe(true);
        expect(
          b.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: body, offset: 0, text: 'Remote ' },
          ]).committed
        ).toBe(true);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        expect(a.collapsedContentControls()).toEqual([]);
        expect(a.session.paragraphIds()).not.toContain(body);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(a.collapsedContentControls()).toEqual([]);
        expect(paragraphTextOf(a.session.part(), body)).toBe('Remote Edited Body peer');
        fold();
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        checkFold();
        const reopened = await createTwoEditorFixture({
          document,
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        });
        try {
          expect(await reopened.savedDigests()).toEqual(saved);
          for (const editor of reopened.editors)
            expect(editor.surface!.collapsedContentControls()).toEqual([]);
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    });
