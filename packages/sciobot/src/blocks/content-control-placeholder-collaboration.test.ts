import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  parentNodeOf,
} from '@docx-editor.dev/core/store';
import { docx, p, sdt } from '../../../core/src/automation/__tests__/support/protocol';
import { createTwoEditorFixture } from '../../test/support/collaboration';

for (const height of [0, 24])
  test(`local headers retain concurrent answer edits, presence, history, deletion, and reopen (height=${height})`, async () => {
    const document = docx(sdt(p('Answer'), p('Details')) + p('After'));
    const pair = await createTwoEditorFixture({ document, clientIds: [41, 73] });
    try {
      const [alice, bob] = pair.editors;
      const a = alice!.surface!,
        b = bob!.surface!;
      const idOf = (text: string) =>
        a.session.paragraphIds().find((id) => paragraphTextOf(a.session.part(), id) === text)!;
      const answer = idOf('Answer'),
        details = idOf('Details'),
        after = idOf('After');
      const owner = parentNodeOf(a.session.part(), parentNodeOf(a.session.part(), answer)!.id)!.id;
      const before = canonicalOoxmlFingerprint(a.session.part());
      const localPoint = { paragraphId: after, offset: 5 };
      a.setSelection({ anchor: localPoint, head: localPoint });
      const fold = () => {
        expect(
          a.setCollapsedContentControls(new Map([[owner, { placeholderHeightPt: height }]]))
        ).toEqual({ ok: true });
        a.flushPendingInput();
      };
      const check = () => {
        expect(a.collapsedContentControls()).toEqual([
          { controlId: owner, summaryParagraphId: null, placeholderHeightPt: height },
        ]);
        expect(b.collapsedContentControls()).toEqual([]);
        expect(caretAt(a.layout(), { paragraphId: details, offset: 1 })).toBeNull();
        expect(caretAt(b.layout(), { paragraphId: details, offset: 1 })).not.toBeNull();
        expect(pair.containers[0]!.textContent).not.toContain('Answer');
      };
      const point = { paragraphId: details, offset: 1 };
      b.setSelection({ anchor: point, head: point });
      await pair.settle();
      expect(pair.containers[0]!.querySelector('.docx-remote-caret')).not.toBeNull();
      fold();
      await pair.settle();
      check();
      expect(pair.containers[0]!.querySelector('.docx-remote-caret')).toBeNull();
      expect(canonicalOoxmlFingerprint(a.session.part())).toBe(before);
      expect(a.state().canUndo).toBe(false);
      pair.pause();
      expect(
        a.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: after, offset: 5, text: ' local' },
        ]).committed
      ).toBe(true);
      expect(
        b.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: answer, offset: 6, text: ' peer' },
        ]).committed
      ).toBe(true);
      pair.reconnect(1);
      await pair.settle();
      check();
      expect(paragraphTextOf(a.session.part(), answer)).toBe('Answer peer');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      check();
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = await createTwoEditorFixture({
        document,
        roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
      });
      try {
        expect(await reopened.savedDigests()).toEqual(saved);
        expect(
          reopened.editors.map((editor) => editor.surface!.collapsedContentControls())
        ).toEqual([[], []]);
      } finally {
        reopened.destroy();
      }
      a.setCollapsedContentControls(new Map());
      a.flushPendingInput();
      expect(pair.containers[0]!.querySelector('.docx-remote-caret')).not.toBeNull();
      fold();
      pair.pause();
      expect(
        a.applyAutomationOps(() => [
          { op: 'removeContentControl', controlId: owner, keepContent: false },
        ]).committed
      ).toBe(true);
      expect(
        b.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: answer, offset: 0, text: 'Remote ' },
        ]).committed
      ).toBe(true);
      pair.reconnect(0);
      await pair.settle();
      expect(a.collapsedContentControls()).toEqual([]);
      expect(a.session.paragraphIds()).not.toContain(answer);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(paragraphTextOf(a.session.part(), answer)).toBe('Remote Answer peer');
      const finalSaved = await pair.savedDigests();
      expect(finalSaved[0]).toEqual(finalSaved[1]);
    } finally {
      pair.destroy();
    }
  });
