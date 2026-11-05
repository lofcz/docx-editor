import { expect, test } from 'bun:test';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import { buildBlockIndex } from './block-index';
import { insertMathAtSelection } from './math-edit';

for (const kind of ['inline', 'block'] as const)
  test(`${kind} math replaces a cross-paragraph selection in one shared history unit and returns the caret`, async () => {
    const prepared = prepareNativeDocument(
      docxFixture(
        ['Before formula after', 'Second sentence', 'Tail']
          .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
          .join('')
      )
    );
    if (!prepared.success) throw Error(prepared.error);
    const pair = await createTwoEditorFixture({ document: prepared.document });
    try {
      const [alice, bob] = pair.editors;
      const surface = alice!.surface!;
      const ids = surface.session.paragraphIds();
      const selection = {
        anchor: { paragraphId: ids[0]!, offset: 7 },
        head: { paragraphId: ids[1]!, offset: 6 },
      };
      surface.setSelection(selection);
      pair.pause();
      expect(
        insertMathAtSelection(alice!, {
          id: 'math',
          kind,
          source: 'x+1',
          selection,
          expectedRevision: surface.session.packageRevision(),
        }).success
      ).toBe(true);
      const caret = surface.state().selection.head;
      expect(caret.offset).toBe(kind === 'inline' ? 10 : 0);
      expect(paragraphTextOf(surface.session.part(), caret.paragraphId)).toBe(
        kind === 'inline' ? 'Before x+1 sentence' : ' sentence'
      );
      const peer = bob!.surface!;
      expect(
        peer.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: ids[2]!, offset: 4, text: ' peer' },
        ]).committed
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      const texts = (index: number) => {
        const current = pair.editors[index]!.surface!;
        return current.session
          .paragraphIds()
          .map((id) => paragraphTextOf(current.session.part(), id));
      };
      const expected =
        kind === 'inline'
          ? ['Before x+1 sentence', 'Tail peer']
          : ['Before ', 'x+1', ' sentence', 'Tail peer'];
      for (const index of [0, 1]) expect(texts(index)).toEqual(expected);
      const controlId = buildBlockIndex(surface.session.currentPackage()).byId.get(
        'math'
      )!.controlId;
      expect(buildBlockIndex(peer.session.currentPackage()).byId.get('math')!.controlId).toBe(
        controlId
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const index of [0, 1])
        expect(texts(index)).toEqual(['Before formula after', 'Second sentence', 'Tail peer']);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const index of [0, 1]) expect(texts(index)).toEqual(expected);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
