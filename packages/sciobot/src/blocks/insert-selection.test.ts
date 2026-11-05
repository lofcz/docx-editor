import { expect, test } from 'bun:test';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from './block-index';
import { insertBlockAtSelection } from './insert';
import { withInitializedPayloadStore } from './payload-opc-codec';

function fixture() {
  const opened = readOoxmlPackage(
    docxFixture(
      ['Before CCO after', 'Second sentence', 'Tail']
        .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
        .join('')
    )
  );
  if (!opened.ok) throw Error(opened.reason);
  const ready = withInitializedPayloadStore(opened.package);
  if (!ready.ok) throw Error(ready.reason);
  return writeOoxmlPackage(ready.pkg);
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const mode of ['collapsed', 'same-paragraph', 'cross-paragraph'] as const)
    test(`atomic insertion replaces ${mode} selection in one shared history unit ${clientIds}`, async () => {
      const pair = await createTwoEditorFixture({ document: fixture(), clientIds });
      try {
        const [alice, bob] = pair.editors;
        const surface = alice!.surface!;
        const ids = surface.session.paragraphIds();
        const anchor = { paragraphId: ids[0]!, offset: 7 };
        const head =
          mode === 'cross-paragraph'
            ? { paragraphId: ids[1]!, offset: 6 }
            : { paragraphId: ids[0]!, offset: mode === 'collapsed' ? 7 : 10 };
        surface.setSelection({ anchor, head });
        pair.pause();
        const result = insertBlockAtSelection(alice!, {
          payload: {
            version: 1,
            id: 'new-molecule',
            type: 'smiles',
            attributes: { smiles: 'CCO' },
          },
          label: 'Chemical structure',
          paragraphs: ['CCO'],
          selection: { anchor, head },
          expectedRevision: surface.session.packageRevision(),
        });
        expect(result.success).toBe(true);
        if (!result.success) throw Error(result.error);
        const controlId = result.controlId;
        expect(surface.state().selection).toEqual({ anchor: result.caret, head: result.caret });
        expect(paragraphTextOf(surface.session.part(), result.caret.paragraphId)).not.toBeNull();
        const tail = bob!.surface!.session.paragraphIds()[2]!;
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: tail, offset: 4, text: ' peer' },
          ]).committed
        ).toBe(true);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const texts = (index: number) => {
          const current = pair.editors[index]!.surface!;
          return current.session
            .paragraphIds()
            .map((id) => paragraphTextOf(current.session.part(), id));
        };
        const expected =
          mode === 'cross-paragraph'
            ? ['Before ', 'CCO', ' sentence', 'Tail peer']
            : [
                'Before ',
                'CCO',
                mode === 'collapsed' ? 'CCO after' : ' after',
                'Second sentence',
                'Tail peer',
              ];
        for (const index of [0, 1]) {
          expect(
            buildBlockIndex(pair.editors[index]!.surface!.session.currentPackage()).byId.get(
              'new-molecule'
            )!.controlId
          ).toBe(controlId);
          expect(texts(index)).toEqual(expected);
          expect(
            buildBlockIndex(pair.editors[index]!.surface!.session.currentPackage()).byId.get(
              'new-molecule'
            )?.payload.attributes.smiles
          ).toBe('CCO');
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const index of [0, 1])
          expect(texts(index)).toEqual(['Before CCO after', 'Second sentence', 'Tail peer']);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const index of [0, 1]) expect(texts(index)).toEqual(expected);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        const before = saved[0];
        expect(
          insertBlockAtSelection(alice!, {
            payload: {
              version: 1,
              id: 'stale-molecule',
              type: 'smiles',
              attributes: { smiles: 'CCN' },
            },
            label: 'Chemical structure',
            selection: { anchor, head },
            expectedRevision: -1,
          })
        ).toMatchObject({ success: false, code: 'stale-revision' });
        expect((await pair.savedDigests())[0]).toEqual(before);
      } finally {
        pair.destroy();
      }
    });
