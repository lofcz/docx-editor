import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage, semanticDigest } from '@docx-editor.dev/core/store';
import { inlineWidgetDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { collectBlockClipboardData, pasteWithBlocks } from './clipboard';
import { readBlockSelection } from './selection';

test('inline selection uses canonical UTF-16 boundaries and carries payloads through native clipboard paste', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: inlineWidgetDocument() });
  try {
    const surface = editor.surface!,
      paragraphId = surface.session.paragraphIds()[0]!;
    const digest = () => semanticDigest(surface.session.currentPackage().parts.values());
    const baseline = digest();
    const select = (anchor: number, head = anchor) =>
      surface.setSelection({
        anchor: { paragraphId, offset: anchor },
        head: { paragraphId, offset: head },
      });
    for (const [anchor, head, relation] of [
      [0, 7, null],
      [13, 19, null],
      [7, 7, null],
      [13, 13, null],
      [8, 8, 'within'],
      [8, 11, 'within'],
      [7, 13, 'contains'],
      [13, 7, 'contains'],
      [6, 9, 'crosses'],
      [11, 15, 'crosses'],
      [0, 19, 'contains'],
    ] as const) {
      select(anchor, head);
      const selection = readBlockSelection(editor)!;
      expect(selection.blocks.map((block) => [block.id, block.relation])).toEqual(
        relation ? [['source-1', relation]] : []
      );
    }
    select(7, 13);
    expect(digest()).toEqual(baseline);
    const data = collectBlockClipboardData(editor)!;
    expect(data.blocks).toHaveLength(1);
    expect(data.blocks[0]!.payload.attributes).toEqual({ name: 'Initial' });
    const flavours = surface.copyFlavours();
    expect(flavours.text).toBe('Source');
    select(19);
    expect(
      pasteWithBlocks(
        editor,
        { text: flavours.text, html: flavours.html, data },
        {
          createId: () => 'source-copy',
        }
      )
    ).toBe(true);
    const index = buildBlockIndex(surface.session.currentPackage());
    expect(index.byId.get('source-copy')!.payload.attributes).toEqual({ name: 'Initial' });
    expect(index.byId.get('source-copy')!.controlId).not.toBe(
      index.byId.get('source-1')!.controlId
    );
    expect(paragraphTextOf(surface.session.part(), paragraphId)).toBe('Before Source afterSource');
    const pasted = digest();
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest()).toEqual(baseline);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(digest()).toEqual(pasted);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(semanticDigest(reopened.package.parts.values())).toEqual(pasted);
  } finally {
    editor.destroy();
    container.remove();
  }
});
