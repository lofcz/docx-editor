// The shared fixture uses the EigenPal Pro License; production paste uses Apache core.
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { inlineWidgetDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { collectBlockClipboardData, pasteWithBlocks } from './clipboard';

function copySource(editor: ReturnType<typeof createDocxEditor>) {
  const surface = editor.surface!,
    paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({ anchor: { paragraphId, offset: 7 }, head: { paragraphId, offset: 13 } });
  const flavours = surface.copyFlavours(),
    data = collectBlockClipboardData(editor)!;
  const end = { paragraphId, offset: 19 };
  surface.setSelection({ anchor: end, head: end });
  return { text: flavours.text, html: flavours.html, data };
}

test('upstream tracked rich paste and private payload repair share one exact undo unit', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: inlineWidgetDocument(),
    author: 'Ada',
  });
  try {
    const surface = editor.surface!,
      transfer = copySource(editor);
    const digest = () => semanticDigest(surface.session.currentPackage().parts.values());
    const before = digest();
    surface.setEditingMode('suggest');
    expect(pasteWithBlocks(editor, transfer, { createId: () => 'tracked-copy' })).toBe(true);
    const pasted = digest();
    expect(
      buildBlockIndex(surface.session.currentPackage()).byId.get('tracked-copy')!.payload.attributes
    ).toEqual({ name: 'Initial' });
    expect(serializeOoxmlPart(surface.session.part())).toMatch(/<w:ins\b[^>]*w:author="Ada"/);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest()).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(digest()).toEqual(pasted);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(semanticDigest(reopened.package.parts.values())).toEqual(pasted);
  } finally {
    editor.destroy();
  }
});

test('tracked payload paste preserves concurrent native peer input through reconnect and history', async () => {
  const pair = await createTwoEditorFixture({
    document: inlineWidgetDocument(),
    clientIds: [73, 41],
  });
  try {
    const [alice, bob] = pair.editors,
      a = alice!.surface!,
      b = bob!.surface!;
    const id = a.session.paragraphIds()[0]!,
      transfer = copySource(alice!);
    a.setEditingMode('suggest');
    pair.pause();
    expect(pasteWithBlocks(alice!, transfer, { createId: () => 'tracked-copy' })).toBe(true);
    const point = { paragraphId: id, offset: 0 };
    b.setSelection({ anchor: point, head: point });
    b.insertPlainText('Bob ');
    b.flushPendingInput();
    const peerOnly = semanticDigest(b.session.currentPackage().parts.values());
    pair.reconnect(0);
    await pair.settle();
    for (const surface of [a, b]) {
      expect(paragraphTextOf(surface.session.part(), id)).toBe('Bob Before Source afterSource');
      expect(
        buildBlockIndex(surface.session.currentPackage()).byId.get('tracked-copy')!.payload
          .attributes
      ).toEqual({ name: 'Initial' });
    }
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const surface of [a, b]) {
      expect(paragraphTextOf(surface.session.part(), id)).toBe('Bob Before Source after');
      expect(buildBlockIndex(surface.session.currentPackage()).byId.has('tracked-copy')).toBe(
        false
      );
      expect(semanticDigest(surface.session.currentPackage().parts.values())).toEqual(peerOnly);
    }
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    for (const surface of [a, b])
      expect(paragraphTextOf(surface.session.part(), id)).toBe('Bob Before Source afterSource');
    const saved = await pair.savedDigests();
    expect(saved[0]).toEqual(saved[1]);
  } finally {
    pair.destroy();
  }
});
