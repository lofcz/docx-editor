import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  createNodeIdAllocator,
  findNode,
  readOoxmlPart,
  replaceChildren,
  semanticDigest,
  withPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  nativeGalleryDocument,
  nativeGalleryDrawings,
  galleryDecodePort,
  galleryDescendants,
} from '../../test/support/native-gallery';
import { groupAdjacentImages } from './gallery-group';
import { readGallery } from './gallery';
import { duplicateBlock, duplicateBlockSelection } from './duplicate';
import { withFreshIds } from './fresh-ids';
import { isNativeGalleryCaption } from './gallery-native-record';

test('registered native gallery copies refuse authored reference ownership before allocating payloads', async () => {
  const pair = await createTwoEditorFixture({
    document: await nativeGalleryDocument(),
    ownedAttributes: true,
    editor: { imageDecodePort: galleryDecodePort },
  });
  try {
    const editor = pair.editors[0]!,
      surface = editor.surface!;
    const grouped = groupAdjacentImages(editor, {
      drawingNodeId: nativeGalleryDrawings(surface.session.currentPackage())[0]!.drawing.id,
    });
    if (!grouped.success) throw Error(grouped.error);
    const gallery = readGallery(surface.session.currentPackage(), grouped.blockId)!;
    const caption = galleryDescendants(findNode(surface.session.part(), gallery.controlId)!).find(
      (node) => node.kind === 'paragraph' && isNativeGalleryCaption(surface.session.part(), node)
    )!;
    const source = readOoxmlPart(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:bookmarkStart w:id="71" w:name="CaptionReference"/><w:bookmarkEnd w:id="71"/></w:p></w:body></w:document>',
      {
        name: '/word/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      }
    );
    if (!source.ok) throw Error(source.reason);
    const body = bodyStoryRoot(source.part)!;
    if (body.kind === 'textValue') throw Error('Missing reference body');
    const paragraph = body.children[0]!;
    if (paragraph.kind !== 'paragraph') throw Error('Missing reference paragraph');
    expect(
      surface.applyAutomationOps(
        () => [],
        { kind: 'body' },
        [
          (working) => {
            const part = working.parts.get(working.mainDocumentPart)!,
              current = findNode(part, caption.id)!;
            if (current.kind !== 'paragraph') throw Error('Missing native caption');
            const nextId = createNodeIdAllocator(part, 'new');
            const markers = paragraph.children.map((node) =>
              withFreshIds(node, nextId, { prefix: null, used: new Set(), seed: 'reference' })
            );
            const properties = current.children.filter(
              (node) => node.kind === 'paragraphProperties'
            );
            const content = current.children.filter((node) => node.kind !== 'paragraphProperties');
            const changed = replaceChildren(part, current.id, [
              ...properties,
              markers[0]!,
              ...content,
              markers[1]!,
            ]);
            if (!changed.ok) throw Error(JSON.stringify(changed.issues));
            return withPart(working, changed.part);
          },
        ],
        false,
        true
      ).committed
    ).toBe(true);
    await pair.settle();
    const revision = surface.session.packageRevision(),
      before = canonicalOoxmlFingerprint(surface.session.part());
    const history = { undo: surface.session.canUndo(), redo: surface.session.canRedo() };
    let allocated = 0;
    const allocateBlockId = () => `protected-copy-${++allocated}`;
    expect(duplicateBlock(editor, { controlId: gallery.controlId, allocateBlockId })).toMatchObject(
      { success: false, code: 'unsupported-operation' }
    );
    expect(
      duplicateBlockSelection(editor, { nodeIds: [gallery.controlId], allocateBlockId })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(allocated).toBe(0);
    expect(surface.session.packageRevision()).toBe(revision);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
    expect({ undo: surface.session.canUndo(), redo: surface.session.canRedo() }).toEqual(history);
    const saved = await pair.savedDigests();
    expect(saved[0]).toEqual(saved[1]);
  } finally {
    pair.destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native gallery item copies retain their owner and refuse before allocation or history (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({
      document: await nativeGalleryDocument(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const editor = pair.editors[0]!;
      const drawing = nativeGalleryDrawings(editor.surface!.session.currentPackage())[0]!.drawing
        .id;
      const grouped = groupAdjacentImages(editor, { drawingNodeId: drawing });
      if (!grouped.success) throw Error(grouped.error);
      await pair.settle();
      const surface = editor.surface!,
        pkg = surface.session.currentPackage();
      const gallery = readGallery(pkg, grouped.blockId)!;
      const before = canonicalOoxmlFingerprint(surface.session.part());
      const digest = semanticDigest(pkg.parts.values()),
        revision = surface.session.packageRevision();
      const history = { undo: surface.session.canUndo(), redo: surface.session.canRedo() };
      let allocated = 0;
      const allocateBlockId = () => `copy-${++allocated}`;
      for (const controlId of [gallery.images[0]!.controlId]) {
        expect(duplicateBlock(editor, { controlId, allocateBlockId })).toMatchObject({
          success: false,
          code: 'unsupported-operation',
        });
        expect(
          duplicateBlockSelection(editor, { nodeIds: [controlId], allocateBlockId })
        ).toMatchObject({ success: false, code: 'unsupported-operation' });
      }
      expect(allocated).toBe(0);
      expect(surface.session.packageRevision()).toBe(revision);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
      expect(semanticDigest(surface.session.currentPackage().parts.values())).toEqual(digest);
      expect({ undo: surface.session.canUndo(), redo: surface.session.canRedo() }).toEqual(history);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
