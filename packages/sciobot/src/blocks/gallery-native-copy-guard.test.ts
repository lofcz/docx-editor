import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, semanticDigest } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  nativeGalleryDocument,
  nativeGalleryDrawings,
  galleryDecodePort,
} from '../../test/support/native-gallery';
import { groupAdjacentImages } from './gallery-group';
import { readGallery } from './gallery';
import { buildBlockIndex } from './block-index';
import { duplicateBlock, duplicateBlockSelection } from './duplicate';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native gallery copies refuse before allocation or history until drawing identity support exists (${clientIds})`, async () => {
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
      const root = buildBlockIndex(pkg).byId.get(grouped.blockId)!;
      const before = canonicalOoxmlFingerprint(surface.session.part());
      const digest = semanticDigest(pkg.parts.values()),
        revision = surface.session.packageRevision();
      const history = { undo: surface.session.canUndo(), redo: surface.session.canRedo() };
      let allocated = 0;
      const allocateBlockId = () => `copy-${++allocated}`;
      for (const controlId of [root.controlId, gallery.images[0]!.controlId]) {
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
