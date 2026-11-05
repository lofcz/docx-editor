import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor } from '../docx-editor.ts';
import type { OoxmlNode } from '../../store/index.ts';
import {
  CT_NS,
  DRAWING_NS,
  IMG_REL,
  OD_REL,
  PNG_1X1,
  REL_NS,
  inlinePicture,
  deferredDecodePort,
  settle,
} from './image-decode-harness.ts';

function imageDocument() {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OD_REL}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL_NS}"><Relationship Id="rIdImg" Type="${IMG_REL}" Target="media/image.png"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document ${DRAWING_NS}><w:body><w:p>${inlinePicture(1)}</w:p><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`
    ),
    'word/media/image.png': PNG_1X1,
  });
}

function drawing(node: OoxmlNode): OoxmlNode | undefined {
  if (node.kind === 'drawing') return node;
  if (node.kind !== 'textValue')
    for (const child of node.children) {
      const found = drawing(child);
      if (found) return found;
    }
}

for (const deselect of [false, true])
  test(`image decoding notifies selected image chrome without moving focus; deselect=${deselect}`, async () => {
    const decode = deferredDecodePort();
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: imageDocument(),
      imageDecodePort: decode.port,
    });
    const field = document.createElement('input');
    container.append(field);
    try {
      const surface = editor.surface!;
      const id = drawing(surface.session.part().root)!.id;
      const paragraphId = surface.session.paragraphIds()[0]!;
      expect(surface.selectDrawing(id, paragraphId)).toBe(true);
      expect(editor.getSelectedImage()?.resourceStatus).toBe('pending');
      if (deselect) {
        const textId = surface.session.paragraphIds()[1]!;
        surface.setSelection({
          anchor: { paragraphId: textId, offset: 1 },
          head: { paragraphId: textId, offset: 1 },
        });
        expect(editor.getSelectedImage()).toBeNull();
      }
      field.focus();
      const before = surface.session.currentPackage();
      const revision = surface.session.packageRevision();
      const undo = editor.can({ type: 'undo' });
      let documentChanges = 0;
      const offChange = editor.on('change', () => {
        documentChanges++;
      });
      const statuses: (string | null)[] = [];
      const off = editor.on('selectionChange', (snapshot) =>
        statuses.push(snapshot.image?.resourceStatus ?? null)
      );
      decode.release();
      await settle();
      expect(statuses.includes('ready')).toBe(!deselect);
      expect(document.activeElement).toBe(field);
      expect(editor.getSelectedImage()?.id ?? null).toBe(deselect ? null : id);
      expect(surface.session.currentPackage()).toBe(before);
      expect(surface.session.packageRevision()).toBe(revision);
      expect(editor.can({ type: 'undo' })).toEqual(undo);
      expect(documentChanges).toBe(0);
      expect(editor.snapshot().image?.resourceStatus ?? null).toBe(deselect ? null : 'ready');
      off();
      offChange();
    } finally {
      editor.destroy();
      container.remove();
    }
  });
