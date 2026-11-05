import {
  escapeXml,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { JsonObject } from '../../src/data/json';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import { prepareNativeDocument } from '../../src/native/prepare-document';
import { docxFixture } from './document';

export const GALLERY_IMAGES: JsonObject[] = [
  { url: 'https://assets.example.test/same.png', caption: 'First image', width: 64, height: 32 },
  { url: 'https://assets.example.test/same.png', caption: 'Second image', width: 64, height: 32 },
  { url: 'https://assets.example.test/third.png', caption: 'Third image' },
];

export function openGallery(document: Uint8Array): OoxmlPackage {
  const opened = readOoxmlPackage(document);
  if (!opened.ok) throw Error(opened.reason);
  return opened.package;
}

export function legacyGallery(images = GALLERY_IMAGES): Uint8Array {
  const body =
    '<w:sdt><w:sdtPr><w:tag w:val="scio:1:gallery"/></w:sdtPr><w:sdtContent>' +
    images
      .map(
        (image) =>
          `<w:p><w:r><w:t>${escapeXml(String(image.caption || image.url))}</w:t></w:r></w:p>`
      )
      .join('') +
    '</w:sdtContent></w:sdt><w:p><w:r><w:t>Tail</w:t></w:r></w:p>';
  const stored = withPayloadRecord(
    openGallery(docxFixture(body)),
    {
      id: 'gallery',
      version: 1,
      type: 'image_gallery',
      attributes: { layout: 'row-fit', images },
    },
    'Gallery'
  );
  if (!stored.ok) throw Error(stored.reason);
  return writeOoxmlPackage(stored.pkg);
}

export function galleryDocument(): Uint8Array {
  const prepared = prepareNativeDocument(legacyGallery());
  if (!prepared.success) throw Error(prepared.error);
  return prepared.document;
}
