import { expect, test } from 'bun:test';
import {
  customXmlNodes,
  replaceChildren,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { withPreparedPayloadFields } from './payload-preparation';
import { readGallery } from './gallery';

function data(pkg: OoxmlPackage) {
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') throw Error('Missing gallery store');
  return store;
}
function node(pkg: OoxmlPackage, id: string) {
  return pkg.parts
    .get(data(pkg).part.partName)!
    .root.children.find(
      (child) =>
        child.kind !== 'textValue' &&
        child.attributes.some((a) => a.localName === 'id' && a.value === `scio_${id}`)
    ) as OoxmlElement;
}

test('prepared gallery image histories recover capacity without changing opaque metadata or base payload data', () => {
  const pkg = openGallery(galleryDocument());
  const id = readGallery(pkg, 'gallery')!.images[0]!.id;
  const record = node(pkg, id);
  const field: OoxmlElement = {
    id: 'caption-field',
    kind: 'generic',
    namespaceUri: record.namespaceUri,
    localName: 'attr',
    namespaceBindings: [],
    children: [],
    attributes: [
      { kind: 'genericExtension', namespaceUri: '', localName: 'name', value: 'caption' },
      {
        kind: 'genericExtension',
        namespaceUri: '',
        localName: 'value',
        value: JSON.stringify('First image'),
      },
    ],
  };
  const opaque: OoxmlElement = {
    ...field,
    id: 'opaque-field',
    attributes: [
      ...field.attributes,
      { kind: 'genericExtension', namespaceUri: '', localName: 'opaque', value: 'keep' },
    ],
  };
  const part = pkg.parts.get(data(pkg).part.partName)!;
  const replaced = replaceChildren(part, record.id, [
    ...record.children,
    opaque,
    ...Array.from({ length: 9_999 }, (_, at) => ({ ...field, id: `history-${at}` })),
  ]);
  if (!replaced.ok) throw Error('Cannot construct gallery field history');
  const full = withPart(pkg, replaced.part);
  const decoded = data(full).entries.get(id)!.decoded;
  if (decoded.status !== 'supported') throw Error('Missing image metadata');
  const changed = {
    ...decoded.value,
    attributes: { ...decoded.value.attributes, caption: 'Next caption' },
  };
  expect(withPayloadRecord(full, changed, 'First image').ok).toBe(false);
  const compact = withPreparedPayloadFields(full);
  expect(readGallery(compact, 'gallery')).toEqual(readGallery(full, 'gallery'));
  expect(customXmlNodes(compact, part.name)).toEqual(customXmlNodes(full, part.name));
  const children = node(compact, id).children;
  expect(children).toContain(opaque);
  expect(
    children.filter((child) => child.kind !== 'textValue' && child.localName === 'attr')
  ).toHaveLength(2);
  expect(withPayloadRecord(compact, changed, 'First image').ok).toBe(true);
  expect(withPreparedPayloadFields(compact)).toBe(compact);
});

test('preparation keeps unbound internal gallery metadata unchanged', () => {
  let pkg = openGallery(galleryDocument());
  const attributes = { url: 'https://example.test/orphan.png', caption: 'Orphan' };
  for (const caption of ['Orphan', 'Second', 'Third']) {
    const saved = withPayloadRecord(
      pkg,
      {
        id: 'unbound-image',
        version: 2,
        type: 'image_gallery_item',
        galleryId: 'gallery',
        attributes: { ...attributes, caption },
      },
      caption
    );
    if (!saved.ok) throw Error(saved.reason);
    pkg = saved.pkg;
  }
  expect(withPreparedPayloadFields(pkg)).toBe(pkg);
});
