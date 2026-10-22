import { expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { readOoxmlPackage, stylesPartOf } from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { prepareBlockControlsDocument, withBlockTypeStyles } from './block-styles';

test('block preparation preserves imported story and opaque bytes and is byte-idempotent', () => {
  const entries = unzipSync(docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>'));
  entries['word/media/unknown.bin'] = new Uint8Array([0, 255, 7, 0]);
  entries['customXml/foreign.xml'] = strToU8('<foreign value=" exact "/>');
  const input = zipSync(entries),
    original = input.slice();
  const result = prepareBlockControlsDocument(input);
  if (!result.success) throw new Error(result.error);
  const output = unzipSync(result.document);
  expect(input).toEqual(original);
  for (const path of ['word/document.xml', 'word/media/unknown.bin', 'customXml/foreign.xml'])
    expect(output[path]).toEqual(entries[path]);
  const loaded = readOoxmlPackage(result.document);
  if (!loaded.ok) throw new Error(loaded.reason);
  const styles = stylesPartOf(loaded.package)!;
  expect(
    styles.root.children.filter(
      (node) =>
        node.kind !== 'textValue' &&
        node.attributes.some(
          (attribute) =>
            attribute.localName === 'styleId' &&
            [
              'Normal',
              'Heading1',
              'Heading2',
              'Heading3',
              'Quote',
              'SciobotHorizontalRule',
            ].includes(attribute.value)
        )
    )
  ).toHaveLength(6);
  const again = prepareBlockControlsDocument(result.document);
  expect(again.success).toBe(true);
  if (!again.success) throw new Error(again.error);
  expect(again.changed).toBe(false);
  expect(again.document).toBe(result.document);
});

for (const definitions of [
  '<w:style w:type="character" w:styleId="Heading1"/>',
  '<w:style w:type="paragraph" w:styleId="Quote"/><w:style w:type="paragraph" w:styleId="Quote"/>',
  '<w:style w:type="paragraph" w:styleId="SciobotHorizontalRule"><w:name w:val="Foreign"/></w:style>',
])
  test(`conflicting block identity refuses immutable preparation: ${definitions}`, () => {
    const input = docxFixture('<w:p/>', {
      styles: `<w:styles xmlns:w="${W}">${definitions}</w:styles>`,
    });
    const original = input.slice(),
      loaded = readOoxmlPackage(input);
    if (!loaded.ok) throw new Error(loaded.reason);
    const pkg = loaded.package;
    expect(withBlockTypeStyles(pkg)).toBeNull();
    expect(prepareBlockControlsDocument(input)).toMatchObject({
      success: false,
      code: 'invalid-document',
    });
    expect(input).toEqual(original);
  });
