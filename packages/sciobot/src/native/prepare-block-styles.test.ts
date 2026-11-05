import { expect, test } from 'bun:test';
import { blankDocumentBytes, createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  stylesPartOf,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from './prepare-document';

const W = WML_NAMESPACE_URI;
const load = (bytes: Uint8Array) => {
  const result = readOoxmlPackage(bytes);
  if (!result.ok) throw Error(result.reason);
  return result.package;
};
const prepare = (bytes: Uint8Array) => {
  const result = prepareNativeDocument(bytes);
  if (!result.success) throw Error(result.error);
  return result;
};

test('the common preparation path adds missing named styles without changing the imported defaults', () => {
  const original = docxFixture('<w:p><w:r><w:t>Unstyled import</w:t></w:r></w:p>');
  const prepared = prepare(original);
  const before = createDocxEditor({ container: document.createElement('div'), document: original });
  const after = createDocxEditor({
    container: document.createElement('div'),
    document: prepared.document,
  });
  try {
    expect(
      after
        .getDocumentStyles()
        .filter((style) =>
          ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Quote'].includes(style.styleId)
        )
    ).toHaveLength(5);
    expect(after.getSelectionFormatting()).toEqual(before.getSelectionFormatting());
    expect(canonicalOoxmlFingerprint(after.surface!.session.part())).toBe(
      canonicalOoxmlFingerprint(before.surface!.session.part())
    );
    const styles = stylesPartOf(load(prepared.document))!;
    expect(
      styles.root.children.some(
        (node) =>
          node.kind !== 'textValue' &&
          node.attributes.some(
            (attribute) => attribute.namespaceUri === W && attribute.localName === 'default'
          )
      )
    ).toBe(false);
    expect(prepare(prepared.document)).toMatchObject({
      changed: false,
      document: prepared.document,
    });
  } finally {
    before.destroy();
    after.destroy();
  }
});

for (const definition of [
  '<w:style w:type="character" w:styleId="Heading1"><w:name w:val="Foreign"/><w:rPr><w:color w:val="765432"/><w:opaque w:val="keep"/></w:rPr></w:style>',
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:rFonts w:ascii="Imported Face"/><w:sz w:val="67"/></w:rPr></w:style>',
  '<w:style w:type="paragraph" w:default="1" w:styleId="Custom"><w:pPr><w:spacing w:after="333"/></w:pPr></w:style>',
  '<w:style w:type="paragraph" w:styleId="SciobotHorizontalRule"><w:name w:val="Foreign"/></w:style>',
])
  test(`imported definitions remain exact while other headings become available: ${definition}`, () => {
    const source = docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>', {
      styles: `<w:styles xmlns:w="${W}">${definition}</w:styles>`,
    });
    const imported = stylesPartOf(load(source))!.root.children.map(canonicalOoxmlFingerprint);
    const prepared = prepare(source);
    const styles = stylesPartOf(load(prepared.document))!.root.children;
    expect(styles.slice(0, imported.length).map(canonicalOoxmlFingerprint)).toEqual(imported);
    for (const id of ['Heading2', 'Heading3'])
      expect(
        styles.filter(
          (node) =>
            node.kind !== 'textValue' &&
            node.attributes.some((entry) => entry.localName === 'styleId' && entry.value === id)
        )
      ).toHaveLength(1);
    expect(prepare(prepared.document).changed).toBe(false);
  });

test('fresh document preparation preserves the complete existing style definitions', () => {
  const source = blankDocumentBytes();
  const before = stylesPartOf(load(source))!.root.children.map(canonicalOoxmlFingerprint);
  const after = stylesPartOf(load(prepare(source).document))!.root.children;
  expect(after.slice(0, before.length).map(canonicalOoxmlFingerprint)).toEqual(before);
});
