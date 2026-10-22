import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  relationshipsOf,
  stylesPartOf,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { inlineCodeStyleId, inlineCodeStyleState } from './inline-code';
import { prepareNativeDocument } from './prepare-document';
import { readReviewGroupStore } from '../review/group-record-store';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BODY =
  '<w:p><w:pPr><w:keepNext/><w:customKeep w:val="keep"/></w:pPr>' +
  '<w:ins w:id="7" w:author="Teacher"><w:r><w:t>Pending</w:t></w:r></w:ins>' +
  '<w:r><w:t xml:space="preserve"> &amp; preserved 😀</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
const ready = (bytes: Uint8Array) => {
  const result = prepareNativeDocument(bytes);
  if (!result.success) throw new Error(result.error);
  return result;
};
const pkg = (bytes: Uint8Array) => {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error(read.reason);
  return read.package;
};

test('preparation adds shared code infrastructure and preserves every unrelated input part byte-for-byte', () => {
  const entries = unzipSync(docxFixture(BODY));
  entries['word/media/imported.bin'] = new Uint8Array([0, 255, 8, 19]);
  entries['customXml/future.xml'] = strToU8('<future  value="opaque">  exact  </future>');
  entries['word/header1.xml'] = strToU8(
    `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`
  );
  const input = zipSync(entries);
  const original = input.slice();
  expect(stylesPartOf(pkg(input))).toBeUndefined();
  const prepared = ready(input);
  expect(prepared).toMatchObject({ changed: true, codeStyleReady: true });
  const after = unzipSync(prepared.document);
  for (const [name, bytes] of Object.entries(entries))
    if (name !== '[Content_Types].xml' && name !== 'word/_rels/document.xml.rels')
      expect(after[name]).toEqual(bytes);
  expect(input).toEqual(original);
  expect(inlineCodeStyleState(pkg(prepared.document))).toBe('present');
  expect(readReviewGroupStore(pkg(prepared.document)).status).toBe('ready');
  expect(strFromU8(after['word/styles.xml']!).match(/w:styleId="SciobotCode"/g)).toHaveLength(1);
  expect(strFromU8(after['word/_rels/document.xml.rels']!)).toContain('Target="styles.xml"');
  expect(ready(prepared.document)).toMatchObject({ changed: false, codeStyleReady: true });
  expect(ready(prepared.document).document).toBe(prepared.document);
});

test('preparation retains imported style definitions and their nonstandard part location', () => {
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="Imported"><w:name w:val="Imported"/><w:rPr><w:b/><w:customKeep w:val="style"/></w:rPr></w:style></w:styles>`;
  const entries = unzipSync(docxFixture(BODY, { styles }));
  entries['word/custom-styles.xml'] = entries['word/styles.xml']!;
  delete entries['word/styles.xml'];
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '/word/styles.xml',
      '/word/custom-styles.xml'
    )
  );
  entries['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(entries['word/_rels/document.xml.rels']!).replace(
      'Target="styles.xml"',
      'Target="custom-styles.xml"'
    )
  );
  const input = zipSync(entries);
  const imported = stylesPartOf(pkg(input))!.root.children[0]!;
  const prepared = ready(input);
  const output = unzipSync(prepared.document);
  expect(output['word/styles.xml']).toBeUndefined();
  expect(output['word/document.xml']).toEqual(entries['word/document.xml']);
  const beforeRelationships = relationshipsOf(pkg(input), '/word/document.xml'),
    afterRelationships = relationshipsOf(pkg(prepared.document), '/word/document.xml');
  for (const relationship of beforeRelationships)
    expect(afterRelationships).toContainEqual(relationship);
  expect(afterRelationships.length).toBe(beforeRelationships.length + 1);
  expect(readReviewGroupStore(pkg(prepared.document)).status).toBe('ready');
  expect(canonicalOoxmlFingerprint(stylesPartOf(pkg(prepared.document))!.root.children[0]!)).toBe(
    canonicalOoxmlFingerprint(imported)
  );
  expect(stylesPartOf(pkg(prepared.document))?.name).toBe('/word/custom-styles.xml');
  expect(inlineCodeStyleState(pkg(prepared.document), inlineCodeStyleId('Imported'))).toBe(
    'present'
  );
  expect(ready(prepared.document).document).toBe(prepared.document);
});

test('an existing code style still seeds imported variants without overwriting a conflicting variant', () => {
  const input = docxFixture(BODY, {
    styles:
      `<w:styles xmlns:w="${W}">` +
      '<w:style w:type="character" w:styleId="SciobotCode"><w:name w:val="Sciobot Code"/></w:style>' +
      '<w:style w:type="character" w:styleId="Imported"><w:name w:val="Imported"/></w:style>' +
      '<w:style w:type="character" w:styleId="Other"><w:name w:val="Other"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="SciobotCode__Other"><w:name w:val="Foreign"/></w:style>' +
      '</w:styles>',
  });
  const before = stylesPartOf(pkg(input))!.root.children;
  const prepared = ready(input);
  expect(prepared.changed).toBe(true);
  const after = pkg(prepared.document);
  expect(inlineCodeStyleState(after, inlineCodeStyleId('Imported'))).toBe('present');
  expect(inlineCodeStyleState(after, inlineCodeStyleId('Other'))).toBe('conflict');
  expect(
    stylesPartOf(after)!
      .root.children.slice(0, before.length)
      .map((node) => canonicalOoxmlFingerprint(node))
  ).toEqual(before.map((node) => canonicalOoxmlFingerprint(node)));
  expect(ready(prepared.document).document).toBe(prepared.document);
});

test('a conflicting imported identity stays intact while review storage initializes', () => {
  const input = docxFixture(BODY, {
    styles: `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="SciobotCode"><w:name w:val="Foreign"/></w:style></w:styles>`,
  });
  const result = ready(input);
  expect(result).toMatchObject({ changed: true, codeStyleReady: false });
  const original = unzipSync(input),
    after = unzipSync(result.document);
  expect(after['word/styles.xml']).toEqual(original['word/styles.xml']);
  expect(after['word/document.xml']).toEqual(original['word/document.xml']);
  expect(readReviewGroupStore(pkg(result.document)).status).toBe('ready');
  expect(ready(result.document)).toMatchObject({ changed: false, codeStyleReady: false });
  expect(ready(result.document).document).toBe(result.document);
});

test('an unrelated part occupying the conventional styles filename stays intact', () => {
  const entries = unzipSync(docxFixture(BODY));
  const foreign = strToU8('<foreign purpose="preserve"/>');
  entries['word/STYLES.xml'] = foreign;
  entries['word/sciobot-styles-1.xml'] = foreign;
  const prepared = ready(zipSync(entries));
  const output = unzipSync(prepared.document);
  expect(output['word/STYLES.xml']).toEqual(foreign);
  expect(output['word/sciobot-styles-1.xml']).toEqual(foreign);
  expect(output['word/document.xml']).toEqual(entries['word/document.xml']);
  expect(stylesPartOf(pkg(prepared.document))?.name).toBe('/word/sciobot-styles-2.xml');
  expect(inlineCodeStyleState(pkg(prepared.document))).toBe('present');
});

test('invalid input fails before creating any document', () => {
  expect(prepareNativeDocument(new Uint8Array([1, 2, 3]))).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  expect(prepareNativeDocument(null as unknown as Uint8Array)).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
});

test('a valid document without owner relationships gains review and style infrastructure before opening', () => {
  const entries = unzipSync(docxFixture(BODY));
  delete entries['word/_rels/document.xml.rels'];
  entries['word/media/untouched.bin'] = new Uint8Array([0, 9, 255]);
  const input = zipSync(entries),
    original = input.slice(),
    prepared = ready(input),
    loaded = pkg(prepared.document),
    bytes = unzipSync(prepared.document);
  expect(prepared).toMatchObject({ changed: true, codeStyleReady: true });
  expect(inlineCodeStyleState(loaded)).toBe('present');
  expect(readReviewGroupStore(loaded).status).toBe('ready');
  expect(
    relationshipsOf(loaded, loaded.mainDocumentPart).some((record) =>
      record.type.endsWith('/styles')
    )
  ).toBe(true);
  expect(
    relationshipsOf(loaded, loaded.mainDocumentPart).some((record) =>
      record.type.endsWith('/customXml')
    )
  ).toBe(true);
  for (const [name, value] of Object.entries(entries))
    if (name !== '[Content_Types].xml') expect(bytes[name]).toEqual(value);
  expect(bytes['word/_rels/document.xml.rels']).toBeDefined();
  expect(input).toEqual(original);
  const again = ready(prepared.document);
  expect(again.changed).toBe(false);
  expect(again.document).toBe(prepared.document);
});

test('opaque existing owner relationship bytes refuse preparation without replacing them', () => {
  const entries = unzipSync(docxFixture(BODY));
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/_rels/document.xml.rels" ContentType="application/octet-stream"/></Types>'
    )
  );
  const input = zipSync(entries),
    original = input.slice();
  expect(prepareNativeDocument(input).success).toBe(false);
  expect(input).toEqual(original);
  expect(unzipSync(input)['word/_rels/document.xml.rels']).toEqual(
    entries['word/_rels/document.xml.rels']
  );
});
