import { expect, test } from 'bun:test';
import { strToU8, strFromU8, unzipSync, zipSync } from 'fflate';
import { readOoxmlPackage, relationshipsOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from './prepare-document';
import {
  nativeNumberingInfrastructureReady,
  withNativeNumberingInfrastructure,
} from './numbering-infrastructure';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const read = (bytes: Uint8Array) => {
  const result = readOoxmlPackage(bytes);
  if (!result.ok) throw Error(result.reason);
  return result.package;
};
const prepared = (bytes: Uint8Array) => {
  const result = prepareNativeDocument(bytes);
  if (!result.success) throw Error(result.error);
  return result;
};
test('preparation creates an empty numbering root with a permanent namespace, relationship, and declaration', () => {
  const bytes = docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>');
  const original = unzipSync(bytes),
    result = prepared(bytes),
    pkg = read(result.document),
    part = pkg.parts.get('/word/numbering.xml')!;
  expect(part.root.namespaceUri).toBe(W);
  expect(part.root.localName).toBe('numbering');
  expect(part.root.children).toEqual([]);
  expect(nativeNumberingInfrastructureReady(pkg)).toBe(true);
  expect(strFromU8(unzipSync(result.document)['word/numbering.xml']!)).toContain(`xmlns:w="${W}"`);
  expect(unzipSync(result.document)['word/document.xml']).toEqual(original['word/document.xml']!);
  const again = prepared(result.document);
  expect(again.changed).toBe(false);
  expect(again.document).toBe(result.document);
});
test('existing numbering and its opaque data remain byte-for-byte unchanged', () => {
  const numbering = `<w:numbering xmlns:w="${W}" xmlns:x="urn:keep" x:value="x:Keep"><x:opaque>Keep</x:opaque></w:numbering>`;
  const bytes = docxFixture('<w:p/>', { numbering }),
    pkg = read(bytes);
  expect(withNativeNumberingInfrastructure(pkg)).toBe(pkg);
  expect(unzipSync(prepared(bytes).document)['word/numbering.xml']).toEqual(strToU8(numbering));
});
test('opaque numbering filename collisions and nonstandard imported numbering stay intact', () => {
  for (const other of [false, true]) {
    const original = unzipSync(
      docxFixture('<w:p/>', { numbering: `<w:numbering xmlns:w="${W}"/>` })
    );
    if (other) {
      original['word/alternate.xml'] = original['word/numbering.xml']!;
      delete original['word/numbering.xml'];
      original['[Content_Types].xml'] = strToU8(
        strFromU8(original['[Content_Types].xml']!).replace(
          '/word/numbering.xml',
          '/word/alternate.xml'
        )
      );
      original['word/_rels/document.xml.rels'] = strToU8(
        strFromU8(original['word/_rels/document.xml.rels']!).replace(
          'Target="numbering.xml"',
          'Target="alternate.xml"'
        )
      );
    } else {
      original['word/numbering.xml'] = strToU8('<foreign xmlns="urn:keep">opaque</foreign>');
    }
    const bytes = zipSync(original),
      pkg = read(bytes);
    expect(withNativeNumberingInfrastructure(pkg)).toBe(pkg);
    expect(nativeNumberingInfrastructureReady(pkg)).toBe(false);
    const output = unzipSync(prepared(bytes).document),
      name = other ? 'word/alternate.xml' : 'word/numbering.xml';
    expect(output[name]).toEqual(original[name]!);
    if (other) expect(output['word/numbering.xml']).toBeUndefined();
  }
});
test('a nonstandard main document gets a correctly relative numbering relationship', () => {
  const entries = unzipSync(docxFixture('<w:p/>'));
  entries['main.xml'] = entries['word/document.xml']!;
  delete entries['word/document.xml'];
  entries['_rels/main.xml.rels'] = entries['word/_rels/document.xml.rels']!;
  delete entries['word/_rels/document.xml.rels'];
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace('/word/document.xml', '/main.xml')
  );
  entries['_rels/.rels'] = strToU8(
    strFromU8(entries['_rels/.rels']!).replace('word/document.xml', 'main.xml')
  );
  const pkg = read(zipSync(entries)),
    next = withNativeNumberingInfrastructure(pkg);
  if (!next) throw Error('Numbering initialization refused');
  expect(nativeNumberingInfrastructureReady(next)).toBe(true);
  expect(
    relationshipsOf(next, '/main.xml').find((rel) => rel.type.endsWith('/numbering'))?.rawTarget
  ).toBe('word/numbering.xml');
  expect(nativeNumberingInfrastructureReady(read(writeOoxmlPackage(next)))).toBe(true);
});

test('case-folded opaque filename collisions do not create a second numbering part', () => {
  for (const name of ['word/numbering.xml', 'WORD/NUMBERING.XML']) {
    const entries = unzipSync(docxFixture('<w:p/>')),
      opaque = strToU8('<opaque xmlns="urn:keep">exact</opaque>');
    entries[name] = opaque;
    const bytes = zipSync(entries),
      pkg = read(bytes);
    expect(withNativeNumberingInfrastructure(pkg)).toBe(pkg);
    const output = unzipSync(prepared(bytes).document);
    expect(output[name]).toEqual(opaque);
    if (name !== 'word/numbering.xml') expect(output['word/numbering.xml']).toBeUndefined();
  }
});
