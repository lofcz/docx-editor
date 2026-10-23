import { expect, test } from 'bun:test';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import { readOoxmlPackage, writeOoxmlPackage } from '../ooxml-package.ts';
import { customXmlDataParts, withCustomXmlDataPart } from '../custom-xml-part.ts';
import { projectContentTypeBytes } from '../content-type-export.ts';
import { TreePackageStore } from '../../store/tree-package-store.ts';
import {
  isCanonicalPrimitiveCaptureActive,
  journalCaptureMark,
  observeCanonicalPrimitiveJournal,
} from '../canonical-primitive-capture.ts';
const TYPES = '/[Content_Types].xml';
const NS = 'urn:application:test';
function source() {
  const read = readOoxmlPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:q="urn:opaque" q:value="q:Keep"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><q:Extension keep="yes"/></Types>'
      ),
      '_rels/.rels': strToU8(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
      ),
      'word/document.xml': strToU8(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'
      ),
    })
  );
  if (!read.ok) throw Error(read.reason);
  return read.package;
}
function retained() {
  const before = source();
  const added = withCustomXmlDataPart(before, before.mainDocumentPart, NS, 'application');
  if (!added.part) throw Error('Missing new data part');
  return {
    before,
    complete: added.pkg,
    stale: {
      ...added.pkg,
      parts: new Map(
        [...added.pkg.parts].map(([name, part]) => [
          name,
          name === added.part!.propsPartName ? { ...part, contentType: 'application/xml' } : part,
        ])
      ),
      partBytes: new Map(added.pkg.partBytes).set(TYPES, before.partBytes.get(TYPES)!),
    },
  };
}
test('export restores retained declaration metadata when historical ZIP bytes omit new parts', () => {
  const { stale } = retained();
  const bytes = stale.partBytes.get(TYPES)!;
  const declared = stale.contentTypes;
  const exported = writeOoxmlPackage(stale);
  const reopened = readOoxmlPackage(exported);
  if (!reopened.ok) throw Error(reopened.reason);
  expect(
    customXmlDataParts(reopened.package, reopened.package.mainDocumentPart).map(
      (part) => part.namespaceUri
    )
  ).toEqual([NS]);
  expect(stale.partBytes.get(TYPES)).toBe(bytes);
  expect(stale.contentTypes).toBe(declared);
  const xml = strFromU8(unzipSync(exported)['[Content_Types].xml']!);
  expect(xml).toContain('q:Keep');
  expect(xml).toContain('urn:opaque');
  expect(xml).toContain('keep="yes"');
  expect(projectContentTypeBytes(stale)).toBe(projectContentTypeBytes(stale));
});
test('consistent packages retain exact content-type bytes, defaults, and unrelated extensions', () => {
  const { complete } = retained();
  expect(projectContentTypeBytes(complete)).toBe(complete);
  expect(unzipSync(writeOoxmlPackage(complete))['[Content_Types].xml']).toEqual(
    complete.partBytes.get(TYPES)!
  );
});
test('cold export in an observed native transaction records no canonical primitive and changes no history', () => {
  const { stale } = retained();
  const store = new TreePackageStore(stale, stale.parts.get(stale.mainDocumentPart)!);
  const initial = store.currentPackage();
  const stop = observeCanonicalPrimitiveJournal(store, () => {
    throw Error('Read-only export published a journal');
  });
  try {
    const result = store.transact({ kind: 'body' }, (ctx) => {
      if (
        !ctx.applyPackage((pkg) => {
          expect(isCanonicalPrimitiveCaptureActive()).toBe(true);
          const mark = journalCaptureMark();
          writeOoxmlPackage(structuredClone(pkg));
          expect(journalCaptureMark()).toBe(mark);
          return pkg;
        })
      )
        throw Error('Read-only completion refused');
    });
    expect(result.ok).toBe(true);
    expect(store.currentPackage()).toBe(initial);
    expect(store.canUndo).toBe(false);
  } finally {
    stop();
  }
});
