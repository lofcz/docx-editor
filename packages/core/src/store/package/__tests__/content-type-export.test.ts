import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readOoxmlPackage, writeOoxmlPackage, type OoxmlPackage } from '../ooxml-package.ts';
import { contentTypeExportPackage } from '../content-type-export.ts';

import { customXmlDataParts, withCustomXmlDataPart } from '../custom-xml-part.ts';
import { TreePackageStore } from '../../store/tree-package-store.ts';
import {
  isCanonicalPrimitiveCaptureActive,
  journalCaptureMark,
  observeCanonicalPrimitiveJournal,
} from '../canonical-primitive-capture.ts';

const TYPES = '/[Content_Types].xml';
const CUSTOM_NS = 'urn:application:test';
const FORK_TYPES =
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:q="urn:opaque" q:value="q:Keep"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><q:Extension keep="yes"/></Types>';
const NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const AUTHORED = `<Types xmlns="${NS}" xmlns:x="urn:extra">
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="${MAIN}"/>
<Override PartName="/Orphan.XML" ContentType="application/xml" x:keep="yes"><x:child/></Override>
<Override PartName="/Orphan.XML" ContentType="application/xml"/>
<Override PartName="/invalid.xml" ContentType="invalid"/>
<Override PartName="/ignored.xml" ContentType="image/.jpg"/>
<Override PartName="/ignored.xml" ContentType="image/jpeg"/>
<x:Override PartName="/foreign.xml" ContentType="foreign"/>
<Override PartName="/missing.xml" x:ContentType="malformed"/>
</Types>`;

function open(authored = AUTHORED, parts: Record<string, Uint8Array> = {}): OoxmlPackage {
  const loaded = readOoxmlPackage(
    zipSync({
      ...parts,
      '[Content_Types].xml': strToU8(authored),
      '_rels/.rels': strToU8(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
      ),
      'word/document.xml': strToU8(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'
      ),
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.package;
}

function types(pkg: OoxmlPackage): string {
  return strFromU8(unzipSync(writeOoxmlPackage(pkg))['[Content_Types].xml']!);
}

test('unchanged index preserves authored bytes, duplicate records, and ignored records', () => {
  const pkg = open();
  expect(contentTypeExportPackage(pkg)).toBe(pkg);
  expect(types(pkg)).toBe(AUTHORED);
});

test('new overrides preserve unrelated authored records and leave the input unchanged', () => {
  const pkg = open();
  const overrides = new Map(pkg.contentTypes.overrides);
  overrides.set('/word/new.xml', 'application/example+xml');
  const next = { ...pkg, contentTypes: { ...pkg.contentTypes, overrides } };
  const xml = types(next);
  expect(xml).toContain('PartName="/word/new.xml"');
  expect(xml.match(/PartName="\/Orphan.XML"/g)).toHaveLength(2);
  for (const name of ['invalid', 'ignored', 'foreign', 'missing']) {
    expect(xml).toContain(`PartName="/${name}.xml"`);
  }
  expect(xml).toContain('keep="yes"');
  expect(xml).toContain('child');
  expect(types(pkg)).toBe(AUTHORED);
  expect(next.partBytes).toBe(pkg.partBytes);
});

test('changed and deleted overrides update the saved bytes without dropping ignored records', () => {
  const pkg = open();
  const overrides = new Map(pkg.contentTypes.overrides);
  overrides.set('/orphan.xml', 'application/example+xml');
  overrides.delete('/ignored.xml');
  const next = { ...pkg, contentTypes: { ...pkg.contentTypes, overrides } };
  const xml = types(next);
  expect(xml).toContain('application/example+xml');
  expect(xml).toContain('<Override ContentType="image/.jpg" PartName="/ignored.xml"/>');
  expect(xml).not.toContain('<Override ContentType="image/jpeg" PartName="/ignored.xml"/>');
  expect(xml).toContain('PartName="/invalid.xml"');
  expect(xml).toContain('keep="yes"');
  const reopened = readOoxmlPackage(writeOoxmlPackage(next));
  expect(reopened.ok).toBe(true);
  if (reopened.ok) {
    const expected = new Map(overrides);
    // The preserved unconventional record resolves once its competing record is removed.
    expected.set('/ignored.xml', 'image/.jpg');
    expect(reopened.package.contentTypes).toEqual({ ...next.contentTypes, overrides: expected });
  }
});

test('changed overrides match percent-encoded authored names in place', () => {
  const authored = AUTHORED.replace(
    '</Types>',
    '<Override PartName="/word/media/my%20photo.png" ContentType="image/png"/></Types>'
  );
  const pkg = open(authored, { 'word/media/my photo.png': new Uint8Array([1, 2, 3]) });
  const overrides = new Map(pkg.contentTypes.overrides);
  overrides.set('/word/media/my photo.png', 'image/jpeg');
  const next = { ...pkg, contentTypes: { ...pkg.contentTypes, overrides } };
  const xml = types(next);
  expect(xml.match(/PartName="\/word\/media\/my(?:%20| )photo.png"/g)).toHaveLength(1);
  expect(xml).toContain(
    '<Override ContentType="image/jpeg" PartName="/word/media/my%20photo.png"/>'
  );
  expect(readOoxmlPackage(writeOoxmlPackage(next)).ok).toBe(true);
});

test('deleted overrides preserve ignored declarations with the same normalized name', () => {
  const ignored = [
    '<Override ContentType="invalid" PartName="/word/media/my photo.png" x:keep="yes"><x:child/></Override>',
    '<Override PartName="/word/media/my photo.png" x:keep="missing"><x:child/></Override>',
    '<Override ContentType="image/.jpg" PartName="/word/media/my photo.png" x:keep="unconventional"/>',
  ];
  const valid = '<Override ContentType="image/png" PartName="/word/media/my photo.png"/>';
  const pkg = open(AUTHORED.replace('</Types>', ignored.join('') + valid + '</Types>'));
  const overrides = new Map(pkg.contentTypes.overrides);
  overrides.delete('/word/media/my photo.png');
  const xml = types({ ...pkg, contentTypes: { ...pkg.contentTypes, overrides } });
  expect(xml).not.toContain(valid);
  for (const declaration of ignored) expect(xml).toContain(declaration);
});

test('new overrides save in the same order whatever the index order', () => {
  const pkg = open();
  const save = (entries: [string, string][]): string => {
    const overrides = new Map(pkg.contentTypes.overrides);
    for (const [name, type] of entries) overrides.set(name, type);
    return types({ ...pkg, contentTypes: { ...pkg.contentTypes, overrides } });
  };
  const header: [string, string] = ['/word/header1.xml', 'application/example-header+xml'];
  const footer: [string, string] = ['/word/footer1.xml', 'application/example-footer+xml'];
  expect(save([header, footer])).toBe(save([footer, header]));
});

test('an index type with invalid syntax is not saved', () => {
  const pkg = open();
  const overrides = new Map(pkg.contentTypes.overrides);
  overrides.set('/word/document.xml', 'not a type');
  overrides.set('/word/new.xml', 'also not a type');
  const next = { ...pkg, contentTypes: { ...pkg.contentTypes, overrides } };
  const xml = types(next);
  expect(xml).toContain(`ContentType="${MAIN}"`);
  expect(xml).not.toContain('not a type');
  const reopened = readOoxmlPackage(writeOoxmlPackage(next));
  expect(reopened.ok).toBe(true);
});

function retained() {
  const before = open(FORK_TYPES);
  const added = withCustomXmlDataPart(before, before.mainDocumentPart, CUSTOM_NS, 'application');
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
  ).toEqual([CUSTOM_NS]);
  expect(stale.partBytes.get(TYPES)).toBe(bytes);
  expect(stale.contentTypes).toBe(declared);
  const xml = strFromU8(unzipSync(exported)['[Content_Types].xml']!);
  expect(xml).toContain('q:Keep');
  expect(xml).toContain('urn:opaque');
  expect(xml).toContain('keep="yes"');
  expect(contentTypeExportPackage(stale)).toBe(contentTypeExportPackage(stale));
});
test('consistent packages retain exact content-type bytes, defaults, and unrelated extensions', () => {
  const { complete } = retained();
  expect(contentTypeExportPackage(complete)).toBe(complete);
  expect(unzipSync(writeOoxmlPackage(complete))['[Content_Types].xml']).toEqual(
    new Uint8Array(complete.partBytes.get(TYPES)!)
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
