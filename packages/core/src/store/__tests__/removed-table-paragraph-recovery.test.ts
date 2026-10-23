import { expect, test } from 'bun:test';
import { runWithTransactionActor } from '../package/actor-scoped-ids.ts';
import { isRetiredParagraph } from '../package/retired-paragraph.ts';
import { zipDoc } from './canonical-primitive-journal-coverage-support.ts';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  validateOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '../index.ts';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const mark = 'w:id="17" w:author="AI" w:date="2026-10-23T00:00:00Z"';
const peer = '<w:r><w:t xml:space="preserve"> Peer</w:t></w:r>';
const owned = `<w:ins ${mark}><w:r><w:t>Owned</w:t></w:r></w:ins>`;
const table = (text: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:ins ${mark}/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p>${text}</w:p></w:tc></w:tr></w:tbl>`;
function load(body: string) {
  const read = readOoxmlPackage(zipDoc({ body }));
  if (!read.ok) throw Error(read.reason);
  return read.package;
}
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
function word(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(main(read.package));
}
for (const actor of [undefined, 'alice']) {
  test(`removed table recovers foreign text without its owned text; actor=${actor}`, () => {
    const pkg = load(p('Before') + table(owned + peer) + p('After'));
    const before = main(pkg);
    const body = before.root.children.find((node) => node.kind === 'body')!;
    const cell = body.children.find((node) => node.kind === 'table')!;
    const paragraph = cell.children
      .find((node) => node.kind === 'tableRow')!
      .children.find((node) => node.kind === 'tableCell')!
      .children.find((node) => node.kind === 'paragraph')!;
    const result = runWithTransactionActor(actor, () =>
      applyTreeOp(before, { op: 'rejectAllRevisions' })
    );
    if (!result.ok) throw Error(result.reason);
    expect(validateOoxmlPart(result.part)).toEqual({ ok: true });
    expect(collectRevisionSites(result.part)).toHaveLength(0);
    expect(isRetiredParagraph(findNode(result.part, paragraph.id)!)).toBe(true);
    expect(paragraphTextOf(result.part, paragraph.id)).toBe(' Peer');
    expect(word(withPart(pkg, result.part))).toBe(
      word(load(p('Before') + '<w:p><w:pPr/>' + peer + '</w:p>' + p('After')))
    );
  });
  test(`removed table without foreign text exports exact original; actor=${actor}`, () => {
    const pkg = load(p('Before') + table(owned) + p('After'));
    const result = runWithTransactionActor(actor, () =>
      applyTreeOp(main(pkg), { op: 'rejectAllRevisions' })
    );
    if (!result.ok) throw Error(result.reason);
    expect(word(withPart(pkg, result.part))).toBe(word(load(p('Before') + p('After'))));
  });
}

for (const kept of ['attribute', 'child', 'source-presence'] as const)
  test(`table recovery preserves ${kept} in the paragraph mark holder`, () => {
    const attrs = kept === 'attribute' ? ' xmlns:x="urn:keep" x:stamp="keep"' : '';
    const child = kept === 'child' ? '<x:data xmlns:x="urn:keep"/>' : '';
    const presence =
      kept === 'source-presence'
        ? ' xmlns:source="urn:docx-editor:inserted-paragraph-property-presence:1" source:properties="pPr rPr"'
        : '';
    const properties = `<w:pPr><w:rPr${attrs}><w:ins ${mark}${presence}/>${child}</w:rPr></w:pPr>`;
    const pkg = load(p('Before') + table(properties + owned + peer) + p('After'));
    const result = runWithTransactionActor('alice', () =>
      applyTreeOp(main(pkg), { op: 'rejectAllRevisions' })
    );
    if (!result.ok) throw Error(result.reason);
    const expected = `<w:p><w:pPr><w:rPr${attrs}>${child}</w:rPr></w:pPr>${peer}</w:p>`;
    expect(word(withPart(pkg, result.part))).toBe(word(load(p('Before') + expected + p('After'))));
  });
