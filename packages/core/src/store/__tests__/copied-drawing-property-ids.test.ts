import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  freshenCopiedDrawingPropertyIds,
  readOoxmlPart,
  type OoxmlNode,
} from '../index.ts';
import { runWithTransactionActor } from '../package/actor-scoped-ids.ts';
import { singlePartPackage } from '../package/ooxml-package.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const picture = (id: string) =>
  `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>` +
  `<wp:docPr id="${id}" name="Picture"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic>` +
  '<pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>' +
  '</wp:inline></w:drawing></w:r></w:p>';
function part(ids: string[], header = false) {
  const tag = header ? 'hdr' : 'document';
  const content = ids.map(picture).join('');
  const read = readOoxmlPart(
    `<w:${tag} xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}">${header ? content : `<w:body>${content}</w:body>`}</w:${tag}>`,
    {
      name: header ? '/word/header1.xml' : '/word/document.xml',
      contentType: header
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!read.ok) throw Error(read.reason);
  return read.part;
}
function propertyIds(nodes: readonly OoxmlNode[]): string[] {
  const result: string[] = [],
    pending = [...nodes];
  while (pending.length) {
    const node = pending.shift()!;
    if (node.kind === 'textValue') continue;
    if (node.kind === 'drawingDocPr')
      result.push(node.attributes.find((attribute) => attribute.localName === 'id')!.value);
    for (const child of node.children) pending.push(child);
  }
  return result;
}

test('copied drawing properties use the native package occupancy without changing their source', () => {
  const source = part(['3', '7']),
    header = part(['99'], true);
  const base = singlePartPackage(source),
    pkg = { ...base, parts: new Map([...base.parts, [header.name, header]]) };
  const blocks = bodyStoryRoot(source)!.children;
  const before = canonicalOoxmlFingerprint(source);
  const copied = freshenCopiedDrawingPropertyIds(pkg, blocks, 500_000, 64);
  expect(propertyIds(copied)).toEqual(['100', '101']);
  expect(propertyIds(blocks)).toEqual(['3', '7']);
  expect(canonicalOoxmlFingerprint(source)).toBe(before);
  expect(copied.map((node) => node.id)).toEqual(blocks.map((node) => node.id));
});

test('copied drawing property allocation follows the transaction actor for every copied image', () => {
  const source = part(['3', '7']),
    pkg = singlePartPackage(source);
  const blocks = bodyStoryRoot(source)!.children;
  const first = runWithTransactionActor('copy-actor-a', () =>
    freshenCopiedDrawingPropertyIds(pkg, blocks, 500_000, 64)
  );
  const second = runWithTransactionActor('copy-actor-b', () =>
    freshenCopiedDrawingPropertyIds(pkg, blocks, 500_000, 64)
  );
  expect(
    new Set([...propertyIds(blocks), ...propertyIds(first), ...propertyIds(second)]).size
  ).toBe(6);
  expect(propertyIds(first)).not.toEqual(['8', '9']);
});

test('copied drawing properties refuse ambiguous identifiers and bounded work before any mutation', () => {
  const source = part(['7', '07']),
    pkg = singlePartPackage(source);
  const blocks = bodyStoryRoot(source)!.children;
  const before = canonicalOoxmlFingerprint(source);
  expect(() => freshenCopiedDrawingPropertyIds(pkg, blocks, 500_000, 64)).toThrow('ambiguous');
  for (const budget of [0, -1, NaN, Infinity])
    expect(() => freshenCopiedDrawingPropertyIds(pkg, blocks, budget, 64)).toThrow(
      'budgets are invalid'
    );
  expect(() => freshenCopiedDrawingPropertyIds(pkg, blocks, 1, 64)).toThrow('fragment limits');
  expect(() => freshenCopiedDrawingPropertyIds(pkg, blocks, 500_000, 1)).toThrow('fragment limits');
  expect(canonicalOoxmlFingerprint(source)).toBe(before);
});
