import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import type { OoxmlElement, OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { readOoxmlPart, serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { ADA, apply, part } from './tracked-edit-fixture.ts';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const body = (source: OoxmlPart) => source.root.children[0] as OoxmlElement;
const paragraphs = (source: OoxmlPart): OoxmlParagraphNode[] =>
  body(source).children.filter((node) => node.kind === 'paragraph') as OoxmlParagraphNode[];
const seed = (xml: string) => paragraphs(part(xml)) as OoxmlParagraphNode[];
const text = (source: OoxmlPart) =>
  paragraphs(source).map((node) => paragraphTextOf(source, node.id));
const attribute = (node: OoxmlElement, name: string) =>
  node.attributes.find((attr) => attr.localName === name)?.value;
const reopen = (source: OoxmlPart) => {
  const read = readOoxmlPart(serializeOoxmlPart(source), {
    name: source.name,
    contentType: source.contentType,
  });
  if (!read.ok) throw new Error(read.reason);
  return read.part;
};

for (const at of [0, 1, 2]) {
  for (const tracked of [false, true]) {
    test(`whole sibling paragraphs preserve anchors at ${at}, tracked=${tracked}`, () => {
      const original = part(p('Opening.') + p('Closing.') + '<w:sectPr/>');
      const originalParagraphs = paragraphs(original);
      const inserted = seed(
        '<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>New section</w:t></w:r></w:p>' +
          p('First addition.') +
          p('Second addition.')
      );
      const op: TreeDocOp = {
        op: 'insertParagraphs',
        parentId: body(original).id,
        beforeId: originalParagraphs[at]?.id ?? null,
        paragraphs: inserted,
        ...(tracked ? { revision: ADA } : {}),
      };
      const next = apply(original, op);
      const expected = ['Opening.', 'Closing.'];
      expected.splice(at, 0, 'New section', 'First addition.', 'Second addition.');
      expect(text(next)).toEqual([...expected]);
      for (const paragraph of originalParagraphs)
        expect(findNode(next, paragraph.id)).toBe(paragraph);
      const ids = paragraphs(next).map((node) => node.id);
      expect(new Set(ids).size).toBe(5);
      expect(
        ids
          .filter((id) => !originalParagraphs.some((node) => node.id === id))
          .some((id) => inserted.some((node) => node.id === id))
      ).toBe(false);
      const last = body(next).children.at(-1);
      expect(last?.kind !== 'textValue' ? last?.localName : undefined).toBe('sectPr');
      expect(canonicalOoxmlFingerprint(reopen(next))).toBe(canonicalOoxmlFingerprint(next));
      if (tracked) {
        const sites = collectRevisionSites(next);
        expect(sites).toHaveLength(6);
        expect(new Set(sites.map((site) => attribute(site.node, 'id'))).size).toBe(1);
        for (const site of sites) {
          expect(attribute(site.node, 'author')).toBe(ADA.author);
          expect(attribute(site.node, 'date')).toBe(ADA.date);
        }
        const accepted = apply(next, { op: 'acceptAllRevisions' });
        const rejected = apply(next, { op: 'rejectAllRevisions' });
        expect(text(accepted)).toEqual([...expected]);
        expect(text(rejected)).toEqual(['Opening.', 'Closing.']);
        expect(paragraphs(rejected).map((node) => node.id)).toEqual(
          originalParagraphs.map((node) => node.id)
        );
        for (const paragraph of originalParagraphs)
          expect(findNode(rejected, paragraph.id)).toBe(paragraph);
        expect(serializeOoxmlPart(rejected)).toBe(serializeOoxmlPart(original));
        expect(serializeOoxmlPart(apply(reopen(next), { op: 'rejectAllRevisions' }))).toBe(
          serializeOoxmlPart(original)
        );
      }
    });
  }
}

test('repeated seeds get fresh identities and independent native revision addresses', () => {
  let source = part(p('Tail'));
  const fragment = seed(p('New'));
  for (let n = 0; n < 2; n++)
    source = apply(source, {
      op: 'insertParagraphs',
      parentId: body(source).id,
      beforeId: null,
      paragraphs: fragment,
      revision: ADA,
    });
  expect(text(source)).toEqual(['Tail', 'New', 'New']);
  expect(new Set(paragraphs(source).map((node) => node.id)).size).toBe(3);
  expect(new Set(collectRevisionSites(source).map((site) => attribute(site.node, 'id'))).size).toBe(
    2
  );
  expect(text(apply(source, { op: 'rejectAllRevisions' }))).toEqual(['Tail']);
});

for (const xml of ['<w:p/>', '<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:p>'])
  test(`empty terminal inserted paragraphs reject completely: ${xml}`, () => {
    const original = part(p('Kept'));
    const next = apply(original, {
      op: 'insertParagraphs',
      parentId: body(original).id,
      beforeId: null,
      paragraphs: seed(xml),
      revision: ADA,
    });
    expect(paragraphs(next)).toHaveLength(2);
    expect(collectRevisionSites(next)).toHaveLength(1);
    expect(paragraphs(apply(next, { op: 'acceptAllRevisions' }))).toHaveLength(2);
    expect(serializeOoxmlPart(apply(next, { op: 'rejectAllRevisions' }))).toBe(
      serializeOoxmlPart(original)
    );
  });

test('terminal native deletion removes its emptied paragraph while keeping the preceding identity', () => {
  const original = part(p('Keep') + p('Delete'));
  const target = paragraphs(original)[1]!;
  let pending = apply(original, {
    op: 'deleteText',
    paragraphId: target.id,
    start: 0,
    end: 6,
    revision: ADA,
  });
  pending = apply(pending, {
    op: 'setParagraphMarkRevision',
    paragraphId: target.id,
    kind: 'del',
    revision: ADA,
  });
  expect(text(apply(pending, { op: 'acceptAllRevisions' }))).toEqual(['Keep']);
  expect(paragraphs(apply(pending, { op: 'acceptAllRevisions' }))[0]).toBe(paragraphs(original)[0]);
  expect(text(apply(pending, { op: 'rejectAllRevisions' }))).toEqual(['Keep', 'Delete']);
});

test.each([
  { name: 'sole required paragraph', xml: p(''), expected: [''] },
  { name: 'peer text', xml: p('Keep') + p('Peer'), expected: ['Keep', 'Peer'] },
  {
    name: 'opaque marker',
    xml: p('Keep') + '<w:p><w:bookmarkStart w:id="2" w:name="peer"/></w:p>',
    expected: ['Keep', ''],
  },
  {
    name: 'section properties',
    xml: p('Keep') + '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>',
    expected: ['Keep', ''],
  },
  {
    name: 'required paragraph after table',
    xml:
      p('Keep') +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc>' +
      p('Cell') +
      '</w:tc></w:tr></w:tbl><w:p/>',
    expected: ['Keep', ''],
  },
])('terminal mark resolution retains $name', ({ xml, expected }) => {
  const original = part(xml);
  const pending = apply(original, {
    op: 'setParagraphMarkRevision',
    paragraphId: paragraphs(original).at(-1)!.id,
    kind: 'ins',
    revision: ADA,
  });
  expect(text(apply(pending, { op: 'rejectAllRevisions' }))).toEqual([...expected]);
});

test.each([
  '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>',
  '<w:p><w:ins w:id="1" w:author="Other"><w:r><w:t>Existing</w:t></w:r></w:ins></w:p>',
  '<w:p><w:pPr><w:numPr><w:numId w:val="42"/></w:numPr></w:pPr></w:p>',
  '<w:p><w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
])('unsafe or resource-bearing paragraph seeds refuse atomically: %s', (xml) => {
  const source = part(p('Keep'));
  const fingerprint = canonicalOoxmlFingerprint(source);
  expect(
    applyTreeOp(source, {
      op: 'insertParagraphs',
      parentId: body(source).id,
      beforeId: null,
      paragraphs: seed(xml),
    })
  ).toMatchObject({ ok: false, reason: 'unsupported' });
  expect(canonicalOoxmlFingerprint(source)).toBe(fingerprint);
});

test('invalid anchors, attribution, payloads and budgets refuse before writes', () => {
  const source = part(p('Keep'));
  const valid: TreeDocOp = {
    op: 'insertParagraphs',
    parentId: body(source).id,
    beforeId: null,
    paragraphs: seed(p('New')),
  };
  const malformed = [
    { ...valid, parentId: 'missing' },
    { ...valid, beforeId: 'missing' },
    { ...valid, parentId: paragraphs(source)[0]!.id },
    { ...valid, paragraphs: [] },
    { ...valid, paragraphs: [null] },
    { ...valid, revision: { author: '' } },
    { ...valid, paragraphs: Array(50_001).fill(valid.paragraphs[0]) },
  ];
  const fingerprint = canonicalOoxmlFingerprint(source);
  for (const op of malformed) {
    expect(applyTreeOp(source, op as TreeDocOp).ok).toBe(false);
    expect(canonicalOoxmlFingerprint(source)).toBe(fingerprint);
  }
  let deep = valid.paragraphs[0]!;
  for (let n = 0; n < 66; n++)
    deep = { ...deep, children: [deep] } as unknown as OoxmlParagraphNode;
  expect(applyTreeOp(source, { ...valid, paragraphs: [deep] })).toMatchObject({
    ok: false,
    reason: 'fragment-too-deep',
  });
});

for (const restriction of [
  '<w:lock w:val="sdtContentLocked"/>',
  '<w:dataBinding w:storeItemID="{123}" w:xpath="/data"/>',
  '<w:temporary/>',
])
  test(`protected ancestry refuses insertion: ${restriction}`, () => {
    const source = part(
      `<w:sdt><w:sdtPr>${restriction}</w:sdtPr><w:sdtContent>${p('Inside')}</w:sdtContent></w:sdt>`
    );
    const control = body(source).children[0] as OoxmlElement;
    const content = control.children.find(
      (node) => node.kind !== 'textValue' && node.localName === 'sdtContent'
    )!;
    const result = applyTreeOp(source, {
      op: 'insertParagraphs',
      parentId: content.id,
      beforeId: null,
      paragraphs: seed(p('New')),
    });
    expect(result.ok).toBe(false);
  });

test('a locked sibling does not prevent insertion outside that control', () => {
  const source = part(
    `<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>${p('Locked')}</w:sdtContent></w:sdt>` +
      p('Tail')
  );
  const next = apply(source, {
    op: 'insertParagraphs',
    parentId: body(source).id,
    beforeId: null,
    paragraphs: seed(p('New')),
  });
  expect(text(next)).toEqual(['Tail', 'New']);
  expect(body(next).children[0]).toBe(body(source).children[0]);
});

test('fresh Word paragraph IDs remain unique across repeated seeds and saved reopen', () => {
  const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const read = readOoxmlPart(
    serializeOoxmlPart(part(p('Keep'))).replace('<w:document', `<w:document xmlns:w14="${w14}"`),
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  let next = read.part;
  const fragment = seed(p('New'));
  for (let n = 0; n < 2; n++)
    next = apply(next, {
      op: 'insertParagraphs',
      parentId: body(next).id,
      beforeId: null,
      paragraphs: fragment,
      revision: ADA,
    });
  const ids = paragraphs(next)
    .slice(1)
    .map(
      (node) =>
        node.attributes.find((attr) => attr.namespaceUri === w14 && attr.localName === 'paraId')
          ?.value
    );
  expect(ids.every((id) => typeof id === 'string' && /^[0-9A-F]{8}$/.test(id))).toBe(true);
  expect(new Set(ids).size).toBe(2);
  expect(serializeOoxmlPart(reopen(next))).toBe(serializeOoxmlPart(next));
});

test('new tracked paragraphs in a native table cell preserve its existing last paragraph', () => {
  const source = part(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc>' +
      p('Cell') +
      '</w:tc></w:tr></w:tbl>' +
      p('Tail')
  );
  const table = body(source).children[0] as OoxmlElement;
  const row = table.children.find((node) => node.kind === 'tableRow') as OoxmlElement;
  const cell = row.children.find((node) => node.kind === 'tableCell') as OoxmlElement;
  const originalParagraph = cell.children.find((node) => node.kind === 'paragraph')!;
  const inserted = apply(source, {
    op: 'insertParagraphs',
    parentId: cell.id,
    beforeId: originalParagraph.id,
    paragraphs: seed(p('New cell')),
    revision: ADA,
  });
  expect(findNode(inserted, originalParagraph.id)).toBe(originalParagraph);
  const rejected = apply(inserted, { op: 'rejectAllRevisions' });
  expect(serializeOoxmlPart(rejected)).toBe(serializeOoxmlPart(source));
  const accepted = apply(inserted, { op: 'acceptAllRevisions' });
  expect(findNode(accepted, originalParagraph.id)).toBe(originalParagraph);
  const acceptedCell = findNode(accepted, cell.id) as OoxmlElement;
  expect(
    acceptedCell.children
      .filter((node) => node.kind === 'paragraph')
      .map((node) => paragraphTextOf(accepted, node.id))
  ).toEqual(['New cell', 'Cell']);
});

test('inserting around unknown imported paragraph metadata leaves its exact tree intact', () => {
  const source = part(
    '<w:p xmlns:foreign="urn:foreign" foreign:paragraph="untouched"><w:pPr><foreign:metadata/></w:pPr><w:r foreign:run="untouched"><w:t>Kept</w:t></w:r></w:p>'
  );
  const original = paragraphs(source)[0]!;
  const next = apply(source, {
    op: 'insertParagraphs',
    parentId: body(source).id,
    beforeId: original.id,
    paragraphs: seed(p('New')),
    revision: ADA,
  });
  expect(findNode(next, original.id)).toBe(original);
  expect(serializeOoxmlPart(apply(next, { op: 'rejectAllRevisions' }))).toBe(
    serializeOoxmlPart(source)
  );
});

for (const properties of [
  '',
  '<w:pPr/>',
  '<w:pPr><w:rPr/></w:pPr>',
  '<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>',
])
  for (const saved of [false, true])
    test(`inserted paragraph retains literal property presence after acceptance: ${properties}/${saved}`, () => {
      const original = part(p('Kept'));
      const fragment = seed(`<w:p>${properties}<w:r><w:t>Copied</w:t></w:r></w:p>`);
      const op = {
        op: 'insertParagraphs' as const,
        parentId: body(original).id,
        beforeId: null,
        paragraphs: fragment,
      };
      const direct = apply(original, op);
      const pending = apply(original, { ...op, revision: ADA });
      const input = saved ? reopen(pending) : pending;
      expect(canonicalOoxmlFingerprint(apply(input, { op: 'acceptAllRevisions' }))).toBe(
        canonicalOoxmlFingerprint(direct)
      );
      expect(canonicalOoxmlFingerprint(apply(input, { op: 'rejectAllRevisions' }))).toBe(
        canonicalOoxmlFingerprint(original)
      );
    });
