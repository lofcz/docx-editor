import { expect, test } from 'bun:test';
import { zipDoc } from './canonical-primitive-journal-coverage-support.ts';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  replaceNode,
  serializeOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '../index.ts';
import { INSERTED_PARAGRAPH_PROPERTY_NS } from '../store/inserted-paragraph-property-presence.ts';

const revision = { author: 'ScioBot', date: '2026-10-19T12:00:00Z' };
const attributes = 'w:id="9" w:author="Imported" w:date="2026-01-01T00:00:00Z"';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
function load(body: string): OoxmlPackage {
  const result = readOoxmlPackage(zipDoc({ body }));
  if (!result.ok) throw Error(result.reason);
  return result.package;
}
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(`${result.reason}: ${result.detail}`);
  return result.part;
}
function reopen(pkg: OoxmlPackage): OoxmlPackage {
  const result = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!result.ok) throw Error(result.reason);
  return result.package;
}
function nodes(part: OoxmlPart, kind: OoxmlNode['kind']): OoxmlElement[] {
  const found: OoxmlElement[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (node.kind === kind) found.push(node);
    node.children.forEach(visit);
  };
  visit(part.root);
  return found;
}
const fingerprints = (pkg: OoxmlPackage) =>
  Object.fromEntries([...pkg.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]));
const source = (mode: 'none' | 'empty' | 'formatted') => {
  const row =
    mode === 'none' ? '' : mode === 'empty' ? '<w:trPr/>' : '<w:trPr><w:cantSplit/></w:trPr>';
  const cell =
    mode === 'none'
      ? ''
      : mode === 'empty'
        ? '<w:tcPr/>'
        : '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>';
  const pPr = mode === 'formatted' ? '<w:pPr><w:rPr><w:i/><w:sz w:val="26"/></w:rPr></w:pPr>' : '';
  return load(
    `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${row}${['Old A', 'Old B'].map((text) => `<w:tc>${cell}${paragraph(text).replace('<w:p>', `<w:p>${pPr}`)}</w:tc>`).join('')}</w:tr></w:tbl>`
  );
};

for (const mode of ['none', 'empty', 'formatted'] as const)
  for (const where of ['above', 'below'] as const)
    for (const populated of [false, true])
      test(`tracked row holders match complete direct and rejected oracles after DOCX reopening: ${mode}/${where}/${populated}`, () => {
        const pkg = source(mode),
          part = main(pkg),
          table = nodes(part, 'table')[0]!,
          row = nodes(part, 'tableRow')[0]!;
        const op: TreeDocOp = populated
          ? {
              op: 'authorTable',
              action: {
                kind: 'existing',
                tableId: table.id,
                mutation: {
                  kind: 'insertRows',
                  rowId: row.id,
                  location: where === 'above' ? 'before' : 'after',
                  count: 1,
                  values: [['New A', 'New B']],
                },
              },
            }
          : { op: 'insertTableRow', tableId: table.id, rowId: row.id, where };
        if (op.op !== 'authorTable' && op.op !== 'insertTableRow')
          throw Error('Wrong insertion kind');
        const direct = reopen(withPart(pkg, apply(part, op)));
        const tracked = reopen(withPart(pkg, apply(part, { ...op, revision })));
        expect(serializeOoxmlPart(main(tracked))).toContain(INSERTED_PARAGRAPH_PROPERTY_NS);
        const accepted = reopen(
          withPart(tracked, apply(main(tracked), { op: 'acceptAllRevisions' }))
        );
        const rejected = reopen(
          withPart(tracked, apply(main(tracked), { op: 'rejectAllRevisions' }))
        );
        expect(fingerprints(accepted)).toEqual(fingerprints(direct));
        expect(fingerprints(rejected)).toEqual(fingerprints(reopen(pkg)));
        expect(serializeOoxmlPart(main(accepted))).not.toContain(INSERTED_PARAGRAPH_PROPERTY_NS);
        expect(serializeOoxmlPart(main(rejected))).not.toContain(INSERTED_PARAGRAPH_PROPERTY_NS);
      });

for (const mode of ['none', 'empty', 'formatted'] as const)
  for (const where of ['left', 'right'] as const)
    test(`tracked column holders match complete direct and rejected oracles after DOCX reopening: ${mode}/${where}`, () => {
      const pkg = source(mode),
        part = main(pkg),
        table = nodes(part, 'table')[0]!;
      const grid = table.children.find(
        (node) => node.kind !== 'textValue' && node.localName === 'tblGrid'
      ) as OoxmlElement;
      const op = {
        op: 'insertTableColumn',
        tableId: table.id,
        gridColumnId: grid.children[0]!.id,
        where,
      } as const;
      const direct = reopen(withPart(pkg, apply(part, op)));
      const tracked = reopen(withPart(pkg, apply(part, { ...op, revision })));
      const accepted = reopen(
        withPart(tracked, apply(main(tracked), { op: 'acceptAllRevisions' }))
      );
      const rejected = reopen(
        withPart(tracked, apply(main(tracked), { op: 'rejectAllRevisions' }))
      );
      expect(fingerprints(accepted)).toEqual(fingerprints(direct));
      expect(fingerprints(rejected)).toEqual(fingerprints(reopen(pkg)));
      expect(serializeOoxmlPart(main(accepted))).not.toContain(INSERTED_PARAGRAPH_PROPERTY_NS);
    });

for (const metadata of [
  '',
  ` xmlns:presence="${INSERTED_PARAGRAPH_PROPERTY_NS}" presence:properties="future"`,
  ` xmlns:presence="${INSERTED_PARAGRAPH_PROPERTY_NS}" presence:properties="pPr"`,
])
  test(`legacy and unknown table holder hints retain their original resolution: ${metadata || 'absent'}`, () => {
    const pkg = load(
      `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:ins ${attributes}${metadata}/></w:trPr><w:tc><w:tcPr><w:cellIns ${attributes}${metadata}/></w:tcPr>${paragraph('Imported')}</w:tc></w:tr></w:tbl>`
    );
    const expected = load(
      `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>${paragraph('Imported')}</w:tc></w:tr></w:tbl>`
    );
    const accepted = reopen(withPart(pkg, apply(main(pkg), { op: 'acceptAllRevisions' })));
    expect(fingerprints(accepted)).toEqual(fingerprints(expected));
  });

for (const kind of ['trPr', 'tcPr'] as const)
  test(`an explicit ${kind} source holder survives an earlier mark-family native address`, () => {
    const pkg = source('none'),
      part = main(pkg),
      table = nodes(part, 'table')[0]!,
      row = nodes(part, 'tableRow')[0]!;
    const op = { op: 'insertTableRow', tableId: table.id, rowId: row.id, where: 'below' } as const;
    const direct = apply(part, op),
      tracked = apply(part, { ...op, revision });
    const fresh = nodes(tracked, 'tableRow').at(-1)!;
    const holders: OoxmlElement[] = [];
    const visit = (node: OoxmlNode): void => {
      if (node.kind === 'textValue') return;
      if (node.localName === kind) holders.push(node);
      node.children.forEach(visit);
    };
    visit(fresh);
    const holder = holders[0]!;
    const replaced = replaceNode(tracked, holder.id, {
      ...holder,
      id: `${part.name}#mark:prior-holder`,
    } as OoxmlElement);
    if (!replaced.ok) throw Error(replaced.reason);
    const accepted = apply(replaced.part, { op: 'acceptAllRevisions' });
    expect(canonicalOoxmlFingerprint(accepted)).toBe(canonicalOoxmlFingerprint(direct));
    expect(fingerprints(reopen(withPart(pkg, accepted)))).toEqual(
      fingerprints(reopen(withPart(pkg, direct)))
    );
  });
