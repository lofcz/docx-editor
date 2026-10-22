import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  serializeOoxmlPart,
  validateOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '../index.ts';
import { runWithTransactionActor } from '../package/actor-scoped-ids.ts';
import { isEmptyRetiredParagraph } from '../package/retired-paragraph.ts';
import {
  zipDoc,
  openStore,
  captureOneJournal,
  replayAndCompare,
} from './canonical-primitive-journal-coverage-support.ts';

const p = (text: string) => '<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>';
function fixture(
  action: 'accept' | 'reject',
  properties: boolean,
  leading: boolean,
  keepColumns: boolean,
  nested = false
) {
  const owner =
      '<w:tc>' +
      (properties ? '<w:tcPr><w:shd w:fill="EEEEEE"/><x:keep x:val="retained"/></w:tcPr>' : '') +
      p('Keep') +
      '</w:tc>',
    content = nested
      ? '<w:sdt><w:sdtPr/><w:sdtContent>' + p('Gone') + '</w:sdtContent></w:sdt>'
      : p('Gone'),
    removed =
      '<w:tc><w:tcPr><w:shd w:fill="BBBBBB"/><x:sourceKeep x:val="original"><x:nested/></x:sourceKeep><w:' +
      (action === 'accept' ? 'cellDel' : 'cellIns') +
      ' w:id="9" w:author="Teacher"/></w:tcPr>' +
      content +
      '</w:tc>';
  return zipDoc({
    extraXmlns: 'xmlns:x="urn:test"',
    body:
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
      (leading ? removed + owner : owner + removed) +
      '</w:tr>' +
      (keepColumns
        ? '<w:tr><w:tc>' + p('One') + '</w:tc><w:tc>' + p('Two') + '</w:tc></w:tr>'
        : '') +
      '</w:tbl>' +
      p('Tail'),
  });
}
const nodes = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(nodes)),
];
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
function resolve(part: OoxmlPart, action: 'accept' | 'reject', actor?: string) {
  const result = runWithTransactionActor(actor, () =>
    applyTreeOp(part, { op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions' })
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
function word(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  return canonicalOoxmlFingerprint(main(read.package));
}
for (const action of ['accept', 'reject'] as const)
  for (const properties of [false, true])
    for (const leading of [false, true])
      for (const keepColumns of [false, true])
        test(
          'table geometry reuses existing shared identities and preserves the independent Word result; ' +
            [action, properties, leading, keepColumns],
          () => {
            const read = readOoxmlPackage(fixture(action, properties, leading, keepColumns));
            if (!read.ok) throw new Error(read.reason);
            const pkg = read.package,
              before = main(pkg),
              original = serializeOoxmlPart(before),
              source = nodes(before.root),
              marker = source.find(
                (node) => node.localName === (action === 'accept' ? 'cellDel' : 'cellIns')
              )!,
              firstColumn = source.find((node) => node.localName === 'gridCol')!,
              deletedParagraph = source.find(
                (node) =>
                  node.kind === 'paragraph' &&
                  node.children.some(
                    (child) =>
                      child.kind === 'run' &&
                      serializeOoxmlPart({ ...before, root: child }).includes('Gone')
                  )
              )!;
            const expected = word(withPart(pkg, resolve(before, action)));
            for (const actor of ['alice', 'bob']) {
              const after = resolve(before, action, actor),
                all = nodes(after.root),
                spans = all.filter((node) => node.localName === 'gridSpan');
              expect(validateOoxmlPart(after)).toEqual({ ok: true });
              expect(new Set(all.map((node) => node.id)).size).toBe(all.length);
              expect(spans).toHaveLength(keepColumns ? 1 : 0);
              if (keepColumns) expect(spans[0]!.id).toBe(marker.id);
              else
                expect(all.find((node) => node.localName === 'gridCol')!.id).toBe(firstColumn.id);
              if (properties) expect(serializeOoxmlPart(after)).toContain('x:val="retained"');
              if (!properties)
                expect(isEmptyRetiredParagraph(findNode(after, deletedParagraph.id)!)).toBe(true);
              expect(word(withPart(pkg, after))).toBe(expected);
            }
            expect(serializeOoxmlPart(before)).toBe(original);
          }
        );

test('a missing vertical property uses its resolved marker identity without losing extension attributes', () => {
  const bytes = zipDoc({
    body:
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      ['rest', 'cont']
        .map(
          (value) =>
            '<w:tr><w:tc><w:tcPr><w:shd w:fill="EEEEEE"/><w:cellMerge w:id="9" w:author="Teacher" w:vMerge="' +
            value +
            '"/></w:tcPr>' +
            p(value) +
            '</w:tc></w:tr>'
        )
        .join('') +
      '</w:tbl>',
  });
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error(read.reason);
  const before = main(read.package),
    markers = nodes(before.root)
      .filter((node) => node.localName === 'cellMerge')
      .map((node) => node.id);
  for (const actor of ['alice', 'bob']) {
    const after = resolve(before, 'accept', actor);
    expect(
      nodes(after.root)
        .filter((node) => node.localName === 'vMerge')
        .map((node) => node.id)
    ).toEqual(markers);
    expect(nodes(after.root).filter((node) => node.localName === 'shd')).toHaveLength(2);
    expect(validateOoxmlPart(after)).toEqual({ ok: true });
  }
});

for (const action of ['accept', 'reject'] as const)
  for (const properties of [false, true])
    for (const nested of [false, true])
      for (const leading of [false, true])
        for (const keepColumns of [false, true])
          test(
            'transferred properties and removed controls retain replayable original slots; ' +
              [action, properties, nested, leading, keepColumns],
            () => {
              const bytes = fixture(action, properties, leading, keepColumns, nested),
                source = openStore(bytes),
                replica = openStore(bytes),
                captured = captureOneJournal(source, () =>
                  source.transact(
                    { kind: 'body' },
                    (context) => {
                      context.apply({
                        op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
                      });
                    },
                    { actorId: 'alice' }
                  )
                );
              expect(captured.result.ok).toBe(true);
              expect(captured.journal).not.toBeNull();
              expect(
                captured.journal!.effects.some(
                  (effect) => effect.kind === 'spliceChildren' && effect.deleteCount > 0
                )
              ).toBe(true);
              replayAndCompare(replica, source, captured.journal!);
            }
          );
