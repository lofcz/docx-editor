import { expect, test } from 'bun:test';
import { storyBlocks } from '../../layout/story-roots.ts';
import { enumerateDocumentSections } from '../../layout/section-properties.ts';
import { runWithTransactionActor } from '../package/actor-scoped-ids.ts';
import { isJoinedParagraph } from '../package/retired-paragraph.ts';
import { zipDoc } from './canonical-primitive-journal-coverage-support.ts';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  readOoxmlPackage,
  serializeOoxmlPart,
  validateOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '../index.ts';

const NS = 'urn:docx-editor.dev:review:retired-paragraph:1';
const attrs = 'w:id="17" w:author="AI" w:date="2026-10-19T09:00:00Z"';
const p = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const inserted = `<w:p><w:pPr><w:rPr><w:ins ${attrs}/></w:rPr></w:pPr><w:ins ${attrs}><w:r><w:t>New</w:t></w:r></w:ins></w:p>`;
const joined = (value = 'Peer', properties = '') =>
  `<w:p><w:pPr><state:joined xmlns:state="${NS}"/>${properties}</w:pPr><w:r><w:t>${value}</w:t></w:r></w:p>`;
const table = (value = 'Cell') =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>${p(value)}</w:tc></w:tr></w:tbl>`;
function load(xml: string) {
  const result = readOoxmlPackage(zipDoc({ body: xml }));
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
const body = (part: OoxmlPart) =>
  part.root.children.find((node) => node.kind === 'body') as OoxmlElement;
function apply(part: OoxmlPart, op: TreeDocOp, actor: string | undefined = 'alice') {
  const result = runWithTransactionActor(actor, () => applyTreeOp(part, op));
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail}`);
  return result.part;
}
const text = (node: OoxmlNode): string =>
  node.kind === 'textValue' ? node.value : node.children.map(text).join('');
const shown = (part: OoxmlPart, mode: 'original' | 'proposed' | 'all-markup' = 'proposed') =>
  storyBlocks(part, mode).map(text);
function reopened(pkg: OoxmlPackage) {
  const result = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}

test('an actor-bound forward decision keeps the source identity and untouched target', () => {
  const source = load(p('Opening') + inserted + p('Closing')),
    part = main(source),
    id = body(part).children[1]!.id,
    tail = body(part).children[2]!;
  const next = apply(part, { op: 'rejectAllRevisions' });
  expect(validateOoxmlPart(next)).toEqual({ ok: true });
  expect(isJoinedParagraph(findNode(next, id)!)).toBe(true);
  expect(findNode(next, tail.id)).toBe(tail);
  expect(collectRevisionSites(next)).toHaveLength(0);
  expect(shown(next)).toEqual(['Opening', 'Closing']);
  const solo = runWithTransactionActor(undefined, () =>
    applyTreeOp(part, { op: 'rejectAllRevisions' })
  );
  if (!solo.ok) throw new Error(solo.reason);
  expect(findNode(solo.part, id)).toBeNull();
  const output = main(reopened(withPart(source, next)));
  expect(serializeOoxmlPart(output)).toBe(serializeOoxmlPart(solo.part));
});

test('late peer content joins forward exactly once in every view and clean Word export', () => {
  const source = load(p('Opening') + inserted + p('Closing')),
    part = main(source),
    id = body(part).children[1]!.id;
  const resolved = apply(part, { op: 'rejectAllRevisions' }),
    before = canonicalOoxmlFingerprint(resolved);
  expect(shown(main(reopened(withPart(source, resolved))))).toEqual(['Opening', 'Closing']);
  expect(canonicalOoxmlFingerprint(resolved)).toBe(before);
  const next = apply(
    resolved,
    { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' },
    'bob'
  );
  for (const mode of ['original', 'proposed', 'all-markup'] as const)
    expect(shown(next, mode)).toEqual(['Opening', 'PeerClosing']);
  const fingerprint = canonicalOoxmlFingerprint(next),
    output = main(reopened(withPart(source, next)));
  expect(serializeOoxmlPart(output)).not.toContain(NS);
  expect(serializeOoxmlPart(output)).toBe(
    serializeOoxmlPart(
      main(
        load(p('Opening') + '<w:p><w:r><w:t>Peer</w:t></w:r><w:r><w:t>Closing</w:t></w:r></w:p>')
      )
    )
  );
  expect(canonicalOoxmlFingerprint(next)).toBe(fingerprint);
});

test('several retired middle marks join into the same unchanged native target', () => {
  const source = load(p('Opening') + joined('One') + joined('Two') + p('Closing'));
  for (const mode of ['original', 'proposed', 'all-markup'] as const)
    expect(shown(main(source), mode)).toEqual(['Opening', 'OneTwoClosing']);
  expect(shown(main(reopened(source)))).toEqual(['Opening', 'OneTwoClosing']);
});

test('the surviving paragraph supplies formatting while source run formatting remains intact', () => {
  const source = load(
      joined('Peer', '<w:outlineLvl w:val="1"/>').replace('<w:r>', '<w:r><w:rPr><w:b/></w:rPr>') +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Closing</w:t></w:r></w:p>'
    ),
    part = main(source),
    target = body(part).children[1]!;
  if (target.kind !== 'paragraph') throw new Error('Missing survivor');
  const properties = target.children.find((node) => node.kind === 'paragraphProperties');
  for (const mode of ['original', 'proposed', 'all-markup'] as const)
    expect(
      storyBlocks(part, mode)[0]!.children.find((node) => node.kind === 'paragraphProperties')
    ).toBe(properties);
  const output = serializeOoxmlPart(main(reopened(source)));
  expect(output).not.toContain('outlineLvl');
  expect(output).toContain('<w:jc w:val="center"/>');
  expect(output).toContain('<w:b/>');
  expect(shown(main(reopened(source)))).toEqual(['PeerClosing']);
});

test('a source with no remaining forward target keeps the required final paragraph', () => {
  for (const value of ['', 'Peer']) {
    const source = load(joined(value));
    for (const mode of ['original', 'proposed', 'all-markup'] as const)
      expect(shown(main(source), mode)).toEqual([value]);
    expect(shown(main(reopened(source)))).toEqual([value]);
  }
});

for (const wrapper of ['control', 'customXml', 'cell'])
  test(`forward joins stay inside native ${wrapper} ownership through Word export`, () => {
    const content = joined() + p('Closing');
    const wrapped =
      wrapper === 'control'
        ? `<w:sdt><w:sdtContent>${content}</w:sdtContent></w:sdt>`
        : wrapper === 'customXml'
          ? `<w:customXml>${content}</w:customXml>`
          : table().replace(p('Cell'), content);
    const source = load(p('Opening') + wrapped);
    expect(shown(main(source))).toEqual(['Opening', 'PeerClosing']);
    expect(shown(main(reopened(source)))).toEqual(['Opening', 'PeerClosing']);
  });

test('a source in a content control cannot join a paragraph outside that owner', () => {
  const source = load(`<w:sdt><w:sdtContent>${joined()}</w:sdtContent></w:sdt>` + p('Closing'));
  expect(shown(main(source))).toEqual(['Peer', 'Closing']);
  expect(shown(main(reopened(source)))).toEqual(['Peer', 'Closing']);
});

test('resolved marks before a table use its first native cell as the join target', () => {
  const source = load(p('Opening') + joined() + table());
  for (const mode of ['original', 'proposed', 'all-markup'] as const)
    expect(shown(main(source), mode)).toEqual(['Opening', 'PeerCell']);
  expect(shown(main(reopened(source)))).toEqual(['Opening', 'PeerCell']);
});

for (const op of ['splitParagraph', 'splitParagraphMany'] as const)
  test(`${op} adds fresh marks before the original resolved forward mark`, () => {
    const source = load(p('Opening') + joined() + p('Closing')),
      part = main(source),
      id = body(part).children[1]!.id;
    const next = apply(
      part,
      op === 'splitParagraph'
        ? { op, paragraphId: id, offset: 2 }
        : { op, paragraphId: id, offsets: [1, 2] }
    );
    expect(validateOoxmlPart(next)).toEqual({ ok: true });
    expect(shown(next)).toEqual(
      op === 'splitParagraph' ? ['Opening', 'Pe', 'erClosing'] : ['Opening', 'P', 'e', 'erClosing']
    );
    expect(body(next).children.filter(isJoinedParagraph)).toHaveLength(1);
    expect(shown(main(reopened(withPart(source, next))))).toEqual(shown(next));
  });

test('an ordinary join inherits the surviving tail mark, including its resolved forward state', () => {
  const source = load(p('Opening') + p('Before') + joined() + p('Closing')),
    part = main(source),
    nodes = body(part).children;
  const next = apply(part, { op: 'joinParagraphs', firstId: nodes[1]!.id, secondId: nodes[2]!.id });
  expect(shown(next)).toEqual(['Opening', 'BeforePeerClosing']);
  expect(isJoinedParagraph(findNode(next, nodes[1]!.id)!)).toBe(true);
  expect(shown(main(reopened(withPart(source, next))))).toEqual(shown(next));
});

test('a copied whole paragraph cannot inherit the source forward-join identity', () => {
  const source = main(load(joined() + p('Closing'))),
    id = body(source).children[0]!.id,
    before = canonicalOoxmlFingerprint(source),
    target = main(load(p('Target'))),
    anchor = body(target).children[0]!.id;
  const pasted = apply(target, {
    op: 'insertFragment',
    paragraphId: anchor,
    offset: 6,
    blocks: [findNode(source, id)!],
    lastMarkCovered: true,
  });
  expect(shown(pasted)).toEqual(['TargetPeer', '']);
  expect(serializeOoxmlPart(pasted)).not.toContain(NS);
  expect(canonicalOoxmlFingerprint(source)).toBe(before);
});

test('a new paragraph format proposal cannot undo a resolved mark through its historical properties', () => {
  const source = load(p('Opening') + joined() + p('Closing')),
    part = main(source),
    id = body(part).children[1]!.id;
  const formatted = apply(part, {
    op: 'setParagraphProperties',
    paragraphId: id,
    properties: [{ localName: 'jc', attributes: { val: 'center' } }],
    revision: { author: 'User', date: '2026-10-19T10:00:00Z' },
  });
  for (const mode of ['original', 'proposed', 'all-markup'] as const)
    expect(shown(formatted, mode)).toEqual(['Opening', 'PeerClosing']);
  const rejected = apply(formatted, { op: 'rejectAllRevisions' });
  expect(isJoinedParagraph(findNode(rejected, id)!)).toBe(true);
  expect(shown(rejected)).toEqual(['Opening', 'PeerClosing']);
});

test('a retained source section mark has the same effective sections as the physical Word join', () => {
  const section = '<w:sectPr><w:pgSz w:w="10000" w:h="12000"/></w:sectPr>';
  const source = load(p('Opening') + joined('Peer', section) + p('Closing')),
    word = main(reopened(source));
  for (const mode of ['original', 'proposed', 'all-markup'] as const) {
    expect(shown(main(source), mode)).toEqual(shown(word, mode));
    expect(enumerateDocumentSections(main(source), mode)).toHaveLength(
      enumerateDocumentSections(word, mode).length
    );
  }
});

for (const marker of [
  '<state:joined flag="future"/>',
  '<state:joined><state:unknown/></state:joined>',
  '<state:other/>',
])
  test(`unknown or malformed resolved state remains opaque: ${marker}`, () => {
    const source = load(
      p('Opening') +
        `<w:p><w:pPr xmlns:state="${NS}">${marker}</w:pPr><w:r><w:t>Peer</w:t></w:r></w:p>` +
        p('Closing')
    );
    expect(shown(main(source))).toEqual(['Opening', 'Peer', 'Closing']);
    expect(serializeOoxmlPart(main(reopened(source)))).toContain(NS);
  });
