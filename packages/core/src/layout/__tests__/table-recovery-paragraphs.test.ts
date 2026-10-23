import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  applyTreeOp,
  paragraphTextOf,
  readOoxmlPart,
  type OoxmlElement,
} from '@docx-editor.dev/core/store';
import { piecesOfParagraph } from '../field-projection.ts';
import {
  mergedFlowBlocks,
  storyBlocks,
  noteStoryBlocks,
  textboxStoryBlocks,
} from '../story-roots.ts';
import {
  revisionAuthorFilter,
  type RevisionAuthorFilter,
  type RevisionDisplayMode,
} from '../revision-projection.ts';
import { tableRecoveryParagraphs } from '../table-recovery-paragraphs.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const mark = (name: string, content = '') =>
  `<w:${name} w:id="17" w:author="AI" w:date="2026-10-23T00:00:00Z">${content}</w:${name}>`;
const row = (own = true, peer = true, kind = 'ins') =>
  `<w:tr><w:trPr>${mark(kind)}</w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p>${own ? `<w:pPr><w:rPr>${mark(kind)}</w:rPr></w:pPr>` : ''}${own ? mark(kind, '<w:r><w:t>Owned</w:t></w:r>') : '<w:r><w:t>Owned</w:t></w:r>'}${peer ? '<w:r><w:t xml:space="preserve"> Peer</w:t></w:r>' : ''}${own ? mark(kind, '<w:r><w:t xml:space="preserve"> again</w:t></w:r>') : ''}</w:p></w:tc></w:tr>`;
const table = (rows: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>${rows}</w:tbl>`;
function fixture(rows = row()) {
  const loaded = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${p('Before')}${table(rows)}${p('After')}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!loaded.ok) throw Error(loaded.reason);
  const body = loaded.part.root.children.find((node) => node.kind === 'body')!;
  const node = body.children.find((node) => node.kind === 'table')!;
  return { part: loaded.part, body, table: node };
}
function pieces(node: OoxmlElement, mode: RevisionDisplayMode, filter?: RevisionAuthorFilter) {
  return piecesOfParagraph(
    node,
    [],
    undefined,
    undefined,
    undefined,
    undefined,
    mode,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    filter
  );
}
function text(
  blocks: readonly OoxmlElement[],
  mode: RevisionDisplayMode,
  filter?: RevisionAuthorFilter
) {
  return blocks
    .filter((node) => node.kind === 'paragraph')
    .map((node) =>
      pieces(node, mode, filter)
        .map((piece) => piece.text)
        .join('')
    );
}
test('Original recovers foreign text at the rejection boundary with exact source offsets and no source writes', () => {
  const source = fixture();
  const before = canonicalOoxmlFingerprint(source.part);
  const recovered = tableRecoveryParagraphs(source.table, 'original');
  expect(recovered).toHaveLength(1);
  expect(
    pieces(recovered[0]!, 'original').map(({ text, start, end }) => ({ text, start, end }))
  ).toEqual([{ text: ' Peer', start: 5, end: 10 }]);
  expect(text(storyBlocks(source.part, 'original'), 'original')).toEqual([
    'Before',
    ' Peer',
    'After',
  ]);
  expect(paragraphTextOf(source.part, recovered[0]!.id)).toBe('Owned Peer again');
  expect(canonicalOoxmlFingerprint(source.part)).toBe(before);
  expect(tableRecoveryParagraphs(source.table, 'original')).toBe(recovered);
  expect(tableRecoveryParagraphs(source.table, 'proposed')).toHaveLength(0);
  expect(tableRecoveryParagraphs(source.table, 'all-markup')).toHaveLength(0);
});
for (const [own, peer] of [
  [false, true],
  [true, false],
] as const)
  test(`Original keeps ordinary row semantics and omits empty owned recovery; own=${own}, peer=${peer}`, () => {
    const source = fixture(row(own, peer));
    expect(tableRecoveryParagraphs(source.table, 'original')).toHaveLength(0);
    expect(text(storyBlocks(source.part, 'original'), 'original')).toEqual(['Before', 'After']);
  });
test('author filters and explicit per-node projection keep independent recovery separate', () => {
  const source = fixture();
  expect(
    tableRecoveryParagraphs(source.table, 'original', revisionAuthorFilter(['AI']))
  ).toHaveLength(0);
  const reject: RevisionAuthorFilter = {
    cacheKey: 'reject-all',
    hiddenAuthors: new Set(),
    includesNode: () => false,
    excludedNodeMode: () => 'original',
  };
  expect(
    text(mergedFlowBlocks(source.body.children, 'all-markup', reject), 'all-markup', reject)
  ).toEqual(['Before', ' Peer', 'After']);
  expect(tableRecoveryParagraphs(source.table, 'original')).toHaveLength(1);
});
test('Proposed recovers independent foreign text beside a removed deletion row', () => {
  const source = fixture(row(true, true, 'del'));
  expect(text(storyBlocks(source.part, 'proposed'), 'proposed')).toEqual([
    'Before',
    ' Peer',
    'After',
  ]);
  expect(tableRecoveryParagraphs(source.table, 'original')).toHaveLength(0);
});
for (const root of ['hdr', 'ftr', 'footnote', 'txbxContent', 'sdtContent'])
  test(`recovery has the same story boundary in ${root}`, () => {
    const parsed = readOoxmlPart(
      `<w:${root} xmlns:w="${W}"${root === 'footnote' ? ' w:id="1"' : ''}>${p('Before')}${table(row())}${p('After')}</w:${root}>`,
      { name: '/word/story.xml', contentType: 'application/xml' }
    );
    if (!parsed.ok) throw Error(parsed.reason);
    const node = parsed.part.root;
    const blocks =
      root === 'hdr' || root === 'ftr'
        ? storyBlocks(parsed.part, 'original')
        : root === 'footnote'
          ? noteStoryBlocks(node, 'original')
          : root === 'txbxContent'
            ? textboxStoryBlocks(node, 'original')
            : mergedFlowBlocks(node.children, 'original');
    expect(text(blocks, 'original')).toEqual(['Before', ' Peer', 'After']);
  });

test('a row retained for a nested table has no duplicate standalone recovery', () => {
  const inner = table(row().replace(`<w:trPr>${mark('ins')}</w:trPr>`, ''));
  const outer = `<w:tr><w:trPr>${mark('ins')}</w:trPr><w:tc>${inner}<w:p/></w:tc></w:tr>`;
  const source = fixture(outer);
  expect(tableRecoveryParagraphs(source.table, 'original')).toHaveLength(0);
});

for (const [kind, mode, op] of [
  ['ins', 'original', 'rejectAllRevisions'],
  ['del', 'proposed', 'acceptAllRevisions'],
] as const)
  test(`${mode} recovery matches its native ${op} paragraph result`, () => {
    const source = fixture(row(true, true, kind));
    const decided = applyTreeOp(source.part, { op });
    if (!decided.ok) throw Error(decided.reason);
    expect(text(storyBlocks(source.part, mode), mode)).toEqual(
      text(storyBlocks(decided.part, mode), mode)
    );
    expect(text(storyBlocks(decided.part, mode), mode)).toEqual(['Before', ' Peer', 'After']);
  });
