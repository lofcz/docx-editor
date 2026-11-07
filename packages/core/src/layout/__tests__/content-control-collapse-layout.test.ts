import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  contentControlContentNodeOf,
  findNode,
  removeNode,
  replaceNode,
  readOoxmlPart,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  createFixedMeasurer,
  createLayoutSession,
  createParagraphLayoutCache,
  layoutSemanticDocument,
  paragraphFragmentsOf,
  type SemanticLayout,
} from '../index.ts';
import { buildNumberingIndex } from '../numbering-index.ts';
import { createContentControlCollapseState } from '../content-control-collapse-state.ts';
import type { PendingLine } from '../paragraph-flow.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const sdt = (body: string) => `<w:sdt><w:sdtPr/><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
function read(xml: string, name = '/word/document.xml'): OoxmlPart {
  const result = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const load = (body: string) =>
  read(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
const options = {
  measurer: createFixedMeasurer(6, 14),
  geometry: {
    width: 300,
    height: 120,
    margin: { top: 10, right: 10, bottom: 10, left: 10 },
  },
};
const shape = (layout: SemanticLayout) =>
  layout.pages.map((page) => ({
    box: page.box,
    fragments: paragraphFragmentsOf(page).map((fragment) => ({
      id: fragment.paragraphId,
      box: fragment.box,
      lines: fragment.lines,
      marker: fragment.marker,
    })),
  }));
function request(part: OoxmlPart) {
  const owner = body(part).children.find((node) => node.kind === 'contentControl')!;
  return new Map([[owner.id, contentControlContentNodeOf(owner)!.children[0]!.id]]);
}
function body(part: OoxmlPart) {
  const node = bodyStoryRoot(part)!;
  if (node.kind === 'textValue') throw Error('Missing body');
  return node;
}

test('folds invalidate paginated flow while unchanged paragraph measurements remain cached', () => {
  const part = load(
    p('Before') +
      sdt(
        p('Summary') + Array.from({ length: 30 }, (_, i) => p(`Hidden paragraph ${i}`)).join('')
      ) +
      p('After')
  );
  const cache = createParagraphLayoutCache<readonly PendingLine[]>();
  const session = createLayoutSession();
  const collapse = createContentControlCollapseState();
  const warm = () =>
    layoutSemanticDocument(part, 1, {
      ...options,
      cache,
      session,
      contentControlCollapse: collapse.option(),
    });
  const opened = warm();
  expect(opened.pages.length).toBeGreaterThan(2);
  const misses = cache.stats.misses;
  expect(collapse.set(part, request(part)).ok).toBe(true);
  const closed = warm();
  expect(closed.pages).toHaveLength(1);
  expect(shape(closed)).toEqual(
    shape(
      layoutSemanticDocument(part, 1, { ...options, contentControlCollapse: collapse.option() })
    )
  );
  expect(cache.stats.misses).toBe(misses);
  expect(warm().pages).toBe(closed.pages);
  expect(collapse.set(part, new Map()).ok).toBe(true);
  expect(shape(warm())).toEqual(shape(opened));
  expect(cache.stats.misses).toBe(misses);
});

test('numbered items and their references retain canonical counters across hidden paragraphs', () => {
  const numbering = read(
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '</w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
    '/word/numbering.xml'
  );
  const numbered = (text: string, bookmark = '') =>
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>' +
    (bookmark ? `<w:bookmarkStart w:id="1" w:name="${bookmark}"/>` : '') +
    `<w:r><w:t>${text}</w:t></w:r>` +
    (bookmark ? '<w:bookmarkEnd w:id="1"/>' : '') +
    '</w:p>';
  const part = load(
    sdt(p('Summary') + numbered('Hidden item', 'Target')) +
      numbered('Visible item') +
      '<w:p><w:fldSimple w:instr=" REF Target \\r "/></w:p>'
  );
  const collapse = createContentControlCollapseState();
  expect(collapse.set(part, request(part)).ok).toBe(true);
  const layout = layoutSemanticDocument(part, 1, {
    ...options,
    numberingIndex: buildNumberingIndex(numbering.root),
    contentControlCollapse: collapse.option(),
  });
  const paragraphs = layout.pages.flatMap((page) => paragraphFragmentsOf(page));
  const text = (fragment: (typeof paragraphs)[number]) =>
    fragment.lines.flatMap((line) => line.spans.map((span) => span.text)).join('');
  expect(paragraphs.map(text)).toEqual(['Summary', 'Visible item', '1']);
  expect(paragraphs.find((fragment) => text(fragment) === 'Visible item')!.marker?.text).toBe('2.');
});

test('an unchanged owner moved into an unsupported container loses its local fold', () => {
  const part = load(
    sdt(p('Summary') + p('Body')) + `<w:customXml>${p('Foreign container')}</w:customXml>`
  );
  const requested = request(part);
  const owner = findNode(part, [...requested.keys()][0]!)!;
  if (owner.kind !== 'contentControl') throw Error('Missing control');
  const container = body(part).children.find(
    (node) => node.kind !== 'textValue' && node.localName === 'customXml'
  )!;
  if (container.kind === 'textValue') throw Error('Missing container');
  const collapse = createContentControlCollapseState();
  expect(collapse.set(part, requested).ok).toBe(true);
  const removed = removeNode(part, owner.id);
  if (!removed.ok) throw Error('Cannot detach owner');
  const moved = replaceNode(removed.part, container.id, {
    ...container,
    children: [owner, ...container.children],
  });
  if (!moved.ok) throw Error('Cannot place owner');
  expect(findNode(moved.part, owner.id)).toBe(owner);
  expect(collapse.reconcile(moved.part)).toBe(true);
  expect(collapse.current()).toEqual([]);
});

test('oversized replacements refuse before changing an existing local fold', () => {
  const part = load(sdt(p('Summary') + p('Body')));
  const collapse = createContentControlCollapseState();
  expect(collapse.set(part, request(part)).ok).toBe(true);
  const snapshot = collapse.current();
  const oversized = new Map(Array.from({ length: 4097 }, (_, i) => [`owner-${i}`, 'summary']));
  expect(collapse.set(part, oversized)).toEqual({ ok: false, reason: 'limit-exceeded' });
  expect(collapse.current()).toBe(snapshot);
});
