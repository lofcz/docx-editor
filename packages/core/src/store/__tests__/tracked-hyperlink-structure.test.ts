import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import {
  findNode,
  createNodeIdAllocator,
  insertChildren,
  replaceChildren,
} from '../package/ooxml-edit.ts';
import { withPart } from '../index.ts';
import { cloneWithNewIds } from '../store/tree-op-nodes.ts';
import {
  TreeDocumentStore,
  collectRevisionSites,
  planRevisionBatch,
  revisionItemsOf,
} from '../index.ts';
import { applyTreeOp, validateTreeOp, type TreeDocOp } from '../store/tree-ops.ts';
const W = WML_NAMESPACE_URI,
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const revision = { author: 'Bot', date: '2026-10-21T21:00:00Z' };
const p = (text: string) =>
  `<w:p><w:pPr><w:rPr/></w:pPr><w:r><w:rPr/><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const run = (text: string, properties = '') =>
  `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const link = (text = 'the site', attrs = 'r:id="rId1"') =>
  `<w:hyperlink ${attrs}>${run(text, '<w:b/>')}</w:hyperlink>`;
function load(body: string) {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:keep="urn:keep"><w:body>${body}${p('Tail')}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const paragraph = (part: OoxmlPart) => kids(kids(part.root)[0]!)[0]!;
const firstLink = (part: OoxmlPart) =>
  kids(paragraph(part)).find((node) => node.kind === 'hyperlink')!;
const apply = (part: OoxmlPart, op: TreeDocOp) => {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${op.op}: ${result.reason}`);
  return result.part;
};
const decide = (part: OoxmlPart, action: 'accept' | 'reject') => {
  const batch = planRevisionBatch(part, action);
  expect(batch.result.skipped).toEqual([]);
  let result = part;
  for (const op of batch.ops) result = apply(result, op);
  expect(revisionItemsOf(result)).toEqual([]);
  return result;
};
function reopen(part: OoxmlPart) {
  const result = readOoxmlPart(serializeOoxmlPart(part), {
    name: part.name,
    contentType: part.contentType,
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
function visible(node: OoxmlNode, mode: 'original' | 'proposed'): string {
  if (node.kind === 'textValue') return node.value;
  if (
    (mode === 'original' && node.kind === 'revisionInsert') ||
    (mode === 'proposed' && node.kind === 'revisionDelete')
  )
    return '';
  return node.children.map((child) => visible(child, mode)).join('');
}
function unnested(node: OoxmlNode, inside = false): void {
  if (node.kind === 'textValue') return;
  const revision = node.kind === 'revisionInsert' || node.kind === 'revisionDelete';
  if (revision) expect(inside).toBe(false);
  for (const child of node.children) unnested(child, inside || revision);
}
type Print = string | Print[];
const JOINABLE = new Set(['t', 'delText', 'tab', 'br']);
const named = (node: Print, name: string): node is Print[] =>
  Array.isArray(node) && node[0] === 'element' && node[1] === W && node[2] === name;
const plain = (run: Print): run is Print[] =>
  named(run, 'r') &&
  (run[3] as Print[]).length === 0 &&
  (run[4] as Print[]).every((part) => named(part, 'rPr') || JOINABLE.has(String(part[2])));
const props = (run: Print[]) =>
  JSON.stringify((run[4] as Print[]).filter((part) => named(part, 'rPr')));
const textOf = (node: Print[]) =>
  String(((node[4] as Print[])[0] as Print[] | undefined)?.[1] ?? '');
/** One text element per text kind and stretch; `xml:space` follows from the joined text. */
function joinText(content: readonly Print[]): Print[] {
  const out: Print[] = [];
  for (const part of content) {
    const last = out.at(-1);
    const text = (node: Print | undefined): node is Print[] =>
      Array.isArray(node) && (node[2] === 't' || node[2] === 'delText');
    if (text(part) && text(last) && last[2] === part[2])
      out[out.length - 1] = [...last.slice(0, 3), [], [['text', textOf(last) + textOf(part)]]];
    else if (text(part)) out.push([...part.slice(0, 3), [], part[4]!]);
    else out.push(part);
  }
  return out;
}
/** Join adjacent same-format plain runs and their text, as a decision joins them. */
function joinRuns(node: Print): Print {
  if (!Array.isArray(node) || node[0] !== 'element') return node;
  const out: Print[] = [];
  for (const child of (node[4] as Print[]).map(joinRuns)) {
    const previous = out.at(-1);
    if (previous && plain(previous) && plain(child) && props(previous) === props(child))
      out[out.length - 1] = [
        ...previous.slice(0, 4),
        [
          ...(previous[4] as Print[]),
          ...(child[4] as Print[]).filter((part) => !named(part, 'rPr')),
        ],
      ];
    else out.push(child);
  }
  return [
    ...node.slice(0, 4),
    out.map((child) =>
      plain(child) ? [...child.slice(0, 4), joinText(child[4] as Print[])] : child
    ),
  ];
}
// Decisions join adjacent same-format runs, so the structural oracle compares joined runs.
const fingerprint = (part: Parameters<typeof canonicalOoxmlFingerprint>[0]) =>
  JSON.stringify(joinRuns(JSON.parse(canonicalOoxmlFingerprint(part)) as Print));

for (const separator of ['', '<w:proofErr w:type="spellStart"/>'])
  test(`tracked link groups deletion wrappers without merging source runs: ${separator || 'adjacent'}`, () => {
    const part = load(`<w:p><w:pPr/>${run('left')}${separator}${run('right')}</w:p>`);
    const operation: TreeDocOp = {
      op: 'insertHyperlink',
      paragraphId: paragraph(part).id,
      start: 0,
      end: 9,
      anchor: 'Place',
    };
    const pending = apply(part, { ...operation, revision });
    const deletions = kids(paragraph(pending)).filter((node) => node.kind === 'revisionDelete');
    expect(deletions).toHaveLength(1);
    const deletedRuns = deletions.flatMap(kids).filter((node) => node.kind === 'run');
    expect(deletedRuns.map((node) => visible(node, 'proposed'))).toEqual(['left', 'right']);
    expect(deletedRuns.every((node) => node.kind === 'run')).toBe(true);
    expect(fingerprint(decide(pending, 'reject'))).toBe(fingerprint(part));
    expect(fingerprint(decide(reopen(pending), 'reject'))).toBe(fingerprint(reopen(part)));
    expect(fingerprint(decide(pending, 'accept'))).toBe(fingerprint(apply(part, operation)));
  });

interface Fixture {
  name: string;
  body: string;
  operation: (part: OoxmlPart) => TreeDocOp;
  partitioned?: boolean;
}
const fixtures: Fixture[] = [];
for (const target of ['external', 'bookmark'] as const)
  for (const style of [false, true])
    fixtures.push({
      name: `insert ${target}, style ${style}`,
      body: `<w:p><w:pPr><w:keepNext/></w:pPr>${run('See ')}${run('the ', '<w:b/>')}${run('site', '<w:i/>')}${run(' and ')}${link('twin')}</w:p>`,
      operation: (part) => ({
        op: 'insertHyperlink',
        paragraphId: paragraph(part).id,
        start: 4,
        end: 12,
        ...(target === 'external' ? { relationshipId: 'rId9' } : { anchor: 'Place' }),
        tooltip: 'Tip',
        ...(style ? { styleId: 'Hyperlink' } : {}),
      }),
    });
for (const [start, end] of [
  [0, 6],
  [7, 13],
  [14, 20],
  [0, 20],
])
  fixtures.push({
    name: `insert partial ordinary run ${start}/${end}`,
    partitioned: start !== 0 || end !== 20,
    body: p('prefix middle suffix'),
    operation: (part) => ({
      op: 'insertHyperlink',
      paragraphId: paragraph(part).id,
      start: start!,
      end: end!,
      anchor: 'Place',
    }),
  });
fixtures.push({
  name: 'insert complete text and run extension data',
  body: '<w:p><w:pPr/><w:r keep:run="same"><w:rPr><w:b/><keep:format keep:value="preserved"/></w:rPr><w:t keep:text="same">word</w:t></w:r></w:p>',
  operation: (part) => ({
    op: 'insertHyperlink',
    paragraphId: paragraph(part).id,
    start: 0,
    end: 4,
    anchor: 'Place',
  }),
});
for (const target of [
  'r:id="rId1" w:tooltip="Old" w:history="1"',
  'w:anchor="Place" w:docLocation="old"',
])
  fixtures.push({
    name: `remove ${target}`,
    body: `<w:p><w:pPr/>${run('See ')}${link('the site', target)}${run(' and ')}${link('twin')}</w:p>`,
    operation: (part) => ({ op: 'removeHyperlink', linkId: firstLink(part).id }),
  });
for (const kind of ['insert', 'remove'] as const)
  fixtures.push({
    name: `${kind} proofing marks`,
    body: `<w:p><w:pPr/>${kind === 'remove' ? '<w:hyperlink r:id="rId1">' : ''}${run('the ')}<w:proofErr w:type="spellStart"/>${run('site')}<w:proofErr w:type="spellEnd"/>${kind === 'remove' ? '</w:hyperlink>' : ''}</w:p>`,
    operation: (part) =>
      kind === 'insert'
        ? {
            op: 'insertHyperlink',
            paragraphId: paragraph(part).id,
            start: 0,
            end: 8,
            anchor: 'Place',
          }
        : { op: 'removeHyperlink', linkId: firstLink(part).id },
  });
fixtures.push({
  name: 'insert partial text with run and text extension metadata',
  body: '<w:p><w:pPr/><w:r keep:run="same" w:rsidR="00000001"><w:rPr><w:b/><keep:format keep:value="preserved"/></w:rPr><w:t keep:text="same">prefix middle suffix</w:t></w:r></w:p>',
  operation: (part) => ({
    op: 'insertHyperlink',
    paragraphId: paragraph(part).id,
    start: 7,
    end: 13,
    anchor: 'Place',
  }),
  partitioned: true,
});
for (const fixture of fixtures)
  test(`tracked hyperlink ${fixture.name}: exact native decisions, one history unit, and saved decisions`, () => {
    const part = load(fixture.body),
      plain = fixture.operation(part),
      op = { ...plain, revision } as TreeDocOp,
      before = fingerprint(part),
      store = new TreeDocumentStore(part),
      tail = kids(kids(part.root)[0]!)[1]!;
    expect(validateTreeOp(part, op)).toBeNull();
    let publications = 0;
    const stop = store.subscribe(() => publications++);
    const result = store.transact((ctx) => ctx.apply(op));
    stop();
    if (!result.ok) throw new Error(result.reason);
    expect(publications).toBe(1);
    expect(store.historyDepth).toBe(1);
    const pending = store.part;
    expect(findNode(pending, tail.id)).toBe(tail);
    expect(revisionItemsOf(pending).length).toBeGreaterThan(0);
    unnested(pending.root);
    const native = apply(part, plain);
    expect(visible(paragraph(pending), 'original')).toBe(visible(paragraph(part), 'original'));
    expect(visible(paragraph(pending), 'proposed')).toBe(visible(paragraph(native), 'proposed'));
    const decided = decide(pending, 'accept');
    expect(fingerprint(decided)).toBe(fingerprint(native));
    const rejectedGolden = fixture.partitioned
      ? apply(native, { op: 'removeHyperlink', linkId: firstLink(native).id })
      : part;
    expect(fingerprint(decide(pending, 'reject'))).toBe(fingerprint(rejectedGolden));
    expect(visible(paragraph(rejectedGolden), 'original')).toBe(
      visible(paragraph(part), 'original')
    );
    expect(store.undo()).not.toBeNull();
    expect(store.part).toBe(part);
    expect(store.redo()).not.toBeNull();
    expect(fingerprint(store.part)).toBe(fingerprint(pending));
    const saved = reopen(pending),
      accepted = decide(saved, 'accept');
    expect(fingerprint(accepted)).toBe(fingerprint(reopen(native)));
    expect(fingerprint(decide(saved, 'reject'))).toBe(fingerprint(reopen(rejectedGolden)));
    expect(fingerprint(part)).toBe(before);
  });
for (const kind of ['insert', 'remove'] as const)
  for (const content of [
    '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="end"/></w:r>',
    '<w:r><w:drawing/></w:r>',
    '<w:r><w:footnoteReference w:id="3"/></w:r>',
    '<w:ins w:id="3" w:author="Other"><w:r><w:t>word</w:t></w:r></w:ins>',
    '<w:r><w:rPr><w:rPrChange w:id="3" w:author="Other"><w:rPr/></w:rPrChange></w:rPr><w:t>word</w:t></w:r>',
    '<w:bookmarkStart w:id="5" w:name="Keep"/><w:r><w:t>word</w:t></w:r><w:bookmarkEnd w:id="5"/>',
  ])
    test(`tracked hyperlink ${kind} refuses non-copyable source ${content.slice(0, 28)} atomically`, () => {
      const part = load(
          `<w:p><w:pPr/>${kind === 'remove' ? '<w:hyperlink r:id="rId1">' : ''}${content}${kind === 'remove' ? '</w:hyperlink>' : ''}</w:p>`
        ),
        p = paragraph(part),
        store = new TreeDocumentStore(part),
        op =
          kind === 'insert'
            ? {
                op: 'insertHyperlink',
                paragraphId: p.id,
                start: 0,
                end: Math.max(1, visible(p, 'proposed').length),
                anchor: 'Place',
                revision,
              }
            : { op: 'removeHyperlink', linkId: firstLink(part).id, revision };
      const before = store.part;
      let count = 0;
      const stop = store.subscribe(() => count++);
      const result = store.transact((ctx) => ctx.apply(op as TreeDocOp));
      stop();
      expect(result.ok).toBe(false);
      expect(store.part).toBe(before);
      expect(count).toBe(0);
      expect(store.canUndo).toBe(false);
    });

test('whole-link removal and insertion mint distinct operation stamps across package callbacks', () => {
  const part = load(`<w:p><w:pPr/>${link()}${run(' plain')}</w:p>`),
    store = new TreeDocumentStore(part);
  let first = new Set<string>();
  const result = store.transact((ctx) => {
    ctx.apply({ op: 'removeHyperlink', linkId: firstLink(part).id, revision });
    first = new Set(
      collectRevisionSites(ctx.currentPackage().parts.get(part.name)!).map(
        (site) =>
          site.node.attributes.find(
            (attribute) => attribute.namespaceUri === W && attribute.localName === 'id'
          )!.value
      )
    );
    ctx.applyPackage((pkg) => pkg);
    ctx.apply({
      op: 'insertHyperlink',
      paragraphId: paragraph(part).id,
      start: 16,
      end: 22,
      anchor: 'Place',
      revision,
    });
  });
  if (!result.ok) throw new Error(result.reason);
  const sites = collectRevisionSites(store.part),
    fresh = sites.filter(
      (site) =>
        !first.has(
          site.node.attributes.find(
            (attribute) => attribute.namespaceUri === W && attribute.localName === 'id'
          )!.value
        )
    );
  expect(fresh.length).toBeGreaterThan(0);
  expect(new Set(sites.map((site) => site.node.id)).size).toBe(sites.length);
});

for (const selection of ['fresh', 'mixed', 'partial-fresh'] as const)
  test(`tracked link composes with ${selection} text from this exact native transaction`, () => {
    const part = load(`<w:p><w:pPr/>${run('left ', '<w:b/>')}${run('right', '<w:i/>')}</w:p>`),
      p = paragraph(part),
      store = new TreeDocumentStore(part),
      insert = { op: 'insertText' as const, paragraphId: p.id, offset: 5, text: 'new word ' },
      range =
        selection === 'fresh'
          ? { start: 5, end: 14 }
          : selection === 'mixed'
            ? { start: 0, end: 14 }
            : { start: 7, end: 12 };
    const plain = { op: 'insertHyperlink' as const, paragraphId: p.id, ...range, anchor: 'Place' };
    const result = store.transact((ctx) => {
      ctx.apply({ ...insert, revision });
      ctx.applyPackage((pkg) => pkg);
      ctx.apply({ ...plain, revision });
    });
    if (!result.ok) throw new Error(result.reason);
    const pending = store.part,
      direct = apply(decide(apply(part, { ...insert, revision }), 'accept'), plain);
    unnested(pending.root);
    expect(visible(paragraph(pending), 'original')).toBe('left right');
    expect(visible(paragraph(pending), 'proposed')).toBe('left new word right');
    expect(fingerprint(decide(pending, 'accept'))).toBe(fingerprint(direct));
    const rejected = decide(pending, 'reject');
    expect(visible(paragraph(rejected), 'original')).toBe('left right');
    expect(revisionItemsOf(rejected)).toEqual([]);
    const saved = reopen(pending);
    expect(fingerprint(decide(saved, 'accept'))).toBe(fingerprint(reopen(direct)));
    expect(visible(paragraph(decide(saved, 'reject')), 'original')).toBe('left right');
    expect(store.undo()).not.toBeNull();
    expect(store.part).toBe(part);
    expect(store.redo()).not.toBeNull();
    expect(fingerprint(store.part)).toBe(fingerprint(pending));
  });

test('a link can be created and then removed inside one current-frame text proposal', () => {
  const part = load(`<w:p><w:pPr/>${run('left ')}${run('right')}</w:p>`),
    p = paragraph(part),
    store = new TreeDocumentStore(part);
  let linkId = '';
  const result = store.transact((ctx) => {
    ctx.apply({ op: 'insertText', paragraphId: p.id, offset: 5, text: 'new ', revision });
    ctx.apply({
      op: 'insertHyperlink',
      paragraphId: p.id,
      start: 5,
      end: 9,
      anchor: 'Place',
      revision,
    });
    const current = findNode(ctx.currentPackage().parts.get(part.name)!, p.id)!;
    linkId = kids(current).find((node) => node.kind === 'hyperlink')!.id;
    ctx.apply({ op: 'removeHyperlink', linkId, revision });
  });
  if (!result.ok) throw new Error(result.reason);
  unnested(store.part.root);
  expect(visible(paragraph(store.part), 'original')).toBe('left right');
  expect(visible(paragraph(store.part), 'proposed')).toBe('left new right');
  const direct = decide(
    apply(part, { op: 'insertText', paragraphId: p.id, offset: 5, text: 'new ', revision }),
    'accept'
  );
  expect(fingerprint(decide(store.part, 'accept'))).toBe(fingerprint(direct));
  expect(fingerprint(decide(store.part, 'reject'))).toBe(fingerprint(part));
});

test('matching a current native revision address cannot grant an imported wrapper transaction ownership', () => {
  const part = load(p('Word')),
    store = new TreeDocumentStore(part),
    first = paragraph(part),
    tail = kids(kids(part.root)[0]!)[1]!;
  let publications = 0;
  const stop = store.subscribe(() => publications++);
  const result = store.transact((ctx) => {
    ctx.apply({ op: 'insertText', paragraphId: first.id, offset: 0, text: 'new', revision });
    const current = ctx.currentPackage().parts.get(part.name)!,
      own = kids(findNode(current, first.id)!).find((node) => node.kind === 'revisionInsert')!;
    ctx.applyPackage((pkg) => {
      const source = pkg.parts.get(part.name)!,
        foreign = cloneWithNewIds(own, createNodeIdAllocator(source));
      if (foreign.kind === 'textValue' || own.kind === 'textValue')
        throw new Error('Missing native wrapper');
      expect(foreign.attributes).toEqual(own.attributes);
      expect(foreign.id).not.toBe(own.id);
      const inserted = insertChildren(source, tail.id, 1, [foreign]);
      if (!inserted.ok) throw new Error(JSON.stringify(inserted.issues));
      return withPart(pkg, inserted.part);
    });
    ctx.apply({
      op: 'insertHyperlink',
      paragraphId: tail.id,
      start: 0,
      end: 3,
      anchor: 'Place',
      revision,
    });
  });
  stop();
  expect(result).toMatchObject({ ok: false, reason: 'unsupported' });
  expect(store.part).toBe(part);
  expect(publications).toBe(0);
  expect(store.canUndo).toBe(false);
});

test('a link spans current-frame deletion gaps without copying old text or nesting content revisions', () => {
  const part = load(`<w:p><w:pPr/>${run('old ')}${run('right', '<w:i/>')}</w:p>`),
    p = paragraph(part),
    store = new TreeDocumentStore(part);
  const result = store.transact((ctx) => {
    ctx.apply({ op: 'deleteText', paragraphId: p.id, start: 0, end: 3, revision });
    ctx.apply({ op: 'insertText', paragraphId: p.id, offset: 3, text: 'new', revision });
    ctx.applyPackage((pkg) => {
      const source = pkg.parts.get(part.name)!,
        node = findNode(source, p.id)!;
      if (node.kind !== 'paragraph') throw new Error('Missing paragraph');
      const insert = node.children.find((child) => child.kind === 'revisionInsert')!,
        deleted = node.children.find((child) => child.kind === 'revisionDelete')!;
      const replaced = replaceChildren(source, p.id, [
        ...node.children.filter((child) => child.localName === 'pPr'),
        insert,
        deleted,
        ...node.children.filter(
          (child) => child.id !== insert.id && child.id !== deleted.id && child.localName !== 'pPr'
        ),
      ]);
      if (!replaced.ok) throw new Error(JSON.stringify(replaced.issues));
      return withPart(pkg, replaced.part);
    });
    ctx.apply({
      op: 'insertHyperlink',
      paragraphId: p.id,
      start: 0,
      end: 12,
      anchor: 'Place',
      revision,
    });
  });
  if (!result.ok) throw new Error(result.reason);
  const pending = store.part;
  unnested(pending.root);
  expect(visible(paragraph(pending), 'original')).toBe('old right');
  expect(visible(paragraph(pending), 'proposed')).toBe('new right');
  const text = decide(
    apply(apply(part, { op: 'deleteText', paragraphId: p.id, start: 0, end: 3, revision }), {
      op: 'insertText',
      paragraphId: p.id,
      offset: 3,
      text: 'new',
      revision,
    }),
    'accept'
  );
  const direct = apply(text, {
    op: 'insertHyperlink',
    paragraphId: p.id,
    start: 0,
    end: 9,
    anchor: 'Place',
  });
  expect(fingerprint(decide(pending, 'accept'))).toBe(fingerprint(direct));
  expect(visible(paragraph(decide(pending, 'reject')), 'original')).toBe('old right');
  const saved = reopen(pending);
  expect(fingerprint(decide(saved, 'accept'))).toBe(fingerprint(reopen(direct)));
  expect(visible(paragraph(decide(saved, 'reject')), 'original')).toBe('old right');
});

test('late native link refusal preserves an earlier valid proposal and prior redo without publication', () => {
  const part = load(p('word')),
    store = new TreeDocumentStore(part),
    id = paragraph(part).id;
  expect(
    store.transact((ctx) =>
      ctx.apply({ op: 'insertText', paragraphId: id, offset: 4, text: ' redo' })
    ).ok
  ).toBe(true);
  expect(store.undo()).not.toBeNull();
  expect(store.canRedo).toBe(true);
  const before = store.part;
  let publications = 0;
  const stop = store.subscribe(() => publications++);
  const result = store.transact((ctx) => {
    ctx.apply({
      op: 'insertHyperlink',
      paragraphId: id,
      start: 0,
      end: 4,
      anchor: 'Place',
      revision,
    });
    ctx.apply({ op: 'removeHyperlink', linkId: 'missing', revision });
  });
  stop();
  expect(result.ok).toBe(false);
  expect(store.part).toBe(before);
  expect(store.canRedo).toBe(true);
  expect(publications).toBe(0);
  expect(store.redo()).not.toBeNull();
  expect(visible(paragraph(store.part), 'proposed')).toBe('word redo');
});

test('tracked unlink refuses a valid native subrange without silently applying it directly', () => {
  const part = load(`<w:p>${link()}</w:p>`),
    plain = {
      op: 'removeHyperlink' as const,
      linkId: firstLink(part).id,
      range: { start: 2, end: 6 },
    },
    store = new TreeDocumentStore(part);
  expect(validateTreeOp(part, plain)).toBeNull();
  expect(validateTreeOp(part, { ...plain, revision })).toBe('unsupported');
  expect(store.transact((ctx) => ctx.apply({ ...plain, revision }))).toMatchObject({
    ok: false,
    reason: 'unsupported',
  });
  expect(store.part).toBe(part);
  expect(store.canUndo).toBe(false);
});

test('native tracked link proposals retain target, attribution, surrogate, and content-lock validation', () => {
  const plain = load(p('A😀B')),
    id = paragraph(plain).id,
    locked = load(
      '<w:p><w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
        run('word') +
        '</w:sdtContent></w:sdt></w:p>'
    );
  const variants: [OoxmlPart, TreeDocOp, string][] = [
    [
      plain,
      {
        op: 'insertHyperlink',
        paragraphId: id,
        start: 0,
        end: 4,
        anchor: 'Place',
        revision: { author: '' },
      },
      'invalidArgs',
    ],
    [
      plain,
      { op: 'insertHyperlink', paragraphId: id, start: 1, end: 2, anchor: 'Place', revision },
      'splits-surrogate-pair',
    ],
    [
      plain,
      { op: 'insertHyperlink', paragraphId: id, start: 0, end: 4, revision },
      'invalid-property-value',
    ],
    [
      locked,
      {
        op: 'insertHyperlink',
        paragraphId: paragraph(locked).id,
        start: 0,
        end: 4,
        anchor: 'Place',
        revision,
      },
      'locked',
    ],
  ];
  for (const [part, op, reason] of variants) {
    const store = new TreeDocumentStore(part);
    expect(validateTreeOp(part, op)).toBe(reason as never);
    expect(store.transact((ctx) => ctx.apply(op))).toMatchObject({ ok: false, reason });
    expect(store.part).toBe(part);
    expect(store.canUndo).toBe(false);
  }
});

test('retarget composes exact current text sites and retains link extension attributes', () => {
  const part = load(
    `<w:p><w:pPr/><w:hyperlink r:id="rId1" w:tooltip="Old" keep:link="kept">${run('old ')}${run('right')}</w:hyperlink></w:p>`
  );
  const id = paragraph(part).id,
    store = new TreeDocumentStore(part);
  const insert: TreeDocOp = {
    op: 'insertText',
    paragraphId: id,
    offset: 4,
    text: 'new ',
    revision,
    bias: 'left',
  };
  const retarget: TreeDocOp = {
    op: 'setHyperlinkTarget',
    linkId: firstLink(part).id,
    relationshipId: 'rId9',
    tooltip: 'New',
  };
  const result = store.transact((ctx) => {
    ctx.apply(insert);
    const current = ctx.currentPackage().parts.get(part.name)!;
    expect(
      kids(findNode(current, firstLink(part).id)!).some((node) => node.kind === 'revisionInsert')
    ).toBe(true);
    ctx.apply({ ...retarget, revision });
  });
  expect(result.ok).toBe(true);
  unnested(store.part.root);
  expect(visible(paragraph(store.part), 'original')).toBe('old right');
  expect(visible(paragraph(store.part), 'proposed')).toBe('old new right');
  const expected = apply(decide(apply(part, insert), 'accept'), retarget);
  expect(fingerprint(decide(store.part, 'accept'))).toBe(fingerprint(expected));
  expect(fingerprint(decide(store.part, 'reject'))).toBe(fingerprint(part));
  const pending = fingerprint(store.part);
  expect(store.undo()).not.toBeNull();
  expect(store.part).toBe(part);
  expect(store.redo()).not.toBeNull();
  expect(fingerprint(store.part)).toBe(pending);
  expect(fingerprint(decide(reopen(store.part), 'accept'))).toBe(fingerprint(expected));
});

for (const marker of ['', '<w:proofErr w:type="spellEnd"/>'])
  test(`unlink of current deleted content proposes nothing and a new link spans the old source wrapper: ${marker || 'plain'}`, () => {
    const part = load(
      `<w:p><w:pPr/>${run('before ')}${link('docs', 'r:id="rId1" keep:link="kept"').replace('</w:hyperlink>', `${marker}</w:hyperlink>`)}${run(' after')}</w:p>`
    );
    const id = paragraph(part).id,
      store = new TreeDocumentStore(part);
    const result = store.transact((ctx) => {
      ctx.apply({ op: 'deleteText', paragraphId: id, start: 7, end: 11, revision });
      const current = ctx.currentPackage().parts.get(part.name)!;
      ctx.apply({ op: 'removeHyperlink', linkId: firstLink(part).id, revision });
      expect(ctx.currentPackage().parts.get(part.name)).toBe(current);
      ctx.apply({ op: 'insertText', paragraphId: id, offset: 0, text: 'New ', revision });
      const content = findNode(ctx.currentPackage().parts.get(part.name)!, id)!;
      const length = (node: OoxmlNode): number =>
        node.kind === 'textValue'
          ? node.value.length
          : kids(node).reduce((n, child) => n + length(child), 0);
      ctx.apply({
        op: 'insertHyperlink',
        paragraphId: id,
        start: 0,
        end: length(content),
        anchor: 'Place',
        revision,
      });
    });
    expect(result.ok).toBe(true);
    unnested(store.part.root);
    expect(visible(paragraph(store.part), 'original')).toBe('before docs after');
    expect(visible(paragraph(store.part), 'proposed')).toBe('New before  after');
    const pending = reopen(store.part);
    expect(visible(paragraph(decide(pending, 'accept')), 'proposed')).toBe('New before  after');
    expect(fingerprint(decide(pending, 'reject'))).toBe(fingerprint(part));
  });

test('retargeting to the current target inside a text transaction adds no revision sites', () => {
  const part = load(
      `<w:p><w:hyperlink r:id="rId1">${run('old ')}${run('right')}</w:hyperlink></w:p>`
    ),
    id = paragraph(part).id,
    store = new TreeDocumentStore(part);
  const result = store.transact((ctx) => {
    ctx.apply({
      op: 'insertText',
      paragraphId: id,
      offset: 4,
      text: 'new ',
      revision,
      bias: 'left',
    });
    const current = ctx.currentPackage().parts.get(part.name)!;
    const sites = collectRevisionSites(current).map((site) => site.node.id);
    ctx.apply({
      op: 'setHyperlinkTarget',
      linkId: firstLink(part).id,
      relationshipId: 'rId1',
      revision,
    });
    expect(ctx.currentPackage().parts.get(part.name)).toBe(current);
    expect(
      collectRevisionSites(ctx.currentPackage().parts.get(part.name)!).map((site) => site.node.id)
    ).toEqual(sites);
  });
  expect(result.ok).toBe(true);
});

test('a deletion-only foreign source link cannot become current transaction ownership', () => {
  const part = load(
    `<w:p>${run('before ')}<w:hyperlink r:id="rId1"><w:del w:id="0" w:author="Bot" w:date="${revision.date}"><w:r><w:rPr/><w:delText>docs</w:delText></w:r></w:del></w:hyperlink>${run(' after')}</w:p>`
  );
  const id = paragraph(part).id,
    store = new TreeDocumentStore(part);
  const result = store.transact((ctx) => {
    ctx.apply({ op: 'insertText', paragraphId: id, offset: 0, text: 'New ', revision });
    ctx.apply({ op: 'removeHyperlink', linkId: firstLink(part).id, revision });
  });
  expect(result).toMatchObject({ ok: false, reason: 'unsupported' });
  expect(store.part).toBe(part);
  expect(store.canUndo).toBe(false);
});
