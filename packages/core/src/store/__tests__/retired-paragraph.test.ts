import { expect, test } from 'bun:test';
import { storyBlocks } from '../../layout/story-roots.ts';
import { runWithTransactionActor } from '../package/actor-scoped-ids.ts';
import { isRetiredParagraph } from '../package/retired-paragraph.ts';
import { storyParagraphs } from '../package/story-blocks.ts';
import { collectFlowBlocks, walkAllStoryParagraphs } from '../package/content-control-walk.ts';
import { zipDoc } from './canonical-primitive-journal-coverage-support.ts';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
  validateOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '../index.ts';

const NS = 'urn:docx-editor.dev:review:retired-paragraph:1';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const rev = 'w:id="17" w:author="AI" w:date="2026-10-19T09:00:00Z"';
const inserted = `<w:p><w:pPr><w:rPr><w:ins ${rev}/></w:rPr></w:pPr><w:ins ${rev}><w:r><w:t>New</w:t></w:r></w:ins></w:p>`;
function load(body: string) {
  const read = readOoxmlPackage(zipDoc({ body }));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
const body = (part: OoxmlPart) =>
  part.root.children.find((node) => node.kind === 'body') as OoxmlElement;
function apply(part: OoxmlPart, op: TreeDocOp, actor: string | null = 'alice') {
  const result = runWithTransactionActor(actor ?? undefined, () => applyTreeOp(part, op));
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail}`);
  return result.part;
}
function rejected(source = load(p('Keep') + inserted)) {
  return withPart(source, apply(main(source), { op: 'rejectAllRevisions' }));
}
function reopened(pkg: OoxmlPackage) {
  const result = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}

test('collaborative rejection keeps a non-rendered native paragraph identity and no pending revision', () => {
  const source = load(p('Keep') + inserted),
    original = main(source),
    id = body(original).children[1]!.id;
  const next = main(rejected(source));
  expect(validateOoxmlPart(next)).toEqual({ ok: true });
  expect(findNode(next, id)).not.toBeNull();
  expect(isRetiredParagraph(findNode(next, id)!)).toBe(true);
  expect(collectRevisionSites(next)).toHaveLength(0);
  expect(storyParagraphs(body(next)).map((node) => node.id)).toEqual([
    body(original).children[0]!.id,
  ]);
  expect(storyBlocks(next).map((node) => node.id)).toEqual([body(original).children[0]!.id]);
  expect(findNode(next, body(original).children[0]!.id)).toBe(body(original).children[0]);
  expect(serializeOoxmlPart(next)).toContain(NS);
});

test('Word export removes the empty collaboration identity and marker without changing the live package', () => {
  const pending = load(p('Keep') + inserted),
    next = rejected(pending);
  const fingerprint = canonicalOoxmlFingerprint(main(next));
  const output = main(reopened(next));
  expect(serializeOoxmlPart(output)).not.toContain(NS);
  expect(storyParagraphs(body(output)).map((node) => paragraphTextOf(output, node.id))).toEqual([
    'Keep',
  ]);
  expect(serializeOoxmlPart(output)).toBe(serializeOoxmlPart(main(load(p('Keep')))));
  expect(canonicalOoxmlFingerprint(main(next))).toBe(fingerprint);
  expect(applyTreeOp(main(next), { op: 'rejectAllRevisions' })).toMatchObject({
    ok: false,
    reason: 'unknown-revision',
  });
  expect(canonicalOoxmlFingerprint(main(next))).toBe(fingerprint);
});

test('peer text reuses the retained native identity, renders once, and exports as a normal Word paragraph', () => {
  const next = rejected(),
    part = main(next),
    id = body(part).children[1]!.id;
  const changed = apply(
    part,
    { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' },
    'bob'
  );
  expect(storyParagraphs(body(changed)).map((node) => node.id)).toEqual(
    body(changed).children.map((node) => node.id)
  );
  expect(storyBlocks(changed).map((node) => node.id)).toContain(id);
  expect(paragraphTextOf(changed, id)).toBe('Peer');
  const exported = main(reopened(withPart(next, changed)));
  expect(serializeOoxmlPart(exported)).not.toContain(NS);
  expect(storyParagraphs(body(exported)).map((node) => paragraphTextOf(exported, node.id))).toEqual(
    ['Keep', 'Peer']
  );
});

test('Word export removes retired metadata from native formatting history and preserves its accept/reject behavior', () => {
  const pkg = rejected(),
    part = main(pkg),
    id = body(part).children[1]!.id;
  const peer = apply(part, { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' }, 'bob');
  const formatted = apply(
    peer,
    {
      op: 'setParagraphProperties',
      paragraphId: id,
      properties: [{ localName: 'jc', attributes: { val: 'right' } }],
      revision: { author: 'Peer' },
    },
    'bob'
  );
  expect(collectRevisionSites(formatted)).toHaveLength(1);
  expect(serializeOoxmlPart(formatted).split(NS).length).toBeGreaterThan(2);
  const saved = main(reopened(withPart(pkg, formatted)));
  expect(serializeOoxmlPart(saved)).not.toContain(NS);
  expect(collectRevisionSites(saved)).toHaveLength(1);
  expect(storyParagraphs(body(saved)).map((node) => paragraphTextOf(saved, node.id))).toEqual([
    'Keep',
    'Peer',
  ]);
  const accepted = apply(saved, { op: 'acceptAllRevisions' });
  const rejectedFormat = apply(saved, { op: 'rejectAllRevisions' });
  expect(serializeOoxmlPart(accepted)).toContain('<w:jc w:val="right"/>');
  expect(serializeOoxmlPart(rejectedFormat)).not.toContain('<w:jc');
  expect(collectRevisionSites(accepted)).toHaveLength(0);
  expect(collectRevisionSites(rejectedFormat)).toHaveLength(0);
});

for (const offset of [0, 4])
  test(`an ordinary split at ${offset} gives both resulting paragraphs their own visible marks`, () => {
    const pkg = rejected(),
      part = main(pkg),
      id = body(part).children[1]!.id;
    const peer = apply(part, { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' }, 'bob');
    const split = apply(peer, { op: 'splitParagraph', paragraphId: id, offset }, 'bob');
    const expected = offset === 0 ? ['Keep', '', 'Peer'] : ['Keep', 'Peer', ''];
    expect(storyParagraphs(body(split)).map((node) => paragraphTextOf(split, node.id))).toEqual(
      expected
    );
    expect(storyBlocks(split)).toHaveLength(3);
    expect(serializeOoxmlPart(split)).not.toContain(NS);
    const saved = main(reopened(withPart(pkg, split)));
    expect(storyParagraphs(body(saved)).map((node) => paragraphTextOf(saved, node.id))).toEqual(
      expected
    );
  });

test('multi-split and joining a revived paragraph do not inherit an earlier removed mark', () => {
  const pkg = rejected(),
    part = main(pkg),
    id = body(part).children[1]!.id;
  const peer = apply(part, { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' }, 'bob');
  const split = apply(peer, { op: 'splitParagraphMany', paragraphId: id, offsets: [0, 4] }, 'bob');
  expect(storyParagraphs(body(split)).map((node) => paragraphTextOf(split, node.id))).toEqual([
    'Keep',
    '',
    'Peer',
    '',
  ]);
  expect(serializeOoxmlPart(split)).not.toContain(NS);
  const source = load(p('Tail'));
  const fresh = body(main(source)).children[0]!;
  if (fresh.kind !== 'paragraph') throw new Error('Missing tail seed');
  // Keep the original retired identity while adding an independent following paragraph.
  const tail = apply(
    peer,
    {
      op: 'insertParagraphs',
      parentId: body(peer).id,
      beforeId: null,
      paragraphs: [fresh],
    },
    'bob'
  );
  expect(isRetiredParagraph(findNode(tail, id)!)).toBe(true);
  const ids = storyParagraphs(body(tail)).map((node) => node.id);
  const joined = apply(tail, { op: 'joinParagraphs', firstId: ids[1]!, secondId: ids[2]! }, 'bob');
  expect(serializeOoxmlPart(joined)).not.toContain(NS);
  expect(storyParagraphs(body(joined)).map((node) => paragraphTextOf(joined, node.id))).toEqual([
    'Keep',
    'PeerTail',
  ]);
});

test('a fresh paragraph fragment copies peer content and formatting without copying retirement identity', () => {
  const source = main(rejected()),
    id = body(source).children[1]!.id;
  const peer = apply(source, { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' }, 'bob');
  const before = canonicalOoxmlFingerprint(peer);
  const target = main(load(p('Target'))),
    anchor = body(target).children[0]!.id;
  const pasted = apply(
    target,
    {
      op: 'insertFragment',
      paragraphId: anchor,
      offset: 6,
      blocks: [findNode(peer, id)!],
      lastMarkCovered: true,
    },
    'bob'
  );
  expect(storyParagraphs(body(pasted)).map((node) => paragraphTextOf(pasted, node.id))).toEqual([
    'TargetPeer',
    '',
  ]);
  expect(serializeOoxmlPart(pasted)).not.toContain(NS);
  expect(canonicalOoxmlFingerprint(peer)).toBe(before);
});

test('solo revision decisions retain their prior physical removal behavior', () => {
  const original = main(load(p('Keep') + inserted)),
    id = body(original).children[1]!.id;
  const next = apply(original, { op: 'rejectAllRevisions' }, null);
  expect(findNode(next, id)).toBeNull();
  expect(serializeOoxmlPart(next)).not.toContain(NS);
});

test('explicit structural deletion removes the retired identity rather than reviving it', () => {
  const next = main(rejected()),
    id = body(next).children[1]!.id;
  const removed = apply(next, { op: 'deleteBlock', blockId: id });
  expect(findNode(removed, id)).toBeNull();
  expect(serializeOoxmlPart(removed)).not.toContain(NS);
  expect(storyBlocks(removed).map((node) => paragraphTextOf(removed, node.id))).toEqual(['Keep']);
});

test('accepting a deleted terminal paragraph also retains later independent peer content', () => {
  const deleted = `<w:p><w:pPr><w:rPr><w:del ${rev}/></w:rPr></w:pPr><w:del ${rev}><w:r><w:delText>Old</w:delText></w:r></w:del></w:p>`;
  const source = load(p('Keep') + deleted),
    original = main(source),
    id = body(original).children[1]!.id;
  const accepted = apply(original, { op: 'acceptAllRevisions' });
  expect(isRetiredParagraph(findNode(accepted, id)!)).toBe(true);
  expect(collectRevisionSites(accepted)).toHaveLength(0);
  expect(storyParagraphs(body(accepted))).toHaveLength(1);
  const peer = apply(
    accepted,
    { op: 'insertText', paragraphId: id, offset: 0, text: 'Peer' },
    'bob'
  );
  const saved = main(reopened(withPart(source, peer)));
  expect(storyParagraphs(body(saved)).map((node) => paragraphTextOf(saved, node.id))).toEqual([
    'Keep',
    'Peer',
  ]);
  expect(serializeOoxmlPart(saved)).not.toContain(NS);
});

test('removing the last ordinary sibling leaves one usable required paragraph in layout and Word', () => {
  const next = rejected(),
    part = main(next),
    [kept, retired] = body(part).children;
  const removed = apply(part, { op: 'deleteBlock', blockId: kept!.id });
  expect(storyParagraphs(body(removed)).map((node) => node.id)).toEqual([retired!.id]);
  expect(storyBlocks(removed).map((node) => node.id)).toEqual([retired!.id]);
  const saved = main(reopened(withPart(next, removed)));
  expect(storyParagraphs(body(saved))).toHaveLength(1);
  expect(serializeOoxmlPart(saved)).not.toContain(NS);
});

for (const marker of [
  '<r:retired/>',
  '<r:retired unexpected="1"/>',
  '<r:retired><r:unknown/></r:retired>',
])
  test(`future or malformed marker syntax stays opaque during export: ${marker}`, () => {
    const namespace = marker === '<r:retired/>' ? NS + ':future' : NS;
    const pkg = load(p('Keep') + `<w:p><w:pPr xmlns:r="${namespace}">${marker}</w:pPr></w:p>`);
    const source = serializeOoxmlPart(main(pkg));
    expect(isRetiredParagraph(body(main(pkg)).children[1]!)).toBe(false);
    expect(serializeOoxmlPart(main(reopened(pkg)))).toBe(source);
  });

test('a peer inline marker or unknown element keeps its paragraph visible', () => {
  const marker = `<r:retired xmlns:r="${NS}"/>`;
  for (const inline of [
    '<w:bookmarkStart w:id="4" w:name="Peer"/>',
    '<peer:data xmlns:peer="urn:peer"/>',
  ]) {
    const pkg = load(p('Keep') + `<w:p><w:pPr>${marker}</w:pPr>${inline}</w:p>`),
      part = main(pkg),
      id = body(part).children[1]!.id;
    expect(storyBlocks(part).map((node) => node.id)).toContain(id);
    const saved = main(reopened(pkg));
    expect(storyParagraphs(body(saved))).toHaveLength(2);
    expect(serializeOoxmlPart(saved)).toContain(
      inline.startsWith('<w:bookmark') ? 'bookmarkStart' : 'urn:peer'
    );
    expect(serializeOoxmlPart(saved)).not.toContain(NS);
  }
});

for (const properties of [
  '<w:sectPr/>',
  '<w:pPrChange w:id="22" w:author="Peer"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>',
  '<w:rPr><w:ins w:id="23" w:author="Peer"/></w:rPr>',
])
  test(`a section or pending peer property proposal keeps a retired paragraph visible: ${properties}`, () => {
    const pkg = load(
      p('Keep') + `<w:p><w:pPr><r:retired xmlns:r="${NS}"/>${properties}</w:pPr></w:p>`
    );
    const part = main(pkg),
      id = body(part).children[1]!.id;
    expect(storyParagraphs(body(part)).map((node) => node.id)).toContain(id);
    expect(storyBlocks(part).map((node) => node.id)).toContain(id);
    const saved = main(reopened(pkg));
    expect(storyParagraphs(body(saved))).toHaveLength(2);
    const expected = main(load(p('Keep') + `<w:p><w:pPr>${properties}</w:pPr></w:p>`));
    expect(serializeOoxmlPart(saved)).toBe(serializeOoxmlPart(expected));
    expect(serializeOoxmlPart(saved)).not.toContain(NS);
  });

test('an identical element in opaque paragraph metadata stays unchanged', () => {
  const pkg = load(
    p('Keep') +
      `<w:p><w:pPr><peer:opaque xmlns:peer="urn:peer"><r:retired xmlns:r="${NS}"/></peer:opaque></w:pPr></w:p>`
  );
  const before = serializeOoxmlPart(main(pkg));
  expect(isRetiredParagraph(body(main(pkg)).children[1]!)).toBe(false);
  expect(serializeOoxmlPart(main(reopened(pkg)))).toBe(before);
});

test('a marker in an unrelated custom XML part stays unchanged', () => {
  const loaded = readOoxmlPackage(
    zipDoc({
      body: p('Keep'),
      extraXml: {
        'customXml/private.xml': `<data xmlns="urn:private" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><r:retired xmlns:r="${NS}"/></w:pPr></w:p></data>`,
      },
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const before = loaded.package.parts.get('/customXml/private.xml');
  expect(before).toBeDefined();
  const output = reopened(loaded.package).parts.get('/customXml/private.xml')!;
  expect(serializeOoxmlPart(output)).toBe(serializeOoxmlPart(before!));
});

for (const [name, wrap] of [
  ['control', (xml: string) => `<w:sdt><w:sdtPr/><w:sdtContent>${xml}</w:sdtContent></w:sdt>`],
  ['custom XML', (xml: string) => `<w:customXml>${xml}</w:customXml>`],
  [
    'cell',
    (xml: string) =>
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${xml}</w:tc></w:tr></w:tbl>`,
  ],
] as const)
  test(`retired paragraphs inside a ${name} stay out of story traversal and Word export`, () => {
    const source = load(wrap(p('Keep') + inserted));
    const ids = storyParagraphs(body(main(source))).map((node) => node.id);
    const next = main(rejected(source));
    expect(findNode(next, ids[1]!)).not.toBeNull();
    expect(storyParagraphs(body(next)).map((node) => node.id)).toEqual([ids[0]!]);
    const collected: string[] = [];
    walkAllStoryParagraphs(body(next).children, 0, (node) => collected.push(node.id));
    expect(collected).toEqual([ids[0]!]);
    if (name !== 'cell')
      expect(collectFlowBlocks(body(next).children).map((node) => node.id)).toEqual([ids[0]!]);
    const saved = main(reopened(withPart(source, next)));
    expect(storyParagraphs(body(saved)).map((node) => paragraphTextOf(saved, node.id))).toEqual([
      'Keep',
    ]);
    expect(serializeOoxmlPart(saved)).not.toContain(NS);
  });

for (const actor of [null, 'alice'])
  for (const prefix of [
    '',
    '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Before</w:t></w:r></w:p></w:sdtContent></w:sdt>',
  ])
    for (const action of ['reject', 'accept'] as const)
      test(`resolved empty paragraph before a block control restores exact Word content; actor=${actor}, prefix=${Boolean(prefix)}, action=${action}`, () => {
        const control =
          '<w:sdt><w:sdtPr><w:tag w:val="opaque-control"/></w:sdtPr><w:sdtContent>' +
          p('Control body') +
          '</w:sdtContent></w:sdt>';
        const proposed =
          action === 'reject'
            ? inserted
            : inserted
                .replaceAll('w:ins', 'w:del')
                .replace('<w:t>New</w:t>', '<w:delText>New</w:delText>');
        const source = load(prefix + proposed + control + p('Tail')),
          part = main(source),
          beforeFingerprint = canonicalOoxmlFingerprint(part),
          target = body(part).children[prefix ? 1 : 0]!,
          siblingControl = body(part).children[prefix ? 2 : 1]!,
          tail = body(part).children.at(-1)!,
          next = apply(
            part,
            { op: action === 'reject' ? 'rejectAllRevisions' : 'acceptAllRevisions' },
            actor
          ),
          output = reopened(withPart(source, next));
        expect(serializeOoxmlPart(main(output))).toBe(
          serializeOoxmlPart(main(load(prefix + control + p('Tail'))))
        );
        expect(collectRevisionSites(next)).toEqual([]);
        expect(findNode(next, siblingControl.id)).toBe(siblingControl);
        expect(findNode(next, tail.id)).toBe(tail);
        expect(validateOoxmlPart(next)).toEqual({ ok: true });
        expect(canonicalOoxmlFingerprint(part)).toBe(beforeFingerprint);
        if (actor) {
          expect(isRetiredParagraph(findNode(next, target.id)!)).toBe(true);
          const revived = apply(
            next,
            { op: 'insertText', paragraphId: target.id, offset: 0, text: 'Peer' },
            'bob'
          );
          expect(paragraphTextOf(revived, target.id)).toBe('Peer');
          expect(storyBlocks(revived).some((node) => node.id === target.id)).toBe(true);
          const saved = main(reopened(withPart(source, revived)));
          expect(
            storyParagraphs(body(saved)).map((node) => paragraphTextOf(saved, node.id))
          ).toContain('Peer');
        } else expect(findNode(next, target.id)).toBeNull();
      });

for (const actor of [null, 'alice'])
  for (const sourceBody of [
    inserted + '<w:sdt><w:sdtPr/><w:sdtContent>' + p('Control') + '</w:sdtContent></w:sdt>',
    p('Keep') +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>' +
      p('Cell') +
      '</w:tc></w:tr></w:tbl>' +
      inserted,
  ])
    test(`empty mark retirement preserves a sole direct paragraph and a terminal paragraph after a table; actor=${actor}, afterTable=${sourceBody.startsWith(p('Keep'))}`, () => {
      const source = load(sourceBody),
        part = main(source),
        target = sourceBody.startsWith(p('Keep'))
          ? body(part).children.at(-1)!
          : body(part).children[0]!,
        next = apply(part, { op: 'rejectAllRevisions' }, actor);
      expect(findNode(next, target.id)).not.toBeNull();
      expect(isRetiredParagraph(findNode(next, target.id)!)).toBe(false);
      expect(paragraphTextOf(next, target.id)).toBe('');
      expect(validateOoxmlPart(next)).toEqual({ ok: true });
    });

for (const actor of [null, 'alice'])
  for (const marker of ['section', 'bookmark'] as const)
    test(`a leading removed mark preserves section and opaque inline content; actor=${actor}, marker=${marker}`, () => {
      const paragraph =
        marker === 'section'
          ? inserted.replace('</w:pPr>', '<w:sectPr/></w:pPr>')
          : inserted.replace(
              '</w:p>',
              '<w:bookmarkStart w:id="90" w:name="Keep"/><w:bookmarkEnd w:id="90"/></w:p>'
            );
      const source = load(
          paragraph +
            '<w:sdt><w:sdtPr/><w:sdtContent>' +
            p('Control') +
            '</w:sdtContent></w:sdt>' +
            p('Tail')
        ),
        part = main(source),
        target = body(part).children[0]!,
        next = apply(part, { op: 'rejectAllRevisions' }, actor);
      expect(findNode(next, target.id)).not.toBeNull();
      expect(isRetiredParagraph(findNode(next, target.id)!)).toBe(false);
      expect(paragraphTextOf(next, target.id)).toBe('');
      const output = serializeOoxmlPart(main(reopened(withPart(source, next))));
      expect(output).toContain(marker === 'section' ? '<w:sectPr' : 'w:name="Keep"');
      expect(validateOoxmlPart(next)).toEqual({ ok: true });
    });

for (const actor of [null, 'alice'])
  for (const prefix of [
    '',
    '<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Before</w:t></w:r></w:p></w:sdtContent></w:sdt>',
  ])
    for (const action of ['reject', 'accept'] as const)
      test(`selected paragraph removal retains a foreign pending paragraph property revision; actor=${actor}, prefix=${Boolean(prefix)}, action=${action}`, () => {
        const control =
          '<w:sdt><w:sdtPr/><w:sdtContent>' + p('Control') + '</w:sdtContent></w:sdt>';
        const proposed = (
          action === 'reject'
            ? inserted
            : inserted
                .replaceAll('w:ins', 'w:del')
                .replace('<w:t>New</w:t>', '<w:delText>New</w:delText>')
        )
          .replace('<w:pPr>', '<w:pPr><w:jc w:val="right"/>')
          .replace(
            '</w:pPr>',
            '<w:pPrChange w:id="22" w:author="Peer"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>'
          );
        const source = load(prefix + proposed + control + p('Tail')),
          part = main(source),
          target = body(part).children[prefix ? 1 : 0]!,
          peer = collectRevisionSites(part).find((site) => site.node.localName === 'pPrChange')!,
          owned = collectRevisionSites(part).filter((site) =>
            site.node.attributes.some(
              (attribute) => attribute.localName === 'author' && attribute.value === 'AI'
            )
          ),
          props =
            target.kind !== 'textValue' &&
            target.children.find((node) => node.kind !== 'textValue' && node.localName === 'pPr'),
          alignment =
            props &&
            props.kind !== 'textValue' &&
            props.children.find((node) => node.kind !== 'textValue' && node.localName === 'jc'),
          unchangedControl = body(part).children[prefix ? 2 : 1]!,
          tail = body(part).children.at(-1)!;
        expect(owned).toHaveLength(2);
        expect(alignment).toBeTruthy();
        const next = apply(
          part,
          {
            op: action === 'reject' ? 'rejectAllRevisions' : 'acceptAllRevisions',
            siteNodeIds: owned.map((site) => site.node.id),
          },
          actor
        );
        expect(findNode(next, peer.node.id)).toBe(peer.node);
        expect(alignment && findNode(next, alignment.id)).toBe(alignment);
        expect(findNode(next, target.id)).not.toBeNull();
        expect(paragraphTextOf(next, target.id)).toBe('');
        expect(collectRevisionSites(next).map((site) => site.node.id)).toEqual([peer.node.id]);
        expect(findNode(next, unchangedControl.id)).toBe(unchangedControl);
        expect(findNode(next, tail.id)).toBe(tail);
        expect(validateOoxmlPart(next)).toEqual({ ok: true });
        const output = main(reopened(withPart(source, next))),
          revisions = collectRevisionSites(output);
        expect(revisions).toHaveLength(1);
        expect(revisions[0]!.node.localName).toBe('pPrChange');
        expect(
          revisions[0]!.node.attributes.find((attribute) => attribute.localName === 'author')?.value
        ).toBe('Peer');
      });
