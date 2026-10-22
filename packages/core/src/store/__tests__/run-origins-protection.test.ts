import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlsIn,
  findNode,
  prepareTrackedRunOrigins,
  readOoxmlPart,
  sealTrackedRunOrigins,
  serializeOoxmlPart,
  storyParagraphs,
  type OoxmlPart,
} from '../index.ts';
import { ADA, apply, paragraphId, part } from './tracked-edit-fixture.ts';
import { RUN_ORIGIN_NS, hasOriginMetadata } from '../store/run-origin-metadata.ts';

const run = (text: string) => `<w:r><w:rPr/><w:t>${text}</w:t></w:r>`;
const fingerprint = canonicalOoxmlFingerprint;
const control = (properties: string) =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${run('LOCKED')}</w:sdtContent></w:sdt>`;

for (const properties of [
  '<w:lock w:val="contentLocked"/>',
  '<w:lock w:val="sdtContentLocked"/>',
  '<w:dataBinding w:storeItemID="{11111111-1111-1111-1111-111111111111}" w:xpath="/data/value"/>',
]) {
  test(`run-origin preparation preserves unrelated inline controls: ${properties}`, () => {
    const source = part(`<w:p>${run('ab')}${control(properties)}${run('cd')}</w:p>`);
    const id = paragraphId(source);
    const held = contentControlsIn(source.root)[0]!.node;
    const node = findNode(source, held.id)!;
    const before = fingerprint(source);
    for (const offset of [1, 2]) {
      const marked = prepareTrackedRunOrigins(source, id, [{ start: offset, end: offset }]);
      expect(marked.scope).not.toBeNull();
      expect(findNode(marked.part, held.id)).toBe(node);
      expect(fingerprint(findNode(marked.part, held.id)!)).toBe(fingerprint(node));
      const p = findNode(marked.part, id)!;
      if (p.kind !== 'paragraph') throw Error('paragraph');
      expect(hasOriginMetadata(p.children.at(-1)!)).toBe(false);
    }
    expect(() => prepareTrackedRunOrigins(source, id, [{ start: 3, end: 3 }])).toThrow(
      'protected or opaque'
    );
    expect(() =>
      prepareTrackedRunOrigins(source, id, [
        { start: 2, end: 2 },
        { start: 4, end: 5 },
      ])
    ).toThrow('protected or opaque');
    expect(fingerprint(source)).toBe(before);
  });

  test(`run-origin preparation refuses a protected paragraph ancestor: ${properties}`, () => {
    const source = part(
      `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent><w:p>${run('AB')}</w:p></w:sdtContent></w:sdt>`
    );
    const paragraph = storyParagraphs(bodyStoryRoot(source)!)[0]!;
    expect(() => prepareTrackedRunOrigins(source, paragraph.id, [{ start: 1, end: 1 }])).toThrow(
      'protected'
    );
  });
}

test('run-origin preparation supports default WML namespaces without rebinding existing prefixes', () => {
  const parsed = readOoxmlPart(
    '<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:runOrigin="urn:keep"><body><p><r><t>AB</t></r></p></body></document>',
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!parsed.ok) throw Error(parsed.reason);
  const marked = prepareTrackedRunOrigins(parsed.part, paragraphId(parsed.part), [
    { start: 1, end: 1 },
  ]);
  expect(marked.scope).not.toBeNull();
  const xml = serializeOoxmlPart(marked.part);
  expect(xml).toContain('xmlns:runOrigin="urn:keep"');
  expect(xml).toContain(RUN_ORIGIN_NS);
});

test('round-trip decisions preserve source prefix collisions and unused source namespace bindings', () => {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:runOrigin="urn:keep" xmlns:reserved="${RUN_ORIGIN_NS}"><w:body><w:p><w:r><w:rPr><runOrigin:opaque runOrigin:value="keep"/></w:rPr><w:t>AB</w:t></w:r></w:p></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!parsed.ok) throw Error(parsed.reason);
  const source = parsed.part;
  const marked = prepareTrackedRunOrigins(source, paragraphId(source), [{ start: 1, end: 1 }]);
  const pending = sealTrackedRunOrigins(
    apply(marked.part, {
      op: 'insertText',
      paragraphId: paragraphId(source),
      offset: 1,
      text: 'X',
      revision: ADA,
    }),
    paragraphId(source),
    marked.scope!,
    source
  );
  const opened = readOoxmlPart(serializeOoxmlPart(pending), {
    name: source.name,
    contentType: source.contentType,
  });
  if (!opened.ok) throw Error(opened.reason);
  for (const action of ['accept', 'reject'] as const) {
    const decided = apply(opened.part, {
      op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
    });
    const golden =
      action === 'accept'
        ? apply(source, {
            op: 'insertText',
            paragraphId: paragraphId(source),
            offset: 1,
            text: 'X',
          })
        : source;
    expect(fingerprint(decided)).toBe(fingerprint(golden));
    expect(serializeOoxmlPart(decided)).toContain(`xmlns:reserved="${RUN_ORIGIN_NS}"`);
    expect(serializeOoxmlPart(decided)).toContain('urn:keep');
  }
});

test('invalid ranges and excessive provenance work refuse before changing the source', () => {
  const source = part(`<w:p>${run('AB')}</w:p>`);
  const before = fingerprint(source);
  for (const range of [
    { start: -1, end: 0 },
    { start: 2, end: 1 },
    { start: 0, end: 3 },
    { start: NaN, end: 1 },
  ])
    expect(() => prepareTrackedRunOrigins(source, paragraphId(source), [range])).toThrow(
      'Invalid run-origin range'
    );
  expect(() =>
    prepareTrackedRunOrigins(
      source,
      paragraphId(source),
      Array.from({ length: 4097 }, () => ({ start: 0, end: 0 }))
    )
  ).toThrow('exceed');
  expect(fingerprint(source)).toBe(before);
  const large: OoxmlPart = part(
    `<w:p>${run('AB')}${'<w:proofErr w:type="spellStart"/>'.repeat(100_001)}</w:p>`
  );
  expect(() => prepareTrackedRunOrigins(large, paragraphId(large), [{ start: 1, end: 1 }])).toThrow(
    'exceeds its limit'
  );
});
