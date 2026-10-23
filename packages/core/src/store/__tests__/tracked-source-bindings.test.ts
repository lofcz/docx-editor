import { expect, test } from 'bun:test';
import { isPackageEditRefusal } from '../store/package-edit-refusal.ts';
import {
  canonicalOoxmlFingerprint,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  findNode,
  replaceNode,
  serializeOoxmlPart,
  readOoxmlPart,
  type OoxmlElement,
} from '../index.ts';
import {
  TRACKED_SOURCE_BINDINGS_NS,
  withTrackedSourceBindings,
  readTrackedSourceBindings,
  materializeTrackedSourceBindings,
  type TrackedSourceBinding,
} from '../package/tracked-source-bindings.ts';
import { part, paragraphId, apply, ADA } from './tracked-edit-fixture.ts';
const scope = '12345678-1234-4234-8234-123456789abc';
const origin = scope + '/12345678-1234-8234-a234-123456789abc';
const original = () =>
  part(
    '<w:p><w:r w:rsidR="ABC"><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> Peer</w:t></w:r><w:r><w:t>Independent</w:t></w:r></w:p>'
  );
function bound() {
  const source = original(),
    p = findNode(source, paragraphId(source))!;
  if (p.kind !== 'paragraph') throw Error('Missing paragraph');
  const run = p.children.find((node) => node.kind === 'run')!;
  if (run.kind !== 'run') throw Error('Missing run');
  const text = run.children.find((node) => node.kind === 'text')!;
  const record: TrackedSourceBinding = {
    scope,
    runId: run.id,
    token: origin,
    absent: false,
    retained: [],
    texts: [{ id: text.id, token: origin, space: 'absent' }],
  };
  return { source, p, run, text, record, marked: withTrackedSourceBindings(p, [record]) };
}
test('source bindings transfer exact run/text identities, preserve independent content, and consume only selected scopes', () => {
  const { source, p, run, text, marked } = bound();
  const before = canonicalOoxmlFingerprint(marked);
  expect(materializeTrackedSourceBindings(marked, new Set(['other']))).toBe(marked);
  const projected = materializeTrackedSourceBindings(marked, new Set([scope]));
  if (projected.kind !== 'paragraph') throw Error('Missing paragraph');
  expect(readTrackedSourceBindings(projected)).toEqual([]);
  expect(canonicalOoxmlFingerprint(marked)).toBe(before);
  expect(projected.children[1]).toBe(p.children[1]);
  const native = projected.children.find((node) => node.id === run.id)!;
  if (native.kind !== 'run') throw Error('Missing run');
  expect(native.attributes.find((attr) => attr.localName === 'origin')?.value).toBe(origin);
  expect(native.attributes.find((attr) => attr.localName === 'rsidR')?.value).toBe('ABC');
  const nativeText = native.children.find((node) => node.id === text.id)!;
  if (nativeText.kind !== 'text') throw Error('Missing text');
  expect(nativeText.attributes.find((attr) => attr.localName === 'textOrigin')?.value).toBe(origin);
  expect(nativeText.children).toBe(text.kind === 'text' ? text.children : null);
  expect(source.root).not.toBe(projected);
});
test('synchronous XML export transfers proofs and removes load-local IDs without mutating the source', () => {
  const { source, p, marked } = bound();
  const changed = replaceNode(source, p.id, marked);
  if (!changed.ok) throw Error('Missing source');
  const before = canonicalOoxmlFingerprint(changed.part);
  const xml = serializeOoxmlPart(changed.part);
  expect(xml).not.toContain(TRACKED_SOURCE_BINDINGS_NS);
  expect(xml).toContain('tracked-run-origins:1/');
  expect(xml).not.toContain('/word/document.xml#');
  expect(canonicalOoxmlFingerprint(changed.part)).toBe(before);
  const reopened = readOoxmlPart(xml, { name: source.name, contentType: source.contentType });
  expect(reopened.ok).toBe(true);
  expect(materializeTrackedSourceBindings(changed.part.root)).toBe(
    materializeTrackedSourceBindings(changed.part.root)
  );
});
for (const malformed of [
  '{',
  JSON.stringify({ schemaVersion: 2, sources: [] }),
  JSON.stringify({ schemaVersion: 1, sources: [{ scope: '__proto__' }] }),
])
  test(`unknown or malformed source bindings remain opaque during export and refuse selected interpretation: ${malformed}`, () => {
    const { p, marked } = bound();
    const invalid = {
      ...marked,
      attributes: marked.attributes.map((attr) =>
        attr.namespaceUri === TRACKED_SOURCE_BINDINGS_NS ? { ...attr, value: malformed } : attr
      ),
    } as OoxmlElement;
    expect(() => readTrackedSourceBindings(invalid)).toThrow();
    expect(materializeTrackedSourceBindings(invalid)).toBe(invalid);
    try {
      materializeTrackedSourceBindings(invalid, new Set([scope]));
      throw Error('Selected malformed bindings must refuse');
    } catch (error) {
      expect(isPackageEditRefusal(error)).toBe(true);
      if (isPackageEditRefusal(error)) expect(error.reason).toBe('unsupported');
    }
    expect(p.attributes.some((attr) => attr.namespaceUri === TRACKED_SOURCE_BINDINGS_NS)).toBe(
      false
    );
  });
test('binding quotas, duplicate identities, and invalid retained prefixes refuse before writing', () => {
  const { p, record } = bound();
  for (const records of [
    [record, record],
    [{ ...record, runId: 'x'.repeat(4097) }],
    [{ ...record, retained: ['not a prefix'] }],
    [{ ...record, retained: ['same', 'same'] }],
    [{ ...record, texts: [{ ...record.texts[0]!, id: record.runId }] }],
    Array.from({ length: 1025 }, () => record),
  ])
    expect(() => withTrackedSourceBindings(p, records)).toThrow();
  expect(readTrackedSourceBindings(p)).toEqual([]);
});

for (const decision of ['accept', 'reject'] as const)
  test(`source bindings preserve inherited opaque prefix meanings through ${decision}`, () => {
    const source = part(
      '<w:p xmlns:sourceBindings="urn:keep"><w:r sourceBindings:value="sourceBindings:keep"><w:t>AB</w:t></w:r></w:p>'
    );
    const id = paragraphId(source);
    const prepared = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
    const inserted = apply(prepared.part, {
      op: 'insertText',
      paragraphId: id,
      offset: 1,
      text: 'X',
      revision: ADA,
    });
    const pending = sealTrackedRunOrigins(inserted, id, prepared.scope!, source);
    const decided = apply(pending, {
      op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
    });
    const expected =
      decision === 'accept'
        ? apply(source, { op: 'insertText', paragraphId: id, offset: 1, text: 'X' })
        : source;
    expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(expected));
    const xml = serializeOoxmlPart(decided);
    expect(xml).toContain('sourceBindings:keep');
    expect(xml).toContain('urn:keep');
  });

test('rich source bindings use bounded chunks and never split a Unicode scalar value', () => {
  const { p, record } = bound();
  const records = Array.from({ length: 200 }, (_, index) => ({
    ...record,
    runId: 'run-' + index + '😀'.repeat(20),
    texts: [{ ...record.texts[0]!, id: 'text-' + index }],
  }));
  const marked = withTrackedSourceBindings(p, records);
  const attrs = marked.attributes.filter(
    (attr) => attr.namespaceUri === TRACKED_SOURCE_BINDINGS_NS
  );
  expect(attrs.length).toBeGreaterThan(1);
  expect(
    attrs.every(
      (attr) =>
        attr.value.length <= 2048 &&
        !/[\uD800-\uDBFF]$/.test(attr.value) &&
        !/^[\uDC00-\uDFFF]/.test(attr.value)
    )
  ).toBe(true);
  expect(readTrackedSourceBindings(marked)).toEqual(records);
  const missing = {
    ...marked,
    attributes: marked.attributes.filter((attr) => attr !== attrs[1]),
  } as OoxmlElement;
  expect(() => readTrackedSourceBindings(missing)).toThrow();
  expect(materializeTrackedSourceBindings(missing)).toBe(missing);
});
