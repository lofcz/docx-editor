import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  prepareTrackedRunOrigins,
  readOoxmlPart,
  sealTrackedRunOrigins,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { replaceNode, createNodeIdAllocator, insertChildren } from '../package/ooxml-edit.ts';
import { cloneWithNewIds } from '../store/tree-op-nodes.ts';
import { RUN_ORIGIN_NS, readRunOrigin } from '../store/run-origin-metadata.ts';
import { ADA, apply, paragraphId, part } from './tracked-edit-fixture.ts';

const run = (text: string, properties = '<w:rPr/>') => `<w:r>${properties}<w:t>${text}</w:t></w:r>`;
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const fingerprint = canonicalOoxmlFingerprint;
function reopen(source: OoxmlPart): OoxmlPart {
  const read = readOoxmlPart(serializeOoxmlPart(source), {
    name: source.name,
    contentType: source.contentType,
  });
  if (!read.ok) throw Error(read.reason);
  return read.part;
}
function proposed(source: OoxmlPart, id = paragraphId(source)) {
  const prepared = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
  expect(prepared.scope).not.toBeNull();
  const inserted = apply(prepared.part, {
    op: 'insertText',
    paragraphId: id,
    offset: 1,
    text: 'X',
    revision: ADA,
  });
  return sealTrackedRunOrigins(inserted, id, prepared.scope!, source);
}

for (const count of [1, 2])
  for (const properties of ['', '<w:rPr/>'])
    for (const action of ['accept', 'reject'] as const)
      test(`run provenance preserves ${count} source runs and ${properties || 'absent'} properties: ${action}`, () => {
        const source = part(
          `<w:p>${count === 1 ? run('AB', properties) : run('A', properties) + run('B', properties)}</w:p>`
        );
        const before = fingerprint(source);
        const pending = proposed(source);
        const golden =
          action === 'accept'
            ? apply(source, {
                op: 'insertText',
                paragraphId: paragraphId(source),
                offset: 1,
                text: 'X',
              })
            : source;
        for (const current of [pending, reopen(pending)]) {
          const decided = apply(current, {
            op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
          });
          expect(fingerprint(decided)).toBe(fingerprint(golden));
          expect(serializeOoxmlPart(decided)).not.toContain(RUN_ORIGIN_NS);
        }
        expect(fingerprint(source)).toBe(before);
      });

for (const action of ['accept', 'reject'] as const)
  test(`resolving one scope preserves an unselected externally decided scope: ${action}`, () => {
    let current = part(`<w:p>${run('AB')}</w:p><w:p>${run('CD')}</w:p>`);
    const first = paragraphId(current);
    const body = current.root.children.find((node) => node.kind === 'body')!;
    const second = kids(body)[1]!.id;
    current = proposed(current, second);
    // An external Word decision removes revision markup without understanding our metadata.
    for (const site of collectRevisionSites(current)) {
      const node = {
        ...site.node,
        attributes: site.node.attributes.filter((attr) => attr.namespaceUri !== RUN_ORIGIN_NS),
      } as OoxmlNode;
      const changed = replaceNode(current, site.node.id, node);
      if (!changed.ok) throw Error('External revision metadata removal failed.');
      current = changed.part;
    }
    current = apply(current, { op: 'acceptAllRevisions' });
    const unselected = fingerprint(findNode(current, second)!);
    const pending = proposed(current, first);
    const decided = apply(pending, {
      op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
    });
    expect(fingerprint(findNode(decided, second)!)).toBe(unselected);
    expect(fingerprint(findNode(decided, first)!)).toBe(
      fingerprint(findNode(part(`<w:p>${run(action === 'accept' ? 'AXB' : 'AB')}</w:p>`), first)!)
    );
  });

test('published provenance stays outside native formatting-history snapshots', () => {
  const source = part(`<w:p>${run('AB')}</w:p>`);
  const pending = proposed(source);
  const id = paragraphId(pending);
  const runNode = kids(findNode(pending, id)!).find((node) => node.kind === 'run')!;
  expect(readRunOrigin(runNode)).not.toBeNull();
  const formatted = apply(pending, {
    op: 'setRunProperties',
    paragraphId: id,
    start: 0,
    end: 1,
    properties: [{ localName: 'b' }],
    revision: { author: 'Other', date: '2026-10-22T12:00:00Z' },
  });
  const history = collectRevisionSites(formatted).filter(
    (site) => site.node.localName === 'rPrChange'
  );
  expect(history.length).toBeGreaterThan(0);
  for (const site of history) expect(fingerprint(site.node)).not.toContain(RUN_ORIGIN_NS);
});

test('sealing without new native revision sites refuses instead of publishing orphan provenance', () => {
  const source = part(`<w:p>${run('AB')}</w:p>`);
  const prepared = prepareTrackedRunOrigins(source, paragraphId(source), [{ start: 1, end: 1 }]);
  expect(() =>
    sealTrackedRunOrigins(prepared.part, paragraphId(source), prepared.scope!, source)
  ).toThrow('needs new native revision sites');
});

test('copied prior revision addresses never acquire a new run-origin scope', () => {
  const source = part(
    `<w:p><w:ins w:id="42" w:author="${ADA.author}" w:date="${ADA.date}">${run('F')}</w:ins>${run('AB')}</w:p>`
  );
  const id = paragraphId(source);
  const prior = collectRevisionSites(source)[0]!.node;
  const marked = prepareTrackedRunOrigins(source, id, [{ start: 2, end: 2 }]);
  let current = apply(marked.part, {
    op: 'insertText',
    paragraphId: id,
    offset: 2,
    text: 'X',
    revision: ADA,
  });
  const copied = cloneWithNewIds(prior, createNodeIdAllocator(current));
  const p = findNode(current, id)!;
  if (p.kind !== 'paragraph') throw Error('paragraph');
  const added = insertChildren(current, id, p.children.length, [copied]);
  if (!added.ok) throw Error('Copied source insertion failed.');
  current = sealTrackedRunOrigins(added.part, id, marked.scope!, source);
  expect(fingerprint(findNode(current, prior.id)!)).toBe(fingerprint(prior));
  expect(fingerprint(findNode(current, copied.id)!)).toBe(fingerprint(prior));
  const newSites = collectRevisionSites(current).filter(
    (site) => site.node.id !== prior.id && site.node.id !== copied.id
  );
  expect(newSites.length).toBeGreaterThan(0);
  expect(
    newSites.every((site) =>
      site.node.attributes.some(
        (attr) => attr.namespaceUri === RUN_ORIGIN_NS && attr.localName === 'scope'
      )
    )
  ).toBe(true);
});
