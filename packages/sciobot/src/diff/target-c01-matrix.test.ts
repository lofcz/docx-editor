import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import { protectedTargetInput } from '../../test/support/target-document';
import { C01_SOURCES } from '../../test/support/c01-sources';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { parseTargetMarkdown } from '../projection/parse-target';
import { planTargetDocument } from './plan-target';
import type { CanonicalEdit } from './plan';

const expectedKinds: Record<string, readonly CanonicalEdit['kind'][]> = {
  'formatting and lists': ['text', 'siblings'],
  'tables and links': ['text', 'insert-row', 'format'],
  'section insertion': ['siblings'],
  'nested MDX and attributes': ['payload', 'text', 'text'],
  'source content': ['payload', 'text', 'text', 'text', 'text'],
  deletion: ['siblings'],
  'complete rewrite': ['text', 'text', 'text', 'text', 'text', 'siblings'],
};
/** Parts each captured pair may change. Every other imported part must keep its exact bytes. */
const changedParts: Record<string, readonly string[]> = {
  'tables and links': ['word/document.xml', 'word/_rels/document.xml.rels'],
  'nested MDX and attributes': ['word/document.xml', 'customXml/item1.xml'],
  'source content': ['word/document.xml', 'customXml/item1.xml'],
};
const protectedParts = [
  'word/header1.xml',
  'word/footer1.xml',
  'word/footnotes.xml',
  'word/media/image.png',
  'custom/opaque.xml',
  'custom/opaque.bin',
];
function ast(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  // Word paragraphs carry no loose/tight list distinction; native projection always emits loose lists.
  return JSON.parse(
    JSON.stringify(result.target.normalizedAst, (key, value) =>
      key === 'spread' ? undefined : value
    )
  ) as unknown;
}

test('the matrix covers exactly the seven captured C01 pairs', () => {
  expect(pairs.cases.map((item) => item.name).sort()).toEqual(Object.keys(C01_SOURCES).sort());
  expect(Object.keys(expectedKinds).sort()).toEqual(Object.keys(C01_SOURCES).sort());
});

for (const fixture of pairs.cases)
  test(`C01 ${fixture.name} plans explicit operations and retains imported non-body content`, async () => {
    const [body, options] = C01_SOURCES[fixture.name]!(),
      input = await protectedTargetInput(body, options),
      bytes = writeOoxmlPackage(input.pkg),
      main = input.pkg.parts.get(input.pkg.mainDocumentPart)!,
      root = bodyStoryRoot(main)!;
    if (root.kind === 'textValue') throw new Error('Invalid fixture body');
    const section = root.children.at(-1)!,
      background = main.root.children.find(
        (node) => node.kind !== 'textValue' && node.localName === 'background'
      )!;
    // The native source is the captured before document: planning it back yields no operations.
    const identity = planTargetDocument(input, fixture.beforeMarkdown);
    if (!identity.success) throw new Error(identity.error);
    expect(identity.plan.operations).toEqual([]);

    const result = planTargetDocument(input, fixture.afterMarkdown);
    if (!result.success) throw new Error(result.error);
    expect(result.plan.operations.map((edit) => edit.kind)).toEqual([
      ...expectedKinds[fixture.name]!,
    ]);
    expect(result.plan.revision).toEqual(input.revision);
    expect(JSON.stringify(result.plan)).not.toContain('replaceStoryBlocks');
    expect(result.plan.operations.every((edit) => edit.partName === main.name)).toBe(true);

    const applied = applyTargetPlanOracle(input.pkg, result.plan),
      appliedMain = applied.parts.get(applied.mainDocumentPart)!;
    expect(findNode(appliedMain, section.id)).toBe(section);
    expect(findNode(appliedMain, background.id)).toBe(background);
    for (const pair of result.pairs) expect(findNode(appliedMain, pair.id)).not.toBeNull();

    const saved = writeOoxmlPackage(applied),
      before = unzipSync(bytes),
      after = unzipSync(saved),
      allowed = new Set(changedParts[fixture.name] ?? ['word/document.xml']);
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    for (const name of protectedParts) expect(after[name]).toEqual(before[name]!);
    for (const [name, content] of Object.entries(before))
      if (!allowed.has(name)) expect({ name, content: after[name] }).toEqual({ name, content });
    if (allowed.has('word/_rels/document.xml.rels'))
      expect(strFromU8(after['word/_rels/document.xml.rels']!)).toContain('https://example.org');

    const reopened = readOoxmlPackage(saved);
    if (!reopened.ok) throw new Error(reopened.reason);
    const reopenedRoot = bodyStoryRoot(reopened.package.parts.get(input.pkg.mainDocumentPart)!)!;
    if (reopenedRoot.kind === 'textValue') throw new Error('Invalid reopened body');
    expect(canonicalOoxmlFingerprint(reopenedRoot.children.at(-1)!)).toBe(
      canonicalOoxmlFingerprint(section)
    );
    const projection = await projectSourcePackage(reopened.package),
      landed = planTargetDocument(
        { revision: input.revision, pkg: reopened.package, projection },
        fixture.afterMarkdown
      );
    if (!landed.success) throw new Error(landed.error);
    expect(landed.plan.operations).toEqual([]);
    // Registered MDX projections serialize attributes in payload order; compare native Markdown exactly.
    if (!options.payloads) expect(ast(projection.text)).toEqual(ast(fixture.afterMarkdown));
    expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  }, 60_000);
