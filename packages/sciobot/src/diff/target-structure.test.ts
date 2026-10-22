import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import { protectedDocument } from '../../test/support/protected-document';
import { targetParagraph } from '../../test/support/target-document';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import { parseTargetMarkdown } from '../projection/parse-target';
import { planTargetDocument } from './plan-target';

const fixtures = [
  {
    name: 'section insertion',
    before: ['Opening.', 'Closing.'],
    after: ['Opening.', 'New section', 'First addition.', 'Second addition.', 'Closing.'],
  },
  { name: 'deletion', before: ['Keep.', 'Remove.', 'Tail.'], after: ['Keep.', 'Tail.'] },
  {
    name: 'complete rewrite',
    before: ['Old document.', 'Old end.'],
    after: ['Replacement', 'New body.', 'New ending.'],
  },
] as const;
function ast(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst;
}

for (const fixture of fixtures)
  test(`C01 ${fixture.name} applies native plans and preserves imported non-body content through reopen`, async () => {
    const entries = unzipSync(protectedDocument()),
      document = strFromU8(entries['word/document.xml']!),
      bodyStart = document.indexOf('<w:body>') + '<w:body>'.length,
      sectionStart = document.lastIndexOf('<w:sectPr>');
    entries['word/document.xml'] = strToU8(
      document.slice(0, bodyStart) +
        fixture.before.map(targetParagraph).join('') +
        document.slice(sectionStart)
    );
    const loaded = readOoxmlPackage(zipSync(entries));
    if (!loaded.ok) throw new Error(loaded.reason);
    const pkg = loaded.package,
      main = pkg.parts.get(pkg.mainDocumentPart)!,
      body = bodyStoryRoot(main)!;
    if (body.kind === 'textValue') throw new Error('Invalid fixture body');
    const section = body.children.at(-1)!,
      background = main.root.children.find(
        (node) => node.kind !== 'textValue' && node.localName === 'background'
      )!,
      originalParagraphs = storyParagraphs(body),
      bytes = writeOoxmlPackage(pkg),
      target = pairs.cases.find((item) => item.name === fixture.name)!.afterMarkdown,
      result = planTargetDocument(
        { pkg, revision: sourceRevision, projection: await projectSourcePackage(pkg) },
        target
      );
    if (!result.success) throw new Error(result.error);
    const applied = applyTargetPlanOracle(pkg, result.plan),
      nextMain = applied.parts.get(pkg.mainDocumentPart)!,
      nextBody = bodyStoryRoot(nextMain)!;
    expect(storyParagraphs(nextBody).map((node) => paragraphTextOf(nextMain, node.id))).toEqual([
      ...fixture.after,
    ]);
    expect(findNode(nextMain, section.id)).toBe(section);
    expect(findNode(nextMain, background.id)).toBe(background);
    for (const pair of result.pairs) {
      const original = originalParagraphs.find((node) => node.id === pair.id)!,
        retained = findNode(nextMain, pair.id)!;
      expect(retained.kind).toBe('paragraph');
      if (retained.kind !== 'paragraph' || original.kind !== 'paragraph')
        throw new Error('Missing retained paragraph');
      expect(retained.children[0]).toBe(original.children[0]);
    }
    for (const [name, part] of pkg.parts)
      if (name !== pkg.mainDocumentPart) expect(applied.parts.get(name)).toBe(part);
    const saved = writeOoxmlPackage(applied),
      beforeEntries = unzipSync(bytes),
      afterEntries = unzipSync(saved);
    expect(Object.keys(afterEntries).sort()).toEqual(Object.keys(beforeEntries).sort());
    for (const [name, content] of Object.entries(beforeEntries))
      if (name !== 'word/document.xml') expect(afterEntries[name]).toEqual(content);
    const reopened = readOoxmlPackage(saved);
    if (!reopened.ok) throw new Error(reopened.reason);
    const reopenedMain = reopened.package.parts.get(pkg.mainDocumentPart)!,
      reopenedBody = bodyStoryRoot(reopenedMain)!;
    if (reopenedBody.kind === 'textValue') throw new Error('Invalid reopened body');
    expect(canonicalOoxmlFingerprint(reopenedBody.children.at(-1)!)).toBe(
      canonicalOoxmlFingerprint(section)
    );
    expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(target));
    expect(writeOoxmlPackage(pkg)).toEqual(bytes);
  });
