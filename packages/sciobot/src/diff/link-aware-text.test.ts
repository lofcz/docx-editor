import { prepareTrackedInlinePlan } from '../review/tracked-inline';
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPart,
  applyTreeOp,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { parseTargetMarkdown } from '../projection/parse-target';
import { compileInlineTarget } from './inline-target';
import { planParagraphText } from './paragraph-text';
import { linkAwareBias } from './link-aware-text';
import { previewTextEdits } from './insertion-marks';
function nodes(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(nodes))];
}
const revision = { documentId: 'link-text', epoch: 'native', revision: 1 } as const;
for (const [label, outside, added, offset, target] of [
  ['label', ' End.', ' ', 11, 'Keep. [label ](https://outer.test) End.'],
  ['aa', 'a End.', 'a', 8, 'Keep. [aaa](https://outer.test)a End.'],
  ['😀', '😀 End.', '😀', 8, 'Keep. [😀😀](https://outer.test)😀 End.'],
  ['label', ' End.', ' ', 6, 'Keep. [ label](https://outer.test) End.'],
] as const)
  test(`link-aware text diff preserves exact native source ownership: ${target}`, () => {
    const loaded = readOoxmlPart(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Keep. </w:t></w:r><w:hyperlink r:id="old" w:history="1"><w:r><w:t>${label}</w:t></w:r></w:hyperlink><w:r><w:t>${outside}</w:t></w:r></w:p></w:body></w:document>`,
      { name: '/word/document.xml', contentType: 'application/xml' }
    );
    if (!loaded.ok) throw Error(loaded.reason);
    const part = loaded.part,
      p = nodes(part.root).find((n) => n.kind === 'paragraph')!,
      link = nodes(part.root).find((n) => n.kind === 'hyperlink')!;
    const parsed = parseTargetMarkdown(target);
    if (!parsed.success) throw Error(parsed.error);
    const ast = (
      parsed.target.normalizedAst.children as import('../projection/parse-target').TargetAstNode[]
    )[0]!;
    const compiled = compileInlineTarget(ast);
    if (!compiled.success || compiled.target.plainText === null) throw Error('Invalid target');
    const planned = planParagraphText(
      { revision, partName: part.name, paragraphId: p.id, text: paragraphTextOf(part, p.id)! },
      compiled.target.plainText
    );
    if (!planned.success) throw Error(planned.error);
    const edits = linkAwareBias(part, p.id, 0, planned.plan.operations, compiled.target);
    const tracked = prepareTrackedInlinePlan(
      part,
      { ...planned.plan, operations: edits },
      { author: 'Alice' },
      '2026-10-24T09:00:00.000Z'
    );
    if (offset === 6) expect(tracked.success).toBe(true);
    else expect(tracked).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(edits).toHaveLength(1);
    expect(edits[0]).toMatchObject({
      kind: 'text',
      start: offset,
      end: offset,
      before: '',
      after: added,
      ...(offset === 6 ? {} : { insideHyperlink: link.id }),
    });
    expect(edits[0]?.kind === 'text' ? edits[0].insideHyperlink : null).toBe(
      offset === 6 ? undefined : link.id
    );
    const preview = previewTextEdits(part, edits);
    if (!preview.success) throw Error(preview.error);
    const expected = applyTreeOp(part, {
      op: 'insertText',
      paragraphId: p.id,
      offset,
      text: added,
      inside: link.id,
    });
    if (!expected.ok) throw Error(expected.reason);
    expect(canonicalOoxmlFingerprint(preview.part)).toBe(canonicalOoxmlFingerprint(expected.part));
    const originalOutside = nodes(part.root).find(
      (n) => n.kind === 'textValue' && n.value === outside
    )!;
    expect(canonicalOoxmlFingerprint(findNode(preview.part, originalOutside.id)!)).toBe(
      canonicalOoxmlFingerprint(originalOutside)
    );
  });

for (const side of ['left', 'right'] as const)
  test(`adjacent native links keep the explicit ${side} insertion owner`, () => {
    const loaded = readOoxmlPart(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Keep. </w:t></w:r><w:hyperlink r:id="left" w:history="1"><w:r><w:t>a</w:t></w:r></w:hyperlink><w:hyperlink r:id="right" w:tgtFrame="_blank"><w:r><w:t>b</w:t></w:r></w:hyperlink><w:r><w:t> End.</w:t></w:r></w:p></w:body></w:document>',
      { name: '/word/document.xml', contentType: 'application/xml' }
    );
    if (!loaded.ok) throw Error(loaded.reason);
    const part = loaded.part,
      p = nodes(part.root).find((n) => n.kind === 'paragraph')!,
      links = nodes(part.root).filter((n) => n.kind === 'hyperlink');
    const target =
      side === 'left'
        ? 'Keep. [aX](https://left.test)[b](https://right.test) End.'
        : 'Keep. [a](https://left.test)[Xb](https://right.test) End.';
    const parsed = parseTargetMarkdown(target);
    if (!parsed.success) throw Error(parsed.error);
    const compiled = compileInlineTarget(
      (
        parsed.target.normalizedAst.children as import('../projection/parse-target').TargetAstNode[]
      )[0]!
    );
    if (!compiled.success || compiled.target.plainText === null) throw Error('Invalid target');
    const planned = planParagraphText(
      { revision, partName: part.name, paragraphId: p.id, text: paragraphTextOf(part, p.id)! },
      compiled.target.plainText
    );
    if (!planned.success) throw Error(planned.error);
    const edits = linkAwareBias(part, p.id, 0, planned.plan.operations, compiled.target),
      owner = links[side === 'left' ? 0 : 1]!;
    expect(edits).toHaveLength(1);
    expect(edits[0]).toMatchObject({
      kind: 'text',
      start: 7,
      end: 7,
      after: 'X',
      ...(side === 'left' ? { insideHyperlink: owner.id } : {}),
    });
    const preview = previewTextEdits(part, edits);
    if (!preview.success) throw Error(preview.error);
    const expected = applyTreeOp(part, {
      op: 'insertText',
      paragraphId: p.id,
      offset: 7,
      text: 'X',
      inside: owner.id,
    });
    if (!expected.ok) throw Error(expected.reason);
    expect(canonicalOoxmlFingerprint(preview.part)).toBe(canonicalOoxmlFingerprint(expected.part));
  });
