import { expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../store/package/ooxml-tree.ts';
import { documentReads } from '../reads.ts';
import {
  joinedParagraphGroups,
  logicalParagraphProjection,
  logicalParagraphPoint,
  logicalParagraphRange,
} from '../joined-paragraphs.ts';
import { docx, p, table, row, cell } from './support/protocol.ts';

const joined = (value: string) =>
  `<w:p><w:pPr><state:joined xmlns:state="urn:docx-editor.dev:review:retired-paragraph:1"/></w:pPr><w:r><w:t>${value}</w:t></w:r></w:p>`;
for (const [body, expected] of [
  [joined('A') + p('B'), ['AB']],
  [joined('same') + joined('😀') + p('same'), ['same😀same']],
  [joined('A') + table(row(cell(p('B')), cell(p('C')))), ['AB', 'C']],
  [`<w:sdt><w:sdtContent>${joined('A')}</w:sdtContent></w:sdt>${p('B')}`, ['A', 'B']],
  [`${joined('A')}<w:sdt><w:sdtContent>${p('B')}</w:sdtContent></w:sdt>${p('C')}`, ['A', 'B', 'C']],
] as const)
  test(`logical paragraph grouping preserves native structure: ${JSON.stringify(expected)}`, () => {
    const loaded = readOoxmlPackage(docx(body));
    if (!loaded.ok) throw Error('Invalid fixture');
    const reads = documentReads(loaded.package).body!;
    const before = canonicalOoxmlFingerprint(reads.part);
    const groups = joinedParagraphGroups(reads).ordered;
    expect(
      groups.map((group) => logicalParagraphProjection(reads, group.id, 'allMarkup')!.text)
    ).toEqual(expected);
    for (const group of groups) {
      const range = logicalParagraphRange(reads, group.id)!;
      expect(range.start.paragraphId).toBe(group.fragments[0]!);
      expect(range.end.paragraphId).toBe(group.fragments.at(-1)!);
      const length = group.fragments.reduce((sum, id) => sum + reads.rawText(id)!.length, 0);
      expect(logicalParagraphPoint(reads, group.id, length)!.offset).toBe(
        reads.rawText(group.fragments.at(-1)!)!.length
      );
    }
    expect(canonicalOoxmlFingerprint(reads.part)).toBe(before);
  });
