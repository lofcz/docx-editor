import { expect } from 'bun:test';
import {
  bodyStoryRoot,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetParagraph as p } from './target-document';
import { applyTargetPlanOracle } from './target-plan-oracle';
import { projectSourcePackage } from './target-source-fixture';
import { nodeChildren, tableRows, rowCells } from './native-table-oracle';
import { parseTargetMarkdown } from '../../src/projection/parse-target';
import { planTargetDocument } from '../../src/diff/plan-target';
import type { TargetSourceInput } from '../../src/diff/target-source-types';

export const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><keep:cell xmlns:keep="urn:keep"/></w:tcPr>${p(text)}</w:tc>`;
export const row = (values: readonly string[]) =>
  `<w:tr><w:trPr><w:cantSplit/><keep:row xmlns:keep="urn:keep"/></w:trPr>${values.map(cell).join('')}</w:tr>`;
export const table = (rows: readonly (readonly string[])[]) =>
  `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/><keep:table xmlns:keep="urn:keep"/></w:tblPr><w:tblGrid>${rows[0]!.map(() => '<w:gridCol w:w="2400"/>').join('')}</w:tblGrid>${rows.map(row).join('')}</w:tbl>`;
export const part = (input: TargetSourceInput) => input.pkg.parts.get(input.pkg.mainDocumentPart)!;
export const nativeTable = (input: TargetSourceInput) =>
  nodeChildren(bodyStoryRoot(part(input))!).find((node) => node.kind === 'table')!;
export const content = (input: TargetSourceInput) =>
  tableRows(part(input), nativeTable(input).id).map((row) =>
    rowCells(row).map((cell) =>
      storyParagraphs(cell).map((node) => paragraphTextOf(part(input), node.id))
    )
  );
export function ast(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst;
}
export function plan(input: TargetSourceInput, target: string) {
  const result = planTargetDocument(input, target);
  if (!result.success) throw new Error(result.error);
  return result;
}
export async function landed(input: TargetSourceInput, target: string) {
  const result = plan(input, target),
    pkg = applyTargetPlanOracle(input.pkg, result.plan),
    reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(target));
  return { result, input: { ...input, pkg } };
}
