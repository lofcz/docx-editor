import { expect, test } from 'bun:test';
import {
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput, targetParagraph as p } from '../../test/support/target-document';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { ast, part, nativeTable } from '../../test/support/target-table-fixture';
import { rowCells, tableRows } from '../../test/support/native-table-oracle';
import { planTargetDocument } from './plan-target';
import { bindLogicalSource } from './bind-source';
import { prepareTargetCellParagraphs } from './target-cell-paragraphs';

const table = (content: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>${content}</w:tc><w:tc>${p('Other')}</w:tc></w:tr></w:tbl>`;
for (const item of [
  {
    name: 'second paragraph text',
    target: 'First<br>Changed',
    texts: ['First', 'Changed'],
    retained: [0, 1],
    kinds: ['text', 'text'],
  },
  {
    name: 'insert middle paragraph',
    target: 'First<br>New<br>Second',
    texts: ['First', 'New', 'Second'],
    retained: [0, 1],
    kinds: ['siblings'],
  },
  {
    name: 'remove second paragraph',
    target: 'First',
    texts: ['First'],
    retained: [0],
    kinds: ['siblings'],
  },
  {
    name: 'remove first paragraph',
    target: 'Second',
    texts: ['Second'],
    retained: [1],
    kinds: ['siblings'],
  },
  {
    name: 'second paragraph marks',
    target: 'First<br>**Second**',
    texts: ['First', 'Second'],
    retained: [0, 1],
    kinds: ['format'],
  },
] as const) {
  test(`GFM multi-paragraph cell plans ${item.name} without replacing surviving owners`, async () => {
    const input = await targetDocumentInput(table(p('First') + p('Second')));
    const original = storyParagraphs(
      rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[0]!
    );
    const target = `| ${item.target} | Other |\n| --- | --- |`;
    const before = writeOoxmlPackage(input.pkg);
    const result = planTargetDocument(input, target);
    if (!result.success) throw new Error(result.error);
    expect(result.plan.operations.map((op) => op.kind)).toEqual([...item.kinds]);
    if (item.name === 'second paragraph text' || item.name === 'second paragraph marks') {
      for (const operation of result.plan.operations) {
        if (operation.kind === 'text') expect(operation.paragraphId).toBe(original[1]!.id);
        if (operation.kind === 'format')
          expect(operation.target).toMatchObject({ paragraphId: original[1]!.id });
      }
    }
    const applied = applyTargetPlanOracle(input.pkg, result.plan);
    const actual = applied.parts.get(applied.mainDocumentPart)!;
    expect(
      storyParagraphs(
        findNode(actual, rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[0]!.id)!
      ).map((node) => paragraphTextOf(actual, node.id))
    ).toEqual([...item.texts]);
    for (const index of item.retained) expect(findNode(actual, original[index]!.id)).not.toBeNull();
    expect(writeOoxmlPackage(input.pkg)).toEqual(before);
    const reopened = readOoxmlPackage(writeOoxmlPackage(applied));
    if (!reopened.ok) throw new Error(reopened.reason);
    const projection = await projectSourcePackage(reopened.package);
    expect(ast(projection.text)).toEqual(ast(target));
    const repeated = planTargetDocument({ ...input, pkg: reopened.package, projection }, target);
    expect(repeated.success && repeated.plan.operations).toEqual([]);
  });
}

test('authored manual breaks remain inside their original paragraph beside generated separators', async () => {
  const manual = '<w:p><w:r><w:t>Manual</w:t><w:br/><w:t>line</w:t></w:r></w:p>';
  const input = await targetDocumentInput(table(manual + p('Second')));
  const target = '| Manual<br>line<br>Changed | Other |\n| --- | --- |';
  const result = planTargetDocument(input, target);
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations.map((op) => op.kind)).toEqual(['text', 'text']);
  const applied = applyTargetPlanOracle(input.pkg, result.plan);
  const actual = applied.parts.get(applied.mainDocumentPart)!;
  const cell = rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[0]!;
  expect(
    storyParagraphs(findNode(actual, cell.id)!).map((node) => paragraphTextOf(actual, node.id))
  ).toEqual(['Manual\nline', 'Changed']);
  expect(storyParagraphs(findNode(actual, cell.id)!)[0]).toBe(storyParagraphs(cell)[0]);
  const reopened = readOoxmlPackage(writeOoxmlPackage(applied));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(target));
});

for (const item of [
  { target: 'Manual<br>edited<br>Second', texts: ['Manual\nedited', 'Second'] },
  { target: 'Manual<br>line<br>New<br>Second', texts: ['Manual\nline', 'New', 'Second'] },
  { target: 'Manual<br>line', texts: ['Manual\nline'] },
] as const) {
  test(`authored break ownership survives target ${item.target}`, async () => {
    const input = await targetDocumentInput(
      table('<w:p><w:r><w:t>Manual</w:t><w:br/><w:t>line</w:t></w:r></w:p>' + p('Second'))
    );
    const cell = rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[0]!;
    const first = storyParagraphs(cell)[0]!;
    const target = `| ${item.target} | Other |\n| --- | --- |`;
    const result = planTargetDocument(input, target);
    if (!result.success) throw new Error(result.error);
    const applied = applyTargetPlanOracle(input.pkg, result.plan);
    const actual = applied.parts.get(applied.mainDocumentPart)!;
    expect(
      storyParagraphs(findNode(actual, cell.id)!).map((node) => paragraphTextOf(actual, node.id))
    ).toEqual([...item.texts]);
    expect(findNode(actual, first.id)).not.toBeNull();
    const reopened = readOoxmlPackage(writeOoxmlPackage(applied));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(target));
  });
}

test('formatting across a generated separator keeps both native paragraphs', async () => {
  const input = await targetDocumentInput(table(p('First') + p('Second')));
  const result = planTargetDocument(input, '| **First<br>Second** | Other |\n| --- | --- |');
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations.map((op) => op.kind)).toEqual(['format', 'format']);
  const applied = applyTargetPlanOracle(input.pkg, result.plan);
  const actual = applied.parts.get(applied.mainDocumentPart)!;
  const oldCell = rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[0]!;
  expect(storyParagraphs(findNode(actual, oldCell.id)!).map((node) => node.id)).toEqual(
    storyParagraphs(oldCell).map((node) => node.id)
  );
  const reopened = readOoxmlPackage(writeOoxmlPackage(applied));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect((await projectSourcePackage(reopened.package)).text).toBe(
    '| **First**<br>**Second** | Other |\n| --- | --- |'
  );
});

test('unchanged composite cell targets preserve the original bound view and need no operations', async () => {
  const input = await targetDocumentInput(table(p('First') + p('Second')));
  const bound = bindLogicalSource(input);
  if (!bound.success) throw new Error(bound.error);
  const result = prepareTargetCellParagraphs(input, bound.source, bound.source.logical);
  if (!result.success) throw new Error(result.error);
  expect(result.source).toBe(bound.source);
  expect(result.target).toBe(bound.source.logical);
  const plan = planTargetDocument(input, input.projection.text);
  expect(plan.success && plan.plan.operations).toEqual([]);
});

for (const texts of [
  ['Same', 'Same'],
  ['😀 first', 'Žluťoučký 😀'],
] as const) {
  test(`repeated and Unicode paragraph identities survive GFM edits: ${texts[0]}`, async () => {
    const input = await targetDocumentInput(table(p(texts[0]) + p(texts[1])));
    const oldCell = rowCells(tableRows(part(input), nativeTable(input).id)[0]!)[0]!;
    const original = storyParagraphs(oldCell);
    const target = `| ${texts[0]}<br>${texts[1]} changed | Other |\n| --- | --- |`;
    const result = planTargetDocument(input, target);
    if (!result.success) throw new Error(result.error);
    for (const operation of result.plan.operations) {
      expect(operation.kind).toBe('text');
      if (operation.kind === 'text') expect(operation.paragraphId).toBe(original[1]!.id);
    }
    const pkg = applyTargetPlanOracle(input.pkg, result.plan);
    const actual = pkg.parts.get(pkg.mainDocumentPart)!;
    expect(findNode(actual, original[0]!.id)).toBe(original[0]);
    expect(paragraphTextOf(actual, original[1]!.id)).toBe(`${texts[1]} changed`);
    const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(target));
  });
}
