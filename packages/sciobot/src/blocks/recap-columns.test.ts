import { expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput, targetParagraph } from '../../test/support/target-document';
import { projectSourcePackage, sourceControl } from '../../test/support/target-source-fixture';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from './block-index';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { withRecapColumns } from './recap-columns';
import { readPresentationColumns } from './column-geometry';

for (const reverse of [false, true])
  test(`overview columns preserve original reading order (${reverse})`, async () => {
    const rows = [
      ['curriculum', 'recap_curriculum', 'RVP: Outcome'],
      ['materials', 'recap_materials', 'Cards'],
    ];
    if (reverse) rows.reverse();
    const old = await targetDocumentInput(
      sourceControl(
        'overview',
        rows.map(([id, , value]) => sourceControl(id!, targetParagraph(value!))).join('')
      ) + targetParagraph('Tail.')
    );
    let pkg = old.pkg;
    for (const [id, type] of [['overview', 'preparation_recap'], ...rows]) {
      const stored = withPayloadRecord(
        pkg,
        { id: id!, type: type!, version: 1, attributes: {} },
        id!
      );
      if (!stored.ok) throw new Error(stored.reason);
      pkg = stored.pkg;
    }
    const before = await projectSourcePackage(pkg);
    const prepared = withRecapColumns(pkg);
    const index = buildBlockIndex(prepared);
    const geometry = readPresentationColumns(prepared, index.byId.get('overview')!, index);
    expect(geometry?.columns.map((column) => column.block.id)).toEqual(rows.map((row) => row[0]!));
    expect(withRecapColumns(prepared)).toBe(prepared);
    const opened = readOoxmlPackage(writeOoxmlPackage(prepared));
    if (!opened.ok) throw new Error(opened.reason);
    const projection = await projectSourcePackage(opened.package);
    expect(projection.text).toBe(before.text);
    const input = { ...old, pkg: opened.package, projection };
    const unchanged = planTargetDocument(input, projection.text);
    if (!unchanged.success) throw new Error(unchanged.error);
    expect(unchanged.plan.operations).toEqual([]);
    const edit = planTargetDocument(input, projection.text.replace('Cards', 'Cards and dice'));
    if (!edit.success) throw new Error(edit.error);
    const applied = previewDirectPlan(input.pkg, edit.plan);
    if (!applied.success) throw new Error(applied.error);
    expect((await projectSourcePackage(applied.pkg)).text).toBe(
      projection.text.replace('Cards', 'Cards and dice')
    );
  });

test('new overviews retain goals and adaptations around the editable native columns', async () => {
  const input = await targetDocumentInput(targetParagraph('Tail.'));
  const target = [
    '<preparation_recap>',
    '<recap_goals>\n\nExplain the diagram\n\n</recap_goals>',
    '<recap_curriculum framework="RVP">\n\nDraw a diagram\n\n</recap_curriculum>',
    '<recap_materials>\n\nCards\n\n</recap_materials>',
    '<recap_adaptations>\n\nUse a large notebook\n\n</recap_adaptations>',
    '</preparation_recap>',
    'Tail.',
  ].join('\n\n');
  let n = 0;
  const created = planTargetDocument(input, target, { allocateBlockId: () => `overview-${++n}` });
  if (!created.success) throw new Error(created.error);
  const applied = previewDirectPlan(input.pkg, created.plan);
  if (!applied.success) throw new Error(applied.error);
  const index = buildBlockIndex(applied.pkg);
  const owner = [...index.byId.values()].find(
    (block) => block.payload.type === 'preparation_recap'
  )!;
  expect(readPresentationColumns(applied.pkg, owner, index)?.columns).toHaveLength(2);
  const projection = await projectSourcePackage(applied.pkg);
  for (const [oldValue, newValue] of [
    ['Cards', 'Dice'],
    ['diagram', 'picture'],
    ['notebook', 'grid'],
  ]) {
    const next = projection.text.replaceAll(oldValue!, newValue!);
    const edit = planTargetDocument({ ...input, pkg: applied.pkg, projection }, next);
    if (!edit.success) throw new Error(edit.error);
    const changed = previewDirectPlan(applied.pkg, edit.plan);
    if (!changed.success) throw new Error(changed.error);
    expect((await projectSourcePackage(changed.pkg)).text).toBe(next);
  }

  for (const removed of [
    ['recap_materials'],
    ['recap_curriculum'],
    ['recap_curriculum', 'recap_materials'],
  ]) {
    let next = projection.text;
    for (const tag of removed)
      next = next.replace(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>\\n\\n`), '');
    const edit = planTargetDocument({ ...input, pkg: applied.pkg, projection }, next);
    if (!edit.success) throw new Error(edit.error);
    const changed = previewDirectPlan(applied.pkg, edit.plan);
    if (!changed.success) throw new Error(changed.error);
    const reopened = readOoxmlPackage(writeOoxmlPackage(changed.pkg));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect((await projectSourcePackage(reopened.package)).text).toBe(next);
    const surviving = buildBlockIndex(reopened.package);
    const records = readPayloadStore(reopened.package);
    if (records.status !== 'ready') throw new Error(records.status);
    for (const block of index.byId.values()) {
      expect(surviving.byId.has(block.id)).toBe(!removed.includes(block.payload.type));
      expect(records.entries.has(block.id)).toBe(!removed.includes(block.payload.type));
    }
  }
});
