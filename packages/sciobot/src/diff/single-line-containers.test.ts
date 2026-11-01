import { expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from './plan-target';

// Generated lessons write a short container on one line: `<lesson_title …>Text</lesson_title>`.
// Markdown parses that as a paragraph that holds an inline element with text and no paragraph.
const SINGLE = '<lesson_title grade="6. třída" lesson_minutes="45">Fotosyntéza</lesson_title>';
const BLOCK =
  '<lesson_title grade="6. třída" lesson_minutes="45">\n\nFotosyntéza\n\n</lesson_title>';

async function land(target: string) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed.</w:t></w:r></w:p>');
  let ids = 0;
  const planned = planTargetDocument(input, `${target}\n\nSeed.`, {
    allocateBlockId: () => `block-${++ids}`,
  });
  if (!planned.success) throw Error(`${planned.code}: ${planned.error}`);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw Error(applied.error);
  return { input, pkg: applied.pkg };
}

test('a one-line container lands as the same native block as its block form', async () => {
  const single = await land(SINGLE),
    block = await land(BLOCK);
  const [title] = [...buildBlockIndex(single.pkg).byId.values()];
  expect(title!.payload.type).toBe('lesson_title');
  expect(title!.payload.attributes).toEqual(
    [...buildBlockIndex(block.pkg).byId.values()][0]!.payload.attributes
  );
  const projection = await projectSourcePackage(single.pkg);
  expect(projection.text).toBe((await projectSourcePackage(block.pkg)).text);
  expect(projection.text).toContain('Fotosyntéza');
  const reopened = readOoxmlPackage(writeOoxmlPackage(single.pkg));
  if (!reopened.ok) throw Error(reopened.reason);
  expect((await projectSourcePackage(reopened.package)).text).toBe(projection.text);
  // Both forms replan the landed document to no change.
  for (const target of [SINGLE, BLOCK]) {
    const again = planTargetDocument(
      {
        ...single.input,
        pkg: reopened.package,
        projection: await projectSourcePackage(reopened.package),
      },
      `${target}\n\nSeed.`
    );
    if (!again.success) throw Error(again.error);
    expect(again.plan.operations).toEqual([]);
  }
});

test('a one-line container keeps inline marks in its paragraph', async () => {
  const { pkg } = await land('<lesson_title grade="7">Light *and* **energy**</lesson_title>');
  const text = (await projectSourcePackage(pkg)).text;
  expect(text).toContain('Light *and* **energy**');
});
