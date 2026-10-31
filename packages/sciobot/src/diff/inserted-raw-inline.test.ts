import { expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { strFromU8 } from 'fflate';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
import { MDX_INLINE } from '../projection/mdx';
import { planTargetDocument } from './plan-target';

// Inline JSX without a registered component keeps its MDX tree and shows a readable label.
for (const [element, label] of [
  ['<qr-code url="https://example.com" label="Scan" />', 'Scan'],
  ['<law law="561/2004" citation="§ 3" />', 'law'],
] as const)
  test(`an unregistered inline element keeps its MDX and reads as "${label}"`, async () => {
    const input = await targetDocumentInput('<w:p><w:r><w:t>Seed.</w:t></w:r></w:p>');
    const target = `See ${element} here.\n\nSeed.`;
    const planned = planTargetDocument(input, target, { allocateBlockId: () => 'raw-inline' });
    if (!planned.success) throw Error(planned.error);
    const applied = previewDirectPlan(input.pkg, planned.plan);
    if (!applied.success) throw Error(applied.error);
    const block = buildBlockIndex(applied.pkg).byId.get('raw-inline')!;
    expect(block.payload.type).toBe(MDX_INLINE);
    expect((block.payload.mdast as { name?: string }).name).toBe(element.slice(1).split(' ')[0]);
    const reopened = readOoxmlPackage(writeOoxmlPackage(applied.pkg));
    if (!reopened.ok) throw Error(reopened.reason);
    const projection = await projectSourcePackage(reopened.package);
    expect(projection.text).toContain(element);
    // Word shows the readable label in place of the element.
    const xml = strFromU8(reopened.package.partBytes.get(reopened.package.mainDocumentPart)!);
    expect(xml).toContain(`>${label}</w:t>`);
    const again = planTargetDocument({ ...input, pkg: reopened.package, projection }, target);
    if (!again.success) throw Error(again.error);
    expect(again.plan.operations).toEqual([]);
  });

test('an inline element with children still refuses', async () => {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed.</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, 'See <qr-code url="x">text</qr-code>.\n\nSeed.');
  expect(planned.success).toBe(false);
});
