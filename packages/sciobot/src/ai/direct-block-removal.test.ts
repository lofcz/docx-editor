import { expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { previewDirectPlan } from './apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';

const BLOCKS: readonly (readonly [string, string])[] = [
  ['code', '```python\ndef f():\n    return 1\n```'],
  ['visual', '<lesson_visual id="v" src="https://example.com/a.png" alt="A" />'],
  ['callout', '<callout variant="info">\n\nInside the callout.\n\n</callout>'],
];

for (const [name, markdown] of BLOCKS)
  test(`a direct edit removes a ${name} block and its payload record`, async () => {
    const input = await targetDocumentInput('<w:p><w:r><w:t>Seed text</w:t></w:r></w:p>');
    let next = 0;
    const inserted = planTargetDocument(input, `${markdown}\n\nSeed text`, {
      allocateBlockId: () => `${name}-${++next}`,
    });
    if (!inserted.success) throw new Error(inserted.error);
    const withBlock = previewDirectPlan(input.pkg, inserted.plan);
    if (!withBlock.success) throw new Error(withBlock.error);
    const ids = [...buildBlockIndex(withBlock.pkg).byId.keys()];
    expect(ids.length).toBeGreaterThan(0);

    const projection = await projectSourcePackage(withBlock.pkg);
    const removal = planTargetDocument({ ...input, pkg: withBlock.pkg, projection }, 'Seed text');
    if (!removal.success) throw new Error(removal.error);
    const removed = previewDirectPlan(withBlock.pkg, removal.plan);
    if (!removed.success) throw new Error(removed.error);

    const reopened = readOoxmlPackage(writeOoxmlPackage(removed.pkg));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(buildBlockIndex(reopened.package).byId.size).toBe(0);
    const store = readPayloadStore(reopened.package);
    if (store.status === 'ready') for (const id of ids) expect(store.entries.has(id)).toBe(false);
    expect((await projectSourcePackage(reopened.package)).text.trim()).toBe('Seed text');
  });
