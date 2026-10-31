import { expect, test } from 'bun:test';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  contentControlTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import { readEditorRevision } from '../projection/revision-cache';
import { DirectTargetBaseline } from '../ai/direct-baseline';
import { applyDirectPlan } from '../ai/apply-direct';
import { recomposeOwnedProposals } from '../review/recompose';
import { readCommittedView } from '../review/committed-view';
import { readWorkingView } from '../review/working-view';
import { activeStorage } from '../registry/active';
import type { BlockPayload } from '../blocks/payload-schema';
import { planTargetDocument, type TargetDocumentOptions } from './plan-target';

const activity = (name: string, duration: number) =>
  `<phase type="evocation">\n\n<activity name="${name}" duration="${duration}">\n\nAsk a question.\n\n</activity>\n\n</phase>`;
const label = (payload: BlockPayload) =>
  payload.type === 'lesson_activity'
    ? `${String(payload.attributes.name)} — ${String(payload.attributes.duration)} min`
    : null;
const options: TargetDocumentOptions = { generatedLabel: label };

function editor() {
  const prepared = prepareNativeDocument(docxFixture('<w:p><w:r><w:t>Seed.</w:t></w:r></w:p>'));
  if (!prepared.success) throw Error(prepared.error);
  const read = readOoxmlPackage(prepared.document);
  if (!read.ok) throw Error(read.reason);
  return propertyEditor(read.package);
}
const labels = (instance: ReturnType<typeof editor>) => {
  const part = instance.surface!.session.part();
  return contentControlsIn(part.root)
    .filter((entry) => contentControlPropertiesOf(entry.node).tag === activeStorage().labelTag)
    .map((entry) => contentControlTextOf(entry.node, ' '));
};

test('a new container gets its Word label, which the projection hides', async () => {
  const instance = editor();
  const baseline = new DirectTargetBaseline();
  const before = await baseline.prepare(instance, propertyContext);
  if (!before.success) throw Error(before.error);
  const planned = planTargetDocument(before.input, `${activity('Warmup', 10)}\n\nSeed.`, options);
  if (!planned.success) throw Error(planned.error);
  const applied = applyDirectPlan(instance, planned.plan, propertyContext);
  if (!applied.success) throw Error(applied.error);
  expect(labels(instance)).toEqual(['Warmup — 10 min']);
  const landed = await baseline.prepare(instance, propertyContext);
  if (!landed.success) throw Error(landed.error);
  expect(landed.input.projection.text).not.toContain('Warmup — 10 min');
  const again = planTargetDocument(landed.input, `${activity('Warmup', 10)}\n\nSeed.`, options);
  if (!again.success) throw Error(again.error);
  expect(again.plan.operations).toEqual([]);

  // A property suggestion now has the label to rewrite.
  const revision = readEditorRevision(instance, propertyContext.documentId, propertyContext.epoch)!;
  const proposed = await recomposeOwnedProposals(
    instance,
    { revision, targetMarkdown: `${activity('Starter', 15)}\n\nSeed.`, ownership: [] },
    { ...propertyContext, author: 'AI', date: '2026-10-31T00:00:00.000Z' },
    {
      target: options,
      preparePayloads: (targets) => ({
        properties: {
          proposalId: 'property-1',
          fallbackTargets: targets.map((target) => ({
            blockId: target.blockId,
            before: target.beforeText,
            after: 'Starter — 15 min',
          })),
        },
      }),
    }
  );
  if (!proposed.success) throw Error(proposed.error);
  const token = readEditorRevision(instance, propertyContext.documentId, propertyContext.epoch)!;
  const [committed, working] = await Promise.all([
    readCommittedView(instance.surface!.session, token),
    readWorkingView(instance.surface!.session, token),
  ]);
  if (!committed.success || !working.success) throw Error('unreadable views');
  expect(committed.view.projection.text).toContain('name="Warmup" duration="10"');
  expect(working.view.projection.text).toContain('name="Starter" duration="15"');
});

test('a multi-line or oversized label refuses the plan', async () => {
  const instance = editor();
  const before = await new DirectTargetBaseline().prepare(instance, propertyContext);
  if (!before.success) throw Error(before.error);
  for (const bad of ['Two\nlines', 'x'.repeat(513)]) {
    const planned = planTargetDocument(before.input, `${activity('Warmup', 10)}\n\nSeed.`, {
      generatedLabel: (payload) => (payload.type === 'lesson_activity' ? bad : null),
    });
    expect(planned.success).toBe(false);
  }
});
