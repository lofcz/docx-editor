import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import type { BlockPayload, JsonObject } from '../../src/blocks/payload-schema';
import { planAtomicBlock } from '../../src/diff/atomic-blocks';
import { parseTargetMarkdown } from '../../src/projection/parse-target';
import { createEditPlan } from '../../src/diff/plan';
import { MDX_BLOCK } from '../../src/projection/mdx';
import { atomicFallbackParagraph } from '../../src/review/property-fallback';
import { trackAtomicPlan } from '../../src/review/track-atomic';
import { docxFixture } from './document';
import { propertyContext, propertyToken } from './property-execution';
import { sourceControl, sourceParagraph } from './target-source-fixture';

export const atomicVisual: BlockPayload = {
  id: 'visual',
  type: 'lesson_visual',
  version: 1,
  attributes: {
    kind: 'explainer',
    phase: 'evocation',
    src: 'vfs://pristine.png',
    fileId: 'original-file',
    alt: 'Old picture',
    prompt: 'Old prompt',
  },
  extension: { originalAsset: 'keep-original', opaque: [1, null, { keep: true }] },
};
export const proposedVisual: BlockPayload = {
  ...atomicVisual,
  attributes: {
    ...atomicVisual.attributes,
    src: 'vfs://new.png',
    fileId: 'new-file',
    alt: 'New picture',
    prompt: 'New prompt',
  },
};
function listeningAst(title: string, audio: string, line: string) {
  const parsed = parseTargetMarkdown(
    `<listening key="budget-keep" title="${title}" audio="${audio}" seed="7">\n\n<speaker id="s1" name="Anna" voice="voice-keep" portrait="https://example.com/anna.png" />\n\n<line speaker="s1">${line}</line>\n\n</listening>`
  );
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.normalizedAst.children[0] as JsonObject;
}
export const atomicListening: BlockPayload = {
  id: 'audio',
  type: MDX_BLOCK,
  version: 1,
  attributes: { cache: { asset: 'cached-audio-keep' } },
  mdast: listeningAst('Old listening', 'original-audio', 'Good morning!'),
  opaqueExtension: { codecVersion: 3, voice: 'voice-keep', portrait: 'portrait-keep' },
};
export const proposedListening: BlockPayload = {
  ...atomicListening,
  attributes: { cache: { asset: 'new-cached-audio' } },
  mdast: listeningAst('New listening', 'new-audio', 'Good evening!'),
};
export function atomicFixture(payloads: readonly BlockPayload[] = [atomicVisual, atomicListening]) {
  const loaded = readOoxmlPackage(
    docxFixture(
      sourceControl(
        'phase',
        sourceParagraph('Phase body') +
          payloads
            .map((payload) =>
              sourceControl(payload.id, sourceParagraph(`Old ${payload.id} fallback`))
            )
            .join('')
      ) + sourceParagraph('Keep tail')
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  let pkg = loaded.package;
  for (const payload of [
    {
      id: 'phase',
      type: 'lesson_phase',
      version: 1,
      attributes: { phaseType: 'evocation' },
    } as BlockPayload,
    ...payloads,
  ]) {
    const written = withPayloadRecord(pkg, payload, `Stored ${payload.id} label`);
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return pkg;
}
export function atomicLabel(editor: DocxEditorInstance, id = 'visual') {
  const session = editor.surface!.session,
    block = buildBlockIndex(session.currentPackage()).byId.get(id)!,
    paragraph = atomicFallbackParagraph(session.part(), block.controlId);
  if (!paragraph) throw new Error('Missing atomic fallback');
  return paragraph;
}
export function atomicPlan(editor: DocxEditorInstance, targets: readonly BlockPayload[]) {
  const session = editor.surface!.session,
    revision = propertyToken(editor),
    operations = targets.flatMap((target) => {
      const result = planAtomicBlock(
        { pkg: session.currentPackage(), revision, blockId: target.id },
        target
      );
      if (!result.success) throw new Error(result.error);
      return result.plan.operations;
    }),
    plan = createEditPlan(revision, operations);
  if (!plan.success) throw new Error(plan.error);
  return plan.plan;
}
export function executeAtomic(
  editor: DocxEditorInstance,
  targets: readonly BlockPayload[] = [proposedVisual],
  proposalId = 'atomic-proposal',
  after?: string
) {
  return trackAtomicPlan(
    editor,
    atomicPlan(editor, targets),
    targets.map((target) => ({
      blockId: target.id,
      before: paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor, target.id).id)!,
      after: after ?? `New ${target.id} fallback`,
    })),
    { ...propertyContext, proposalId }
  );
}
