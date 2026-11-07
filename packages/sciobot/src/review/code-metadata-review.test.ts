import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode, readOoxmlPackage, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument } from '../../test/support/code';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { updateAtomicBlock } from '../blocks/update-atomic';
import { nativeCodeSource } from '../blocks/code-native-source';
import { planAtomicBlock } from '../diff/atomic-blocks';
import { readReviewGroups } from './groups';
import { projectPendingProperties } from './property-projection';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { trackAtomicPlan } from './track-atomic';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const ORIGINAL = 'one\nmiddle\nlast';
const PROPOSED = 'one\n\tmid\ndle\nlast';
function propose(editor: DocxEditorInstance, field: 'source' | 'language' | 'both') {
  const pkg = editor.surface!.session.currentPackage();
  const payload = buildBlockIndex(pkg).byId.get('code')!.payload;
  const plan = planAtomicBlock(
    { pkg, revision: propertyToken(editor), blockId: 'code' },
    {
      ...payload,
      attributes: {
        ...payload.attributes,
        ...(field !== 'language' ? { value: PROPOSED } : {}),
        ...(field !== 'source' ? { language: 'typescript' } : {}),
      },
    }
  );
  if (!plan.success) throw Error(plan.error);
  const result = trackAtomicPlan(
    editor,
    plan.plan,
    [{ blockId: 'code', before: ORIGINAL, after: field !== 'language' ? PROPOSED : ORIGINAL }],
    {
      ...propertyContext,
      author: field === 'source' ? 'Alice' : 'Bob',
      date: '2026-11-06T12:00:00.000Z',
      proposalId: `native_code_${field}`,
    }
  );
  if (!result.success) throw Error(result.error);
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of ['source', 'language', 'both'] as const)
    test(`complete-payload metadata preserves pending code source and language: ${clientIds}/${first}`, async () => {
      const pair = await createTwoEditorFixture({
        document: await codeDocument(),
        clientIds,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        propose(alice, first);
        await pair.settle();
        const originalPart = bob.surface!.session.part();
        const originalXml = serializeOoxmlPart(originalPart);
        const payload = buildBlockIndex(bob.surface!.session.currentPackage()).byId.get(
          'code'
        )!.payload;
        expect(
          updateAtomicBlock(bob, {
            blockId: 'code',
            attributes: {
              ...payload.attributes,
              ...(first === 'source' ? { language: 'javascript' } : {}),
              opaque: { peer: true },
            },
          })
        ).toMatchObject({ success: true });
        await pair.settle();
        for (const editor of pair.editors) {
          const pkg = editor!.surface!.session.currentPackage();
          expect(serializeOoxmlPart(editor!.surface!.session.part())).toBe(originalXml);
          const stored = readStoredPropertyProposals(pkg, propertyToken(editor!));
          if (!stored.success) throw Error(stored.error);
          const projected = projectPendingProperties(
            pkg,
            propertyToken(editor!),
            stored.pending,
            'working'
          );
          if (!projected.success) throw Error(projected.error);
          const shown = buildBlockIndex(projected.pkg).byId.get('code')!.payload;
          expect(shown.attributes.language).toBe(first === 'source' ? 'javascript' : 'typescript');
          expect(shown.attributes.opaque).toEqual({ peer: true });
          const part = editor!.surface!.session.part();
          const control = findNode(part, buildBlockIndex(pkg).byId.get('code')!.controlId)!;
          expect(nativeCodeSource(part, control, 'committed')!.value).toBe(ORIGINAL);
          expect(nativeCodeSource(part, control, 'working')!.value).toBe(
            first === 'language' ? ORIGINAL : PROPOSED
          );
          const groups = readReviewGroups(pkg);
          if (!groups.success) throw Error(groups.error);
          expect(groups.groups).toHaveLength(1);
        }
        expect(bob.surface!.session.part()).toBe(originalPart);
        const beforeRefusal = bob.surface!.session.currentPackage();
        if (first !== 'language')
          expect(
            updateAtomicBlock(bob, {
              blockId: 'code',
              attributes: { ...payload.attributes, value: 'replace pending source' },
            })
          ).toMatchObject({ success: false, code: 'unsupported-operation' });
        expect(bob.surface!.session.currentPackage()).toBe(beforeRefusal);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(serializeOoxmlPart(editor!.surface!.session.part())).toBe(originalXml);
          expect(
            buildBlockIndex(editor!.surface!.session.currentPackage()).byId.get('code')!.payload
              .attributes.opaque
          ).toBeUndefined();
        }
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        const saved = readOoxmlPackage(new Uint8Array(await alice.save()));
        if (!saved.ok) throw Error(saved.reason);
        const stored = readStoredPropertyProposals(saved.package, propertyToken(alice));
        if (!stored.success) throw Error(stored.error);
        expect(stored.pending).toHaveLength(1);
        expect(buildBlockIndex(saved.package).byId.get('code')!.payload.attributes.opaque).toEqual({
          peer: true,
        });
        const reopenedBlock = buildBlockIndex(saved.package).byId.get('code')!;
        const reopenedPart = saved.package.parts.get(reopenedBlock.partName)!;
        const reopenedControl = findNode(reopenedPart, reopenedBlock.controlId)!;
        expect(nativeCodeSource(reopenedPart, reopenedControl, 'committed')!.value).toBe(ORIGINAL);
        expect(nativeCodeSource(reopenedPart, reopenedControl, 'working')!.value).toBe(
          first === 'language' ? ORIGINAL : PROPOSED
        );
        const reopenedView = projectPendingProperties(
          saved.package,
          propertyToken(alice),
          stored.pending,
          'working'
        );
        if (!reopenedView.success) throw Error(reopenedView.error);
        expect(
          buildBlockIndex(reopenedView.pkg).byId.get('code')!.payload.attributes.language
        ).toBe(first === 'source' ? 'javascript' : 'typescript');
        expect(serializeOoxmlPart(alice.surface!.session.part())).toBe(originalXml);
      } finally {
        pair.destroy();
      }
    });
