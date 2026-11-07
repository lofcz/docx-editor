import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  findNode,
  isJoinedParagraph,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument, openCode } from '../../test/support/code';
import { buildBlockIndex } from '../blocks/block-index';
import { updateCodeLanguage, updateCodeSource } from '../blocks/code-edit';
import { nativeCodeSource } from '../blocks/code-native-source';
import { codeSourceReviewLocks, stageTrackedCodeSource } from './code-source-review';
import type { TrackedStageWriter } from './tracked-stage-writer';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { planAtomicBlock } from '../diff/atomic-blocks';
import { indexTargetSource } from '../diff/target-source';
import { createPendingPropertyProposal, projectPendingProperties } from './property-projection';
import { decideAtomicProposal, trackAtomicPlan } from './track-atomic';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readReviewGroups } from './groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

const ORIGINAL = 'one\nmiddle\nlast';
const PROPOSED = 'one\n\tmid\ndle\nlast';
const attribution = { author: 'Alice', date: '2026-11-05T12:00:00.000Z' };
function transaction(editor: DocxEditorInstance, stage: (writer: TrackedStageWriter) => void) {
  const surface = editor.surface!;
  const result = surface.applyAutomationOps(
    () => [],
    { kind: 'body' },
    undefined,
    false,
    true,
    'before',
    (ctx) =>
      stage({
        currentPackage: () => ctx.currentPackage(),
        edit(change) {
          if (!ctx.applyPackage(change)) throw Error('The code review package stage failed.');
        },
        apply(partName, op) {
          if (!ctx.applyTo(partName, op)) throw Error(`The code review operation failed: ${op.op}`);
        },
      })
  );
  if (result.rejected || !result.committed) throw Error(String(result.reason));
}
function state(editor: DocxEditorInstance) {
  const pkg = editor.surface!.session.currentPackage();
  const block = buildBlockIndex(pkg).byId.get('code')!;
  const part = pkg.parts.get(block.partName)!;
  const control = findNode(part, block.controlId)!;
  return {
    pkg,
    block,
    part,
    committed: nativeCodeSource(part, control, 'committed')!,
    working: nativeCodeSource(part, control, 'working')!,
  };
}

for (const decision of ['accept', 'reject'] as const)
  for (const executor of ['proposal', 'group'] as const)
    test(`the native code review executor ${decision}/${executor} retains source, peer metadata, and native group membership`, async () => {
      const pair = await createTwoEditorFixture({
        document: await codeDocument(),
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const alice = pair.editors[0]!;
        const initial = state(alice);
        const plan = planAtomicBlock(
          { pkg: initial.pkg, revision: propertyToken(alice), blockId: 'code' },
          {
            ...initial.block.payload,
            attributes: { ...initial.block.payload.attributes, value: PROPOSED },
          }
        );
        if (!plan.success) throw Error(plan.error);
        const staged = trackAtomicPlan(
          alice,
          plan.plan,
          [{ blockId: 'code', before: ORIGINAL, after: PROPOSED }],
          {
            ...propertyContext,
            ...attribution,
            proposalId: 'native_code_executor',
          }
        );
        if (!staged.success) throw Error(staged.error);
        expect(staged.proposal!.nativeRevisions.length).toBeGreaterThan(0);
        await pair.settle();
        for (const editor of pair.editors) {
          const current = state(editor!);
          expect(current.committed.value).toBe(ORIGINAL);
          expect(current.working.value).toBe(PROPOSED);
          const groups = readReviewGroups(current.pkg);
          if (!groups.success) throw Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          expect(groups.groups[0]!.readOnly).toBe(false);
          const projected = await readWorkingView(editor!.surface!.session, propertyToken(editor!));
          if (!projected.success) throw Error(projected.error);
          expect(projected.view.projection.text).toContain('```python\n' + PROPOSED + '\n```');
        }
        const bob = pair.editors[1]!;
        expect(updateCodeLanguage(bob, { blockId: 'code', language: 'typescript' })).toMatchObject({
          success: true,
        });
        await pair.settle();
        const groups = readReviewGroups(state(alice).pkg);
        if (!groups.success) throw Error(groups.error);
        const decided =
          executor === 'proposal'
            ? decideAtomicProposal(
                alice,
                'native_code_executor',
                decision,
                propertyToken(alice),
                propertyContext
              )
            : (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
                alice,
                groups.groups[0]!.id,
                propertyToken(alice),
                propertyContext
              );
        if (!decided.success) throw Error(decided.error);
        await pair.settle();
        for (const editor of pair.editors) {
          const current = state(editor!);
          expect(current.committed.value).toBe(decision === 'accept' ? PROPOSED : ORIGINAL);
          expect(current.working.value).toBe(current.committed.value);
          expect(current.block.payload.attributes.language).toBe('typescript');
        }
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const current = state(editor!);
          expect(current.committed.value).toBe(ORIGINAL);
          expect(current.working.value).toBe(PROPOSED);
          expect(current.block.payload.attributes.language).toBe('typescript');
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(state(editor!).working.value).toBe(decision === 'accept' ? PROPOSED : ORIGINAL);
      } finally {
        pair.destroy();
      }
    });

test('code review reads exact native source halves without publishing payload cache changes', async () => {
  const pair = await createTwoEditorFixture({
    document: await codeDocument(),
    ownedAttributes: true,
  });
  try {
    const editor = pair.editors[0]!;
    transaction(editor, (writer) =>
      stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
    );
    await pair.settle();
    for (const peer of pair.editors) {
      const before = peer!.surface!.session.currentPackage();
      const committed = await readCommittedView(peer!.surface!.session, propertyToken(peer!));
      const working = await readWorkingView(peer!.surface!.session, propertyToken(peer!));
      if (!committed.success) throw Error(committed.error);
      if (!working.success) throw Error(working.error);
      expect(committed.view.projection.text).toContain('```python\n' + ORIGINAL + '\n```');
      expect(working.view.projection.text).toContain('```python\n' + PROPOSED + '\n```');
      expect(peer!.surface!.session.currentPackage()).toBe(before);
      expect(state(peer!).block.payload.attributes.value).toBe(ORIGINAL);
    }
  } finally {
    pair.destroy();
  }
});

test('pending native source review preserves a peer language change in both payload views', async () => {
  const pair = await createTwoEditorFixture({
    document: await codeDocument(),
    ownedAttributes: true,
  });
  try {
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    const initial = state(alice);
    const plan = planAtomicBlock(
      { pkg: initial.pkg, revision: propertyToken(alice), blockId: 'code' },
      {
        ...initial.block.payload,
        attributes: { ...initial.block.payload.attributes, value: PROPOSED },
      }
    );
    if (!plan.success) throw Error(plan.error);
    const pending = createPendingPropertyProposal('native_code', plan.plan);
    if (!pending.success) throw Error(pending.error);
    transaction(alice, (writer) =>
      stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
    );
    await pair.settle();
    expect(updateCodeLanguage(bob, { blockId: 'code', language: 'typescript' })).toMatchObject({
      success: true,
    });
    await pair.settle();
    for (const editor of pair.editors) {
      const current = state(editor!);
      for (const view of ['committed', 'working'] as const) {
        const projected = projectPendingProperties(
          current.pkg,
          propertyToken(editor!),
          [pending.proposal],
          view
        );
        if (!projected.success) throw Error(projected.error);
        expect(projected.pkg).toBe(current.pkg);
        expect(buildBlockIndex(projected.pkg).byId.get('code')!.payload.attributes.language).toBe(
          'typescript'
        );
      }
      expect(current.committed.value).toBe(ORIGINAL);
      expect(current.working.value).toBe(PROPOSED);
    }
  } finally {
    pair.destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`native code source ${decision} retains peer language, exact history, cursors, and reopen: ${clientIds}`, async () => {
      const pair = await createTwoEditorFixture({
        document: await codeDocument(),
        clientIds,
        ownedAttributes: true,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const initial = state(alice);
        const originalXml = serializeOoxmlPart(initial.part);
        const caret = { paragraphId: initial.committed.paragraphs[2]!.id, offset: 2 };
        bob.surface!.setSelection({ anchor: caret, head: caret });
        await pair.settle();
        pair.pause();
        transaction(alice, (writer) =>
          stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
        );
        expect(updateCodeLanguage(bob, { blockId: 'code', language: 'typescript' })).toMatchObject({
          success: true,
        });
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const verify = (committed: string, working: string) => {
          for (const editor of pair.editors) {
            const current = state(editor);
            expect(current.committed.value).toBe(committed);
            expect(current.working.value).toBe(working);
            expect(current.block.payload.attributes.language).toBe('typescript');
          }
        };
        verify(ORIGINAL, PROPOSED);
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote).toBeDefined();
        expect(
          caretAt(alice.surface!.layout(), {
            paragraphId: remote.head.nodeId,
            offset: remote.head.offset,
          })
        ).not.toBeNull();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, ORIGINAL);
        expect(serializeOoxmlPart(state(alice).part)).toBe(originalXml);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, PROPOSED);
        // Metadata stays independently editable while the source proposal is pending.
        expect(updateCodeLanguage(bob, { blockId: 'code', language: 'javascript' })).toMatchObject({
          success: true,
        });
        await pair.settle();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, PROPOSED);
        transaction(bob, (writer) => {
          const current = state(bob);
          const locks = codeSourceReviewLocks(current.pkg, ['code']);
          const ids = collectRevisionSites(current.part).map((site) => site.node.id);
          writer.edit(locks.before);
          writer.apply(current.part.name, {
            op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
            siteNodeIds: ids,
          });
          writer.edit(locks.after);
        });
        await pair.settle();
        const result = decision === 'accept' ? PROPOSED : ORIGINAL;
        verify(result, result);
        if (decision === 'reject') {
          expect(state(alice).committed.paragraphs.map((paragraph) => paragraph.id)).toContain(
            initial.committed.paragraphs[1]!.id
          );
          expect(state(alice).committed.paragraphs.filter(isJoinedParagraph)).toHaveLength(1);
          // Shared decisions retain source identities. Word export performs their physical join.
          const exported = openCode(new Uint8Array(await alice.save()));
          expect(serializeOoxmlPart(exported.parts.get(exported.mainDocumentPart)!)).toBe(
            originalXml
          );
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, PROPOSED);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(result, result);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        for (const editor of pair.editors) {
          const reopened = openCode(new Uint8Array(await editor.save()));
          const payload = buildBlockIndex(reopened).byId.get('code')!.payload;
          expect(payload.attributes).toEqual({ language: 'typescript', value: result });
        }
      } finally {
        pair.destroy();
      }
    }, 30_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const tail of [false, true])
    test(`late peer text survives source rejection, both histories, and later code authoring: ${clientIds}/${tail}`, async () => {
      const pair = await createTwoEditorFixture({
        document: await codeDocument(),
        clientIds,
        ownedAttributes: true,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        transaction(alice, (writer) =>
          stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
        );
        await pair.settle();
        const original = state(alice);
        const target = original.committed.paragraphs[tail ? 2 : 1]!;
        const offset = tail ? 1 : 2; // The first physical fragment also contains the proposed tab.
        pair.pause();
        transaction(alice, (writer) => {
          const current = state(alice);
          const locks = codeSourceReviewLocks(current.pkg, ['code']);
          const ids = collectRevisionSites(current.part).map((site) => site.node.id);
          writer.edit(locks.before);
          writer.apply(current.part.name, { op: 'rejectAllRevisions', siteNodeIds: ids });
          writer.edit(locks.after);
        });
        transaction(bob, (writer) => {
          const current = state(bob);
          const locks = codeSourceReviewLocks(current.pkg, ['code']);
          writer.edit(locks.before);
          writer.apply(current.part.name, {
            op: 'insertText',
            paragraphId: target.id,
            offset,
            text: 'PEER',
          });
          writer.edit(locks.after);
        });
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const joined = tail ? 'one\nmiddPEERle\nlast' : 'one\nmPEERiddle\nlast';
        const pending = tail ? 'one\n\tmid\ndPEERle\nlast' : 'one\n\tmPEERid\ndle\nlast';
        const verify = (committed: string, working: string) => {
          for (const editor of pair.editors) {
            expect(state(editor).committed.value).toBe(committed);
            expect(state(editor).working.value).toBe(working);
          }
        };
        verify(joined, joined);
        const projected = async (expected: string) => {
          for (const editor of pair.editors) {
            const revision = propertyToken(editor);
            const view = await readWorkingView(editor.surface!.session, revision);
            if (!view.success) throw Error(view.error);
            expect(view.view.projection.text).toContain(`\`\`\`python\n${expected}\n\`\`\``);
            expect(
              indexTargetSource({
                pkg: editor.surface!.session.currentPackage(),
                revision,
                projection: view.view.projection,
              })
            ).toMatchObject({ success: true });
          }
        };
        await projected(joined);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, pending);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, joined);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, ORIGINAL);
        await projected(ORIGINAL);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, joined);
        await projected(joined);
        const authored = joined.replace('PEER', 'USER\n\t');
        expect(
          updateCodeSource(bob, { blockId: 'code', expectedSource: joined, value: authored })
        ).toMatchObject({ success: true });
        await pair.settle();
        verify(authored, authored);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, joined);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(authored, authored);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        for (const editor of pair.editors) {
          const reopened = openCode(new Uint8Array(await editor.save()));
          expect(buildBlockIndex(reopened).byId.get('code')!.payload.attributes.value).toBe(
            authored
          );
        }
      } finally {
        pair.destroy();
      }
    }, 30_000);
