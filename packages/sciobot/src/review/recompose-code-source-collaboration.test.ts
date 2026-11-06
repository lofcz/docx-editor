import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument, openCode } from '../../test/support/code';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { updateCodeLanguage } from '../blocks/code-edit';
import { buildBlockIndex } from '../blocks/block-index';
import { nativeCodeSource } from '../blocks/code-native-source';
import { removeBlock } from '../blocks/remove';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { codeSourceReviewLocks } from './code-source-review';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';

async function views(editor: DocxEditorInstance) {
  const before = await readCommittedView(editor.surface!.session, propertyToken(editor));
  const after = await readWorkingView(editor.surface!.session, propertyToken(editor));
  if (!before.success || !after.success) throw Error('Missing code review views.');
  return [before.view.projection.text, after.view.projection.text];
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`owned native code replacement preserves peer source, metadata, history, and reopen: ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await codeDocument(),
      documentId: propertyContext.documentId,
      ownedAttributes: true,
      clientIds,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const initial = (await views(alice))[1]!;
      const options = (proposalId: string, after: string) => ({
        payloads: {
          atomic: {
            proposalId,
            fallbackTargets: [{ blockId: 'code', before: 'one\nmiddle\nlast', after }],
          },
        },
      });
      const first = await recomposeOwnedProposals(
        alice,
        {
          revision: propertyToken(alice),
          targetMarkdown: initial.replace('middle', '\tmid\ndle'),
          ownership: [],
          contentProofs: [],
        },
        propertyContext,
        options('code_first', 'one\n\tmid\ndle\nlast')
      );
      if (!first.success) throw Error(first.error);
      await pair.settle();
      expect(first.committed).toBe(true);
      expect(first.ownership).toHaveLength(1);
      expect(first.ownership[0]!.properties).toEqual([
        { proposalId: 'code_first', blockId: 'code' },
      ]);
      expect(await views(alice)).toEqual([initial, initial.replace('middle', '\tmid\ndle')]);
      expect(updateCodeLanguage(bob, { blockId: 'code', language: 'typescript' })).toMatchObject({
        success: true,
      });
      await pair.settle();
      const target = initial
        .replace('```python', '```typescript')
        .replace('middle', 'second\nline');
      pair.pause();
      const sourcePkg = bob.surface!.session.currentPackage();
      const sourceBlock = buildBlockIndex(sourcePkg).byId.get('code')!;
      const sourcePart = sourcePkg.parts.get(sourceBlock.partName)!;
      const native = nativeCodeSource(
        sourcePart,
        findNode(sourcePart, sourceBlock.controlId)!,
        'committed'
      )!;
      const last = native.paragraphs.at(-1)!;
      const locks = codeSourceReviewLocks(sourcePkg, ['code']);
      expect(
        bob.surface!.applyAutomationOps(
          () => [],
          { kind: 'body' },
          undefined,
          false,
          true,
          'before',
          (ctx) => {
            if (
              !ctx.applyPackage(locks.before) ||
              !ctx.applyTo(sourcePart.name, {
                op: 'insertText',
                paragraphId: last.id,
                offset: 4,
                text: ' PEER',
                bias: 'right',
              }) ||
              !ctx.applyPackage(locks.after)
            )
              throw Error('Peer code text failed.');
          }
        )
      ).toMatchObject({ committed: true, rejected: false });
      const second = await recomposeOwnedProposals(
        alice,
        {
          revision: propertyToken(alice),
          targetMarkdown: target,
          ownership: first.ownership,
          contentProofs: first.contentProofs,
        },
        propertyContext,
        options('code_second', 'one\nsecond\nline\nlast')
      );
      if (!second.success) throw Error(second.error);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      const original = initial.replace('```python', '```typescript').replace('last', 'last PEER');
      const proposed = target.replace('last', 'last PEER');
      const verify = async (before: string, after: string) => {
        for (const editor of pair.editors) {
          expect(await views(editor)).toEqual([before, after]);
          expect(
            buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('code')!.payload
              .attributes.language
          ).toBe('typescript');
        }
      };
      await verify(original, proposed);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      await verify(original, original.replace('middle', '\tmid\ndle'));
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await verify(original, proposed);
      for (const decision of ['accept', 'reject'] as const) {
        const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          bob,
          propertyToken(bob),
          propertyContext
        );
        if (!decided.success) throw Error(decided.error);
        await pair.settle();
        const expected = decision === 'accept' ? proposed : original;
        await verify(expected, expected);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        for (const editor of pair.editors)
          expect(
            buildBlockIndex(openCode(new Uint8Array(await editor.save()))).byId.get('code')!.payload
              .attributes.value
          ).toBe(decision === 'accept' ? 'one\nsecond\nline\nlast PEER' : 'one\nmiddle\nlast PEER');
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await verify(original, proposed);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await verify(expected, expected);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await verify(original, proposed);
      }
      const reopened = await createTwoEditorFixture({
        document: new Uint8Array(await alice.save()),
        documentId: propertyContext.documentId,
        roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        for (const editor of reopened.editors)
          expect(await views(editor)).toEqual([original, proposed]);
      } finally {
        reopened.destroy();
      }
      const beforeDeletion = propertyToken(alice);
      const controlId = buildBlockIndex(bob.surface!.session.currentPackage()).byId.get(
        'code'
      )!.controlId;
      expect(removeBlock(bob, { controlId, mode: 'delete' })).toMatchObject({ success: true });
      await pair.settle();
      for (const editor of pair.editors)
        expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.has('code')).toBe(
          false
        );
      const deleted = alice.surface!.session.currentPackage();
      expect(
        await recomposeOwnedProposals(
          alice,
          {
            revision: beforeDeletion,
            targetMarkdown: proposed,
            ownership: second.ownership,
            contentProofs: second.contentProofs,
          },
          propertyContext,
          options('code_deleted', 'one\nsecond\nline\nlast PEER')
        )
      ).toMatchObject({ success: false, code: 'stale-revision' });
      expect(alice.surface!.session.currentPackage()).toBe(deleted);
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      await verify(original, proposed);
      pair.ownership.forEach((handle) => handle.assertHealthy());
    } finally {
      pair.destroy();
    }
  }, 30_000);
