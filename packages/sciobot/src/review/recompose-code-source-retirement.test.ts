import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument } from '../../test/support/code';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { updateCodeLanguage } from '../blocks/code-edit';
import { indexTargetSource } from '../diff/target-source';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { rejectAllReviewGroups } from './reject-all';
import { acceptAllReviewGroups } from './accept-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const ORIGINAL = 'one\nmiddle\nlast';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`a new code proposal follows rejection undo and redo without exposing retired lines: ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await codeDocument(ORIGINAL),
      documentId: propertyContext.documentId,
      ownedAttributes: true,
      clientIds,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      let ownership: Parameters<typeof recomposeOwnedProposals>[1]['ownership'] = [];
      let contentProofs: Parameters<typeof recomposeOwnedProposals>[1]['contentProofs'] = [];
      const propose = async (source: string, proposalId: string) => {
        const view = await readWorkingView(alice.surface!.session, propertyToken(alice));
        if (!view.success) throw Error(view.error);
        const input = {
          pkg: alice.surface!.session.currentPackage(),
          revision: propertyToken(alice),
          projection: view.view.projection,
        };
        expect(indexTargetSource(input)).toMatchObject({ success: true });
        const result = await recomposeOwnedProposals(
          alice,
          {
            revision: propertyToken(alice),
            targetMarkdown: view.view.projection.text.replace(
              /```python[\s\S]*?```|```typescript[\s\S]*?```/,
              `\`\`\`typescript\n${source}\n\`\`\``
            ),
            ownership,
            contentProofs,
          },
          propertyContext,
          {
            payloads: {
              atomic: {
                proposalId,
                fallbackTargets: [{ blockId: 'code', before: ORIGINAL, after: source }],
              },
            },
          }
        );
        if (!result.success) throw Error(result.error);
        ownership = result.ownership;
        contentProofs = result.contentProofs;
        await pair.settle();
      };
      expect(updateCodeLanguage(bob, { blockId: 'code', language: 'typescript' }).success).toBe(
        true
      );
      await pair.settle();
      await propose(ORIGINAL + '\nfirst proposal', 'code_first');
      const replacement = ORIGINAL + '\nreplacement\n';
      await propose(replacement, 'code_replacement');
      expect(rejectAllReviewGroups(alice, propertyToken(alice), propertyContext).success).toBe(
        true
      );
      await pair.settle();
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await propose(replacement, 'code_after_rejection');
      expect(acceptAllReviewGroups(alice, propertyToken(alice), propertyContext).success).toBe(
        true
      );
      await pair.settle();
      for (const editor of pair.editors) {
        const committed = await readCommittedView(editor.surface!.session, propertyToken(editor));
        const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
        if (!committed.success || !working.success) throw Error('Missing source view');
        expect(committed.view.projection.text).toBe(working.view.projection.text);
        expect(working.view.projection.text).toContain(`\`\`\`typescript\n${replacement}\n\`\`\``);
        expect(
          indexTargetSource({
            pkg: editor.surface!.session.currentPackage(),
            revision: propertyToken(editor),
            projection: working.view.projection,
          })
        ).toMatchObject({ success: true });
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  }, 30_000);
