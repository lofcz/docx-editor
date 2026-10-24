// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { acceptAllReviewGroups } from './accept-all';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const causal of [false, true])
    test(`terminal tracked append ${clients}/${causal ? 'causal' : 'concurrent'} retains peer text and history`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture('<w:p><w:r><w:t>a</w:t></w:r></w:p>'),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const paragraphId = alice!.surface!.session.paragraphIds()[0]!;
        if (!causal) pair.pause();
        const staged = await recomposeOwnedProposals(
          alice!,
          { revision: propertyToken(alice!), targetMarkdown: 'an', ownership: [] },
          propertyContext
        );
        if (!staged.success) throw Error(staged.error);
        if (causal) await pair.settle();
        expect(
          bob!.surface!.applyAutomationOps(() => [
            {
              op: 'insertText',
              paragraphId,
              offset: causal ? 2 : 1,
              text: ' Peer',
            },
          ]).rejected
        ).toBe(false);
        if (!causal) pair.reconnect();
        await pair.settle();
        const check = async (expected: string) => {
          for (const editor of pair.editors) {
            const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
            expect(working.success && working.view.projection.text.trimEnd()).toBe(expected);
          }
        };
        await check('an Peer');
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check('an');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check('an Peer');
        const accepted = acceptAllReviewGroups(alice!, propertyToken(alice!), propertyContext);
        if (!accepted.success) throw Error(accepted.error);
        await pair.settle();
        const fingerprints = [];
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), paragraphId)).toBe('an Peer');
          const reopened = readOoxmlPackage(
            writeOoxmlPackage(editor.surface!.session.currentPackage())
          );
          if (!reopened.ok) throw Error(reopened.reason);
          fingerprints.push(
            canonicalOoxmlFingerprint(
              reopened.package.parts.get(reopened.package.mainDocumentPart)!
            )
          );
        }
        expect(fingerprints[0]).toBe(fingerprints[1]);
      } finally {
        pair.destroy();
      }
    });
