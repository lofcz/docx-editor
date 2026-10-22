// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { findNode, paragraphOffsetIndex, type OoxmlNode } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  document_,
  context,
  plan,
  projection,
  ast,
  paragraph,
} from '../../test/support/link-structure-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';
const url = 'https://example.org/new';
for (const imported of [false, true])
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`tracked links keep two peer-owned text fields through acceptance ${clientIds}/${receiver}${imported ? ' with imported metadata' : ''}`, async () => {
        const pair = await createTwoEditorFixture({
          document: document_(
            imported
              ? paragraph(
                  '<w:r xmlns:keep="urn:keep" keep:run="original"><w:rPr/><w:t keep:text="first">Alpha </w:t><w:t keep:text="second">end</w:t></w:r>'
                )
              : undefined
          ),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          clientIds,
        });
        try {
          const [alice, bob] = pair.editors;
          const result = trackHyperlinkPlan(
            alice!,
            await plan(alice!, `[Alpha new end](${url})\n\nTail`),
            context
          );
          if (!result.success) throw Error(result.error);
          await pair.settle();
          const part = bob!.surface!.session.part(),
            id = bob!.surface!.session.paragraphIds()[0]!,
            p = findNode(part, id)!;
          if (p.kind !== 'paragraph') throw Error('paragraph');
          const link = p.children.find((n) => n.kind === 'hyperlink')!;
          const span = paragraphOffsetIndex(p).spanOf(link)!;
          pair.pause();
          expect(
            bob!.surface!.applyAutomationOps(
              () => [
                { op: 'insertText', paragraphId: id, offset: span.end - 1, text: 'P2' },
                { op: 'insertText', paragraphId: id, offset: span.start + 1, text: 'P1' },
              ],
              { kind: 'body' },
              [],
              false,
              true
            )
          ).toMatchObject({ committed: true, rejected: false });
          pair.reconnect(receiver);
          await pair.settle();
          const fields: OoxmlNode[] = [];
          const walk = (n: OoxmlNode) => {
            if (n.kind === 'textValue') {
              if (n.value.includes('P1') || n.value.includes('P2')) fields.push(n);
            } else n.children.forEach(walk);
          };
          walk(bob!.surface!.session.part().root);
          expect(fields).toHaveLength(2);
          expect(fields[0]!.id).not.toBe(fields[1]!.id);
          const check = async (text: string, original = text) => {
            for (const e of pair.editors) {
              expect(ast((await projection(e)).text)).toBe(ast(text + '\n\nTail'));
              expect(ast((await projection(e, true)).text)).toBe(ast(original + '\n\nTail'));
            }
          };
          const plain = `[Alpha new end](${url})`,
            peer = `[AP1lpha new enP2d](${url})`;
          await check(peer, 'Alpha end');
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(plain, 'Alpha end');
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(peer, 'Alpha end');
          expect(alice!.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
          await pair.settle();
          await check(peer);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(plain);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(peer);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(peer, 'Alpha end');
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(peer);
        } finally {
          pair.destroy();
        }
      }, 30000);
