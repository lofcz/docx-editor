import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { docxFixture } from '../../test/support/document';
import { groups, groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { rejectReviewGroup } from './reject-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const options = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function fingerprints(pkg: OoxmlPackage, parts: readonly string[]) {
  return parts.map((name) => [name, canonicalOoxmlFingerprint(pkg.parts.get(name)!)]);
}
for (const mode of ['body-header', 'imported-middle'] as const)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`${mode} rejection preserves native neighbors and original pending identity through both actors history; ${clients}/${receiver}`, async () => {
        const document =
          mode === 'body-header'
            ? writeOoxmlPackage(crossStoryGroupFixture())
            : docxFixture(
                ['Before', 'Middle', 'After']
                  .map(
                    (author, index) =>
                      `<w:p><w:r><w:t>${author} </w:t></w:r><w:ins w:id="${index + 1}" w:author="${author}"><w:r><w:t>Pending</w:t></w:r></w:ins></w:p>`
                  )
                  .join('')
              );
        const pair = await createTwoEditorFixture({
          document,
          clientIds: clients,
          ownedAttributes: true,
          editor: options,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const before = groups(alice),
            ids = before.map((group) => group.id);
          expect(before).toHaveLength(mode === 'body-header' ? 2 : 3);
          const selected =
            mode === 'body-header'
              ? before.find((group) => group.id === 'g1:stored:rg_cross_story')!
              : before[1]!;
          expect(selected.readOnly).toBe(false);
          expect(selected.native).toHaveLength(mode === 'body-header' ? 2 : 1);
          if (mode === 'imported-middle') expect(selected.id.startsWith('g1:stored:')).toBe(false);
          const remaining = before.filter((group) => group.id !== selected.id),
            parts =
              mode === 'body-header'
                ? ['/word/document.xml', '/word/header1.xml']
                : ['/word/document.xml'];
          const paragraph = selected.anchor!.paragraphId!,
            old = mode === 'body-header' ? 'Body ' : 'Middle ';
          const base = source(alice),
            text = applyTreeOp(base.parts.get(base.mainDocumentPart)!, {
              op: 'insertText',
              paragraphId: paragraph,
              offset: 0,
              text: 'Peer: ',
            });
          if (!text.ok) throw new Error(text.reason);
          const pending = withPart(base, text.part);
          let expected = pending;
          const address =
            mode === 'body-header' ? { id: '7', author: 'Teacher' } : { id: '2', author: 'Middle' };
          for (const name of parts) {
            const actual = applyTreeOp(expected.parts.get(name)!, {
              op: 'rejectRevision',
              revision: address,
            });
            if (!actual.ok) throw new Error(actual.reason);
            expected = withPart(expected, actual.part);
          }
          const pendingPrint = fingerprints(pending, parts),
            resolvedPrint = fingerprints(expected, parts);
          pair.pause();
          expect(
            rejectReviewGroup(alice, selected.id, groupToken(alice), propertyContext).success
          ).toBe(true);
          bob.surface!.setSelection({
            anchor: { paragraphId: paragraph, offset: 0 },
            head: { paragraphId: paragraph, offset: 0 },
          });
          bob.surface!.insertPlainText('Peer: ');
          bob.surface!.flushPendingInput();
          expect(
            rejectReviewGroup(bob, selected.id, groupToken(bob), propertyContext).success
          ).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          const state = (isPending: boolean) => {
            for (const editor of pair.editors) {
              const current = groups(editor);
              expect(current.map((group) => group.id)).toEqual(
                isPending ? ids : remaining.map((group) => group.id)
              );
              expect(fingerprints(source(editor), parts)).toEqual(
                isPending ? pendingPrint : resolvedPrint
              );
              expect(paragraphTextOf(editor.surface!.session.part(), paragraph)).toBe(
                `Peer: ${old}${isPending ? 'Pending' : ''}`
              );
              expect(
                current.filter((group) => group.id !== selected.id).map((group) => group.native)
              ).toEqual(remaining.map((group) => group.native));
            }
          };
          state(false);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          state(false);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          state(true);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          state(false);
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          for (const editor of pair.editors) {
            const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!saved.ok) throw new Error(saved.reason);
            expect(fingerprints(saved.package, parts)).toEqual(resolvedPrint);
            expect(
              collectRevisionSites(saved.package.parts.get(saved.package.mainDocumentPart)!)
            ).toHaveLength(mode === 'body-header' ? 1 : 2);
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60000);
