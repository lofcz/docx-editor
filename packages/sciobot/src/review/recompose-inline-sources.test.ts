// Collaboration transport is test-only under the EigenPal Pro License.
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlTextOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readEditorRevision } from '../projection/revision-cache';
import { recomposeOwnedProposals, type RecomposeOptions } from './recompose';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { inspectExternalReview } from './external-review-state';

const context = {
  documentId: 'inline-source-review',
  epoch: 'mounted',
  author: 'AI',
  date: '2026-10-26T12:00:00Z',
};
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (id: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${run('Book')}</w:sdtContent></w:sdt>`;
function fixture() {
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p>${run('Lead ')}${control('a')}${run(' between ')}${control('b')}${run(' tail')}</w:p>`
    )
  );
  if (!read.ok) throw Error(read.reason);
  let pkg = read.package;
  for (const id of ['a', 'b']) {
    const stored = withPayloadRecord(
      pkg,
      {
        version: 1,
        id,
        type: 'source',
        attributes: {
          title: 'Book',
          url: `https://source.test/${id}`,
          author: 'Writer',
          bookId: id === 'a' ? 123 : 456,
          trust: 'verified',
          sourceType: 'web',
        },
        opaqueExtension: { keep: [id, true] },
      },
      'Citation'
    );
    if (!stored.ok) throw Error(stored.reason);
    pkg = stored.pkg;
  }
  return writeOoxmlPackage(pkg);
}
const mount = (document = fixture()) =>
  createDocxEditor({
    container: window.document.createElement('div'),
    document,
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, context.documentId, context.epoch)!;
async function text(editor: DocxEditorInstance, committed = false) {
  const result = await (committed ? readCommittedView : readWorkingView)(
    editor.surface!.session,
    token(editor)
  );
  if (!result.success) throw Error(result.error);
  return result.view.projection.text;
}
const options = (id: string): RecomposeOptions => ({
  preparePayloads: (targets) => ({
    atomic: {
      proposalId: id,
      fallbackTargets: targets.map((target) => ({
        blockId: target.blockId,
        before: target.beforeText,
        after: String(target.after.attributes.title),
      })),
    },
  }),
});
const payloads = (editor: DocxEditorInstance) =>
  Object.fromEntries(
    [...buildBlockIndex(editor.surface!.session.currentPackage()).byId].map(([id, block]) => [
      id,
      block.payload,
    ])
  );
const value = (editor: DocxEditorInstance, id: string) =>
  contentControlTextOf(
    findNode(
      editor.surface!.session.part(),
      buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(id)!.controlId
    )!,
    '\n'
  );
for (const both of [false, true])
  for (const decision of ['accept', 'reject'] as const)
    test(`inline source waves retain neighboring atoms and same-paragraph manual text: both=${both}/${decision}`, async () => {
      const editor = mount();
      let reopened: DocxEditorInstance | undefined;
      try {
        const initial = await text(editor),
          original = payloads(editor);
        const firstTarget = initial.replace('title="Book"', 'title="First"');
        const first = await recomposeOwnedProposals(
          editor,
          {
            revision: token(editor),
            targetMarkdown: firstTarget,
            ownership: [],
            contentProofs: [],
          },
          context,
          options('source-one')
        );
        if (!first.success) throw Error(first.error);
        expect(await text(editor)).toBe(firstTarget);
        expect(await text(editor, true)).toBe(initial);
        expect(payloads(editor)).toEqual(original);
        expect(value(editor, 'a')).toBe('First');
        expect(value(editor, 'b')).toBe('Book');
        const paragraphId = editor.surface!.session.paragraphIds()[0]!;
        const end = paragraphTextOf(editor.surface!.session.part(), paragraphId)!.length;
        expect(
          editor.surface!.applyAutomationOps(
            () => [{ op: 'insertText', paragraphId, offset: end, text: ' peer' }],
            { kind: 'body' },
            [],
            false,
            true
          ).rejected
        ).toBe(false);
        const secondTarget = (firstTarget + ' peer')
          .replace('title="First"', 'title="Final"')
          .replace(both ? 'title="Book"' : 'unmatched-token', 'title="Second"');
        const second = await recomposeOwnedProposals(
          editor,
          {
            revision: token(editor),
            targetMarkdown: secondTarget,
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          context,
          options('source-two')
        );
        if (!second.success) throw Error(second.error);
        expect(await text(editor)).toBe(secondTarget);
        expect(await text(editor, true)).toBe(initial + ' peer');
        expect(payloads(editor)).toEqual(original);
        expect(value(editor, 'a')).toBe('Final');
        expect(value(editor, 'b')).toBe(both ? 'Second' : 'Book');
        const groups = readReviewGroups(editor.surface!.session.currentPackage());
        if (!groups.success) throw Error(groups.error);
        expect(groups.groups.every((group) => !group.readOnly)).toBe(true);
        expect(inspectExternalReview(editor.surface!.session.currentPackage())).toMatchObject({
          success: true,
          report: { hasConflicts: false },
        });
        const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(await text(editor)).toBe(firstTarget + ' peer');
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
        reopened = mount(new Uint8Array(await editor.save()));
        expect(await text(reopened)).toBe(secondTarget);
        const result = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          editor,
          token(editor),
          context
        );
        if (!result.success) throw Error(result.error);
        expect(await text(editor)).toBe(decision === 'accept' ? secondTarget : initial + ' peer');
        const expected = structuredClone(original);
        if (decision === 'accept') {
          expected.a = {
            ...expected.a!,
            attributes: { ...expected.a!.attributes, title: 'Final' },
          };
          if (both)
            expected.b = {
              ...expected.b!,
              attributes: { ...expected.b!.attributes, title: 'Second' },
            };
        }
        expect(payloads(editor)).toEqual(expected);
        expect(paragraphTextOf(editor.surface!.session.part(), paragraphId)).toBe(
          `Lead ${decision === 'accept' ? 'Final' : 'Book'} between ${decision === 'accept' && both ? 'Second' : 'Book'} tail peer`
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(await text(editor)).toBe(secondTarget);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(payloads(editor)).toEqual(expected);
      } finally {
        reopened?.destroy();
        editor.destroy();
      }
    });

for (const metadataOnly of [false, true])
  test(`independent source groups in one paragraph retain separate ownership: metadata=${metadataOnly}`, async () => {
    const editor = mount();
    try {
      const original = await text(editor);
      const a = await recomposeOwnedProposals(
        editor,
        {
          revision: token(editor),
          targetMarkdown: original.replace('title="Book"', 'title="First"'),
          ownership: [],
          contentProofs: [],
        },
        context,
        options('independent-a')
      );
      if (!a.success) throw Error(a.error);
      const afterA = await text(editor);
      const target = metadataOnly
        ? afterA.replace('https://source.test/b', 'https://source.test/new-b')
        : afterA.replace('title="Book"', 'title="Second"');
      const b = await recomposeOwnedProposals(
        editor,
        { revision: token(editor), targetMarkdown: target, ownership: [], contentProofs: [] },
        context,
        options('independent-b')
      );
      if (!b.success) throw Error(b.error);
      expect(await text(editor)).toBe(target);
      expect(await text(editor, true)).toBe(original);
      expect(readReviewGroups(editor.surface!.session.currentPackage())).toMatchObject({
        success: true,
        groups: [{ readOnly: false }, { readOnly: false }],
      });
      expect(inspectExternalReview(editor.surface!.session.currentPackage())).toMatchObject({
        success: true,
        report: { hasConflicts: false },
      });
      const { decideReviewGroup } = await import('./decide-group');
      const groups = readReviewGroups(editor.surface!.session.currentPackage());
      if (!groups.success) throw Error(groups.error);
      const groupA = groups.groups.find((group) =>
        group.properties.some((member) => member.blockId === 'a')
      )!;
      const accepted = decideReviewGroup(editor, groupA.id, token(editor), context, 'accept');
      if (!accepted.success) throw Error(accepted.error);
      expect(await text(editor)).toBe(target);
      expect(await text(editor, true)).toBe(afterA);
      const rejected = rejectAllReviewGroups(editor, token(editor), context);
      if (!rejected.success) throw Error(rejected.error);
      expect(await text(editor)).toBe(afterA);
      expect(value(editor, 'a')).toBe('First');
      expect(value(editor, 'b')).toBe('Book');
    } finally {
      editor.destroy();
    }
  });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`shared inline source waves retain peer text, histories, geometry and cold room: ${clients}/${receiver}/${decision}`, async () => {
        const pair = await createTwoEditorFixture({
          document: fixture(),
          documentId: context.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        let cold: Awaited<ReturnType<typeof createTwoEditorFixture>> | undefined;
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const initial = await text(alice),
            original = payloads(alice);
          const firstTarget = initial.replace('title="Book"', 'title="First"');
          const first = await recomposeOwnedProposals(
            alice,
            {
              revision: token(alice),
              targetMarkdown: firstTarget,
              ownership: [],
              contentProofs: [],
            },
            context,
            options('shared-one')
          );
          if (!first.success) throw Error(first.error);
          await pair.settle();
          pair.pause();
          const id = bob.surface!.session.paragraphIds()[0]!;
          const end = paragraphTextOf(bob.surface!.session.part(), id)!.length;
          expect(
            bob.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId: id, offset: end, text: ' peer' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          const target = firstTarget
            .replace('title="First"', 'title="Final"')
            .replace('title="Book"', 'title="Second"');
          const second = await recomposeOwnedProposals(
            alice,
            {
              revision: token(alice),
              targetMarkdown: target,
              ownership: first.ownership,
              contentProofs: first.contentProofs,
            },
            context,
            options('shared-two')
          );
          if (!second.success) throw Error(second.error);
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(await text(editor)).toBe(target + ' peer');
            expect(await text(editor, true)).toBe(initial + ' peer');
            expect(payloads(editor)).toEqual(original);
          }
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await text(alice)).toBe(target);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await text(alice)).toBe(target + ' peer');
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await text(alice)).toBe(firstTarget + ' peer');
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await text(alice)).toBe(target + ' peer');
          const point = { paragraphId: id, offset: 2 };
          bob.surface!.setSelection({ anchor: point, head: point });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(id);
          expect(remote?.head.offset).toBe(2);
          const aCaret = caretAt(alice.surface!.layout(), point),
            bCaret = caretAt(bob.surface!.layout(), point);
          expect(aCaret).not.toBeNull();
          expect(bCaret).not.toBeNull();
          expect({ x: aCaret!.x, y: aCaret!.y, height: aCaret!.height }).toEqual({
            x: bCaret!.x,
            y: bCaret!.y,
            height: bCaret!.height,
          });
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          cold = await createTwoEditorFixture({
            document: fixture(),
            documentId: context.documentId,
            ownedAttributes: true,
            clientIds: [101, 303],
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          for (const editor of cold.editors) {
            expect(await text(editor)).toBe(target + ' peer');
            expect(await text(editor, true)).toBe(initial + ' peer');
            expect(inspectExternalReview(editor.surface!.session.currentPackage())).toMatchObject({
              success: true,
              report: { hasConflicts: false },
            });
          }
          const result = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            alice,
            token(alice),
            context
          );
          if (!result.success) throw Error(result.error);
          await pair.settle();
          const expected = decision === 'accept' ? target + ' peer' : initial + ' peer';
          for (const editor of pair.editors) expect(await text(editor)).toBe(expected);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await text(bob)).toBe(target + ' peer');
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(await text(editor)).toBe(expected);
            expect(value(editor, 'a')).toBe(decision === 'accept' ? 'Final' : 'Book');
            expect(value(editor, 'b')).toBe(decision === 'accept' ? 'Second' : 'Book');
          }
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          cold?.destroy();
          pair.destroy();
        }
      });
