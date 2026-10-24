// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listNumbering, listParagraph, listPPr } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { replaceOwnedProposals } from '../review/replace-owned-proposals';
import { planParagraphText } from '../diff/paragraph-text';
import { readEditorRevision } from '../projection/revision-cache';
import { setManualList } from './lists';
const context = { documentId: 'list-sequence-peers', epoch: 'mount' };
const numbering = listNumbering.replace(
  '</w:numbering>',
  '<w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
);
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`rejecting a bridge keeps a peer's conflicting tail format ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture(
        listParagraph('One', '1') +
          listParagraph('Gap') +
          listParagraph('Three', '4') +
          listParagraph('Four', '4'),
        { numbering }
      ),
      documentId: context.documentId,
      clientIds: clients,
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      const ids = a.session.paragraphIds();
      const first = { paragraphId: ids[1]!, offset: 1 },
        last = { paragraphId: ids[3]!, offset: 1 };
      a.setSelection({ anchor: first, head: first });
      b.setSelection({ anchor: last, head: last });
      expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      pair.pause();
      expect(setManualList(alice!, 'decimal', context).success).toBe(true);
      expect(setManualList(bob!, 'circle', context).success).toBe(true);
      const expected = canonicalOoxmlFingerprint(b.session.part());

      pair.reconnect(clients[0] === 41 ? 0 : 1);
      await pair.settle();
      const groups = readReviewGroups(a.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      expect(groups.groups).toHaveLength(1);
      const result = rejectReviewGroup(
        alice!,
        groups.groups[0]!.id,
        readEditorRevision(alice!, context.documentId, context.epoch)!,
        context
      );
      if (!result.success) throw new Error(result.error);
      await pair.settle();
      for (const peer of pair.editors)
        expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const boundary of [false, true])
    for (const decision of ['direct', 'accept', 'reject'] as const)
      test(`sequence bridge preserves pending AI ownership, peer text and independent history: boundary=${boundary}/${decision}/${clients}`, async () => {
        const body =
          listParagraph('One', '1') +
          (boundary
            ? `<w:p><w:pPr><w:pStyle w:val="Heading1"/>${listPPr('5')}</w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>`
            : '') +
          listParagraph('Gap') +
          listParagraph('Three', '4') +
          listParagraph('Four', '4');
        const pair = await createTwoEditorFixture({
          document: docxFixture(body, { numbering }),
          documentId: context.documentId,
          clientIds: clients,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          const ids = a.session.paragraphIds(),
            gap = boundary ? 2 : 1,
            tail = ids.at(-1)!;
          const aiPlan = planParagraphText(
            {
              revision: readEditorRevision(alice!, context.documentId, context.epoch)!,
              partName: a.session.part().name,
              paragraphId: tail,
              text: 'Four',
            },
            'Four AI'
          );
          if (!aiPlan.success) throw new Error(aiPlan.error);
          const proposed = replaceOwnedProposals(alice!, [], aiPlan.plan, {
            ...context,
            author: 'Bot',
            date: '2026-10-24T12:00:00Z',
          });
          if (!proposed.success) throw new Error(proposed.error);
          await pair.settle();
          const existing = readReviewGroups(a.session.currentPackage());
          if (!existing.success) throw new Error(existing.error);
          expect(existing.groups).toHaveLength(1);
          const old = existing.groups[0]!;
          const initial = a.session.part(),
            before = canonicalOoxmlFingerprint(initial);
          const point = { paragraphId: ids[gap]!, offset: 1 };
          a.setSelection({ anchor: point, head: point });
          const peerPoint = { paragraphId: ids[gap + 1]!, offset: 5 };
          b.setSelection({ anchor: peerPoint, head: peerPoint });
          if (decision !== 'direct')
            expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          pair.pause();
          const changed = setManualList(alice!, 'decimal', context);
          if (!changed.success) throw new Error(changed.error);
          const authored = a.session.part();
          const expected = applyTreeOp(authored, {
            op: 'insertText',
            paragraphId: peerPoint.paragraphId,
            offset: 5,
            text: '!',
          });
          const recovered = applyTreeOp(initial, {
            op: 'insertText',
            paragraphId: peerPoint.paragraphId,
            offset: 5,
            text: '!',
          });
          if (!expected.ok || !recovered.ok)
            throw new Error('Independent peer text insertion refused.');
          b.insertPlainText('!');
          b.flushPendingInput();
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const after = canonicalOoxmlFingerprint(expected.part);
          for (const peer of pair.editors)
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(after);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
            canonicalOoxmlFingerprint(recovered.part)
          );
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(before);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
          const groups = readReviewGroups(a.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups.find((group) => group.id === old.id)).toEqual(old);
          const fresh = groups.groups.filter((group) => group.id !== old.id);
          expect(fresh).toHaveLength(decision === 'direct' ? 0 : 1);
          if (decision !== 'direct') {
            const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
            expect(
              decide(
                alice!,
                fresh[0]!.id,
                readEditorRevision(alice!, context.documentId, context.epoch)!,
                context
              ).success
            ).toBe(true);
            await pair.settle();
          }
          let word = recovered.part;
          if (decision !== 'reject') {
            for (let index = gap; index < ids.length; index++) {
              if (boundary && index !== gap) continue;
              const changed = applyTreeOp(word, {
                op: 'setListNumbering',
                paragraphId: ids[index]!,
                numId: boundary ? '4' : '1',
                level: 0,
              });
              if (!changed.ok) throw new Error(changed.reason);
              word = changed.part;
            }
          }
          for (const peer of pair.editors) {
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(
              canonicalOoxmlFingerprint(word)
            );
            expect(paragraphTextOf(peer.surface!.session.part(), peerPoint.paragraphId)).toBe(
              'Three!'
            );
            const remaining = readReviewGroups(peer.surface!.session.currentPackage());
            if (!remaining.success) throw new Error(remaining.error);
            expect(remaining.groups).toEqual([old]);
          }
          const saved = [];
          for (const peer of pair.editors) {
            const result = readOoxmlPackage(new Uint8Array(await peer.save()));
            if (!result.ok) throw new Error(result.reason);
            saved.push(semanticDigest(result.package.parts.values()));
          }
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          pair.destroy();
        }
      });
