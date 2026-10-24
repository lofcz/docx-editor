// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { applyTreeOp, canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { listNumbering, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { replaceOwnedProposals } from '../review/replace-owned-proposals';
import { planParagraphText } from '../diff/paragraph-text';
import { readEditorRevision } from '../projection/revision-cache';
import {
  PROPERTY_GROUP_RECEIPTS,
  decodePropertyGroupBinding,
} from '../collaboration/property-group-codec';
import {
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  decodeFieldTarget,
  decodeFieldIntent,
} from '../collaboration/review-field-codec';
import { manualListState, setManualList } from './lists';
const context = { documentId: 'bound-property-group', epoch: 'mount' };
const numbering = listNumbering.replace(
  '</w:numbering>',
  '<w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
);
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const reopen of [false, true])
    for (const observed of [false, true])
      test(`bound group rejection preserves offline peer format and pending AI: ${clients}/cold=${reopen}/observed=${observed}`, async () => {
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
          experimentalPropertyReview: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        let cold: Awaited<ReturnType<typeof createTwoEditorFixture>> | undefined;
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          const ids = a.session.paragraphIds();
          const plan = planParagraphText(
            {
              revision: readEditorRevision(alice!, context.documentId, context.epoch)!,
              partName: a.session.part().name,
              paragraphId: ids[3]!,
              text: 'Four',
            },
            'Four AI'
          );
          if (!plan.success) throw new Error(plan.error);
          const prior = replaceOwnedProposals(alice!, [], plan.plan, {
            ...context,
            author: 'Bot',
            date: '2026-10-24T12:00:00Z',
          });
          if (!prior.success) throw new Error(prior.error);
          await pair.settle();
          const oldGroups = readReviewGroups(a.session.currentPackage());
          if (!oldGroups.success) throw new Error(oldGroups.error);
          expect(oldGroups.groups).toHaveLength(1);
          const old = oldGroups.groups[0]!;
          const originalPart = a.session.part();
          const original = canonicalOoxmlFingerprint(originalPart);
          const first = { paragraphId: ids[1]!, offset: 1 },
            last = { paragraphId: ids[3]!, offset: 1 };
          a.setSelection({ anchor: first, head: first });
          b.setSelection({ anchor: last, head: last });
          expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          pair.pause();
          let publications = 0;
          let boundWithNative = false;
          const count = (_update: Uint8Array, origin: unknown) => {
            if ((origin as { kind?: string } | null)?.kind === 'docx-document-local')
              publications++;
          };
          const observe = (transaction: Y.Transaction) => {
            const changed = new Set<unknown>(transaction.changed.keys());
            if (changed.has(pair.documents[0]!.getMap(PROPERTY_GROUP_RECEIPTS)))
              boundWithNative = changed.has(pair.documents[0]!.getMap(REVIEW_FIELD_INTENTS));
          };
          pair.documents[0]!.on('update', count);
          pair.documents[0]!.on('afterTransaction', observe);
          try {
            const changed = setManualList(alice!, 'decimal', context);
            if (!changed.success) throw new Error(changed.error);
          } finally {
            pair.documents[0]!.off('update', count);
            pair.documents[0]!.off('afterTransaction', observe);
          }
          expect(publications).toBe(1);
          expect(boundWithNative).toBe(true);
          const map = pair.documents[0]!.getMap(PROPERTY_GROUP_RECEIPTS);
          expect(map.size).toBe(1);
          const [key, raw] = [...map][0]!;
          const binding = decodePropertyGroupBinding(key, raw);
          const proposedTree = canonicalOoxmlFingerprint(a.session.part());
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(original);
          expect(map.has(key)).toBe(false);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(proposedTree);
          expect(map.get(key)).toBe(raw);
          // This checks actual native scalar selection, not just child-edge cancellation.
          expect(
            binding.fields.some((id) => {
              const intent = decodeFieldIntent(
                id,
                pair.documents[0]!.getMap(REVIEW_FIELD_INTENTS).get(id)
              );
              const target = decodeFieldTarget(
                intent.target,
                pair.documents[0]!.getMap(REVIEW_FIELD_TARGETS).get(intent.target)
              );
              return target.root === 'docx-package-attributes-v1';
            })
          ).toBe(true);
          if (observed) {
            pair.reconnect();
            await pair.settle();
            pair.pause();
          }
          const changed = setManualList(bob!, 'circle', context);
          if (!changed.success) throw new Error(changed.error);
          const membership = manualListState(bob!, context, ids[3]!)!.numbering!;
          const word = applyTreeOp(originalPart, {
            op: 'setListNumbering',
            paragraphId: ids[3]!,
            numId: membership.numId,
            level: membership.level,
          });
          if (!word.ok) throw new Error(word.reason);
          const expected = canonicalOoxmlFingerprint(word.part);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const pending = canonicalOoxmlFingerprint(a.session.part());
          let active = pair;
          if (reopen) {
            cold = await createTwoEditorFixture({
              document: docxFixture(listParagraph('Unused')),
              documentId: context.documentId,
              ownedAttributes: true,
              experimentalPropertyReview: true,
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            });
            active = cold;
            await cold.settle();
          }
          const reviewer = active.editors[0]!;
          const groups = readReviewGroups(reviewer.surface!.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(2);
          const group = groups.groups.find((group) => group.id !== old.id)!;
          expect(group.id).toBe(binding.groupId);
          const decided = rejectReviewGroup(
            reviewer,
            group.id,
            readEditorRevision(reviewer, context.documentId, context.epoch)!,
            context
          );
          if (!decided.success) throw new Error(decided.error);
          await active.settle();
          for (const peer of active.editors) {
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
            const remaining = readReviewGroups(peer.surface!.session.currentPackage());
            if (!remaining.success) throw new Error(remaining.error);
            expect(remaining.groups).toEqual([old]);
          }
          expect(reviewer.exec({ type: 'undo' }).ok).toBe(true);
          await active.settle();
          for (const peer of active.editors)
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(pending);
          expect(reviewer.exec({ type: 'redo' }).ok).toBe(true);
          await active.settle();
          for (const peer of active.editors)
            expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
          if (!reopen) {
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const peer of pair.editors)
              expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(original);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const peer of pair.editors)
              expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toEqual(expected);
          }
          const saved = await active.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          cold?.destroy();
          pair.destroy();
        }
      });
