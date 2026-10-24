// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  findNode,
} from '@docx-editor.dev/core/store';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  manualContext,
  manualTrackingDocument,
  manualEdits,
  manualWave,
  manualRevisionFingerprints,
  manualToken,
  manualWorking,
  selectManual,
} from '../../test/support/manual-tracking';
import { setManualSuggestionMode, manualSuggestionState } from './manual-tracking';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`manual tracked edits survive concurrent AI recomposition, ${decision}, independent history, peer deletion and exact saved convergence ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: manualTrackingDocument(),
        documentId: manualContext.documentId,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const first = await manualWave(alice!, 'ONE');
        await pair.settle();
        manualEdits(alice!);
        await pair.settle();
        const manual = manualRevisionFingerprints(alice!);
        expect(manual).toHaveLength(3);
        const before = a.session.part(),
          id = a.session.paragraphIds()[1]!,
          offset = paragraphTextOf(before, id)!.length;
        selectManual(bob!, 1, offset);
        pair.pause();
        const second = await manualWave(alice!, 'TWO', first),
          authored = a.session.part();
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const withPeer = (part: typeof before) => {
          const r = applyTreeOp(part, { op: 'insertText', paragraphId: id, offset, text: '!' });
          if (!r.ok) throw new Error(r.reason);
          return canonicalOoxmlFingerprint(r.part);
        };
        const after = withPeer(authored);
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(after);
          expect(manualRevisionFingerprints(e, 'Alice')).toEqual(manual);
          const peer = applyTreeOp(before, {
            op: 'insertText',
            paragraphId: id,
            offset,
            text: '!',
          });
          if (!peer.ok) throw new Error(peer.reason);
          for (const paragraphId of a.session.paragraphIds().slice(1, 4))
            expect(
              canonicalOoxmlFingerprint(findNode(e.surface!.session.part(), paragraphId)!)
            ).toEqual(canonicalOoxmlFingerprint(findNode(peer.part, paragraphId)!));
        }
        expect(manualSuggestionState(alice!)).toMatchObject({ active: true, author: 'Alice' });
        expect(bob!.getConfiguredAuthor()).toBe('Bob');
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(withPeer(before));
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
        const selection = {
          anchor: { paragraphId: id, offset: 1 },
          head: { paragraphId: id, offset: 6 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const e of pair.editors) {
          expect(caretAt(e.surface!.layout(), selection.head)).not.toBeNull();
          expect(
            selectionRects(e.surface!.layout(), selection, e.surface!.session.paragraphIds()).length
          ).toBeGreaterThan(0);
        }
        const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
        expect(
          decide(alice!, second.ownership.at(-1)!.groupId, manualToken(alice!), manualToken(alice!))
            .success
        ).toBe(true);
        await pair.settle();
        const resolved = canonicalOoxmlFingerprint(a.session.part());
        expect(manualRevisionFingerprints(alice!)).toEqual(manual);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(resolved);
        expect(setManualSuggestionMode(alice!, false).ok).toBe(true);
        const beforeDeletion = a.session.part();
        selectManual(alice!, 1, 0, paragraphTextOf(beforeDeletion, id)!.length);
        selectManual(bob!, 1, 2);
        pair.pause();
        a.deleteForward();
        b.insertPlainText('?');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const recovered = applyTreeOp(beforeDeletion, {
          op: 'insertText',
          paragraphId: id,
          offset: 2,
          text: '?',
        });
        if (!recovered.ok) throw new Error(recovered.reason);
        const saved = [];
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(recovered.part)
          );
          expect(manualRevisionFingerprints(e, 'Alice')).toEqual(manual);
          const read = readOoxmlPackage(new Uint8Array(await e.save()));
          if (!read.ok) throw new Error(read.reason);
          expect(
            canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
          ).toEqual(canonicalOoxmlFingerprint(recovered.part));
          saved.push(semanticDigest(read.package.parts.values()));
          expect(paragraphTextOf(e.surface!.session.part(), id)).toBe('Ma?nual new text!');
          expect(await manualWorking(e)).toContain(String.raw`Ma?nual new text\!`);
        }
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });
