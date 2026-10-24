// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyBoundaryBody, propertyBoundaryOps } from '../../test/support/property-boundaries';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { persistReviewGroups, readReviewGroups } from './groups';
import { freshPropertyGroupScopes } from './property-group-scopes';
import { captureReviewGroupHistory } from './group-history';
import { groupRemovalBudget } from './group-removal-budget';
import { groupReviewFieldBudget } from './group-field-budget';
import { rejectReviewGroup } from './reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import type { PropertyGroupScope } from './property-history-contract';
const context = { documentId: 'independent-property-groups', epoch: 'mount' };
function propose(editor: DocxEditorInstance) {
  const surface = editor.surface!,
    before = surface.session.currentPackage();
  const beforeGroups = readReviewGroups(before);
  if (!beforeGroups.success) throw new Error(beforeGroups.error);
  const existing = new Set(beforeGroups.groups.map((group) => group.id));
  let current: readonly PropertyGroupScope[] = [];
  const complete = (pkg: OoxmlPackage) => {
    const next = persistReviewGroups(before, pkg);
    const read = readReviewGroups(next);
    if (!read.success) throw new Error(read.error);
    current = freshPropertyGroupScopes(
      before,
      next,
      new Set(read.groups.filter((group) => !existing.has(group.id)).map((group) => group.id))
    );
    return next;
  };
  const ops = propertyBoundaryOps(surface.session.paragraphIds()[0]!);
  let preview = before,
    flagWrites = 0,
    fieldWrites = 0;
  for (const op of ops) {
    const applied = applyTreeOp(preview.parts.get(preview.mainDocumentPart)!, op);
    if (!applied.ok) throw new Error(applied.reason);
    const next = withPart(preview, applied.part);
    flagWrites += groupRemovalBudget(preview, next);
    fieldWrites += groupReviewFieldBudget(preview, next);
    preview = next;
  }
  const finished = complete(preview);
  flagWrites += groupRemovalBudget(preview, finished);
  fieldWrites += groupReviewFieldBudget(preview, finished);
  const predicted = current;
  expect(predicted).toHaveLength(2);
  current = [];
  let committed = false;
  const refusal = surface.commitReviewOps((historyToken) => {
    const applied = captureReviewGroupHistory(
      surface.collaborationSession(),
      () =>
        surface.session.applyTreeOps(
          ops,
          undefined,
          undefined,
          { kind: 'body' },
          {
            historyToken,
            packageEditsAfterOps: [complete],
          }
        ),
      flagWrites,
      fieldWrites,
      { preview: predicted, current: () => current }
    );
    committed = applied.committed;
    return applied;
  }, 'revision-resolve');
  if (refusal || !committed)
    throw new Error(String(refusal ?? 'Property publication did not commit.'));
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const author of ['Alice', 'Bot'])
    test(`independent paragraph and mark-format decisions retain break revisions, section state, peer formatting, and histories: ${clients}/${author}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(propertyBoundaryBody),
        documentId: context.documentId,
        clientIds: clients,
        ownedAttributes: true,
        experimentalPropertyReview: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const id = a.session.paragraphIds()[0]!;
        pair.pause();
        propose(alice!);
        const authored = a.session.part();
        const point = { paragraphId: id, offset: 2 };
        b.setSelection({ anchor: point, head: point });
        const peerOp = {
          op: 'setRunProperties' as const,
          paragraphId: id,
          start: 0,
          end: 4,
          properties: [{ localName: 'i', attributes: {} }],
        };
        expect(b.applyAutomationOps(() => [peerOp]).committed).toBe(true);
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const groups = readReviewGroups(a.session.currentPackage());
        if (!groups.success) throw new Error(groups.error);
        expect(groups.groups).toHaveLength(3);
        const selected = groups.groups.find((group) =>
          group.native.every((member) => member.address.author === author)
        )!;
        const siteIds = selected.native.flatMap((member) => member.siteNodeIds);
        const before = a.session.part(),
          pending = canonicalOoxmlFingerprint(before);
        const expected = applyTreeOp(before, { op: 'rejectAllRevisions', siteNodeIds: siteIds });
        if (!expected.ok) throw new Error(expected.reason);
        const independent = collectRevisionSites(before)
          .filter((site) => !siteIds.includes(site.node.id))
          .map((site) => ({ id: site.node.id, fingerprint: canonicalOoxmlFingerprint(site.node) }));
        const rejected = rejectReviewGroup(
          alice!,
          selected.id,
          readEditorRevision(alice!, context.documentId, context.epoch)!,
          context
        );
        if (!rejected.success) throw new Error(rejected.error);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(expected.part)
          );
          for (const site of independent)
            expect(
              canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), site.id)!)
            ).toEqual(site.fingerprint);
          const read = readReviewGroups(editor.surface!.session.currentPackage());
          if (!read.success) throw new Error(read.error);
          expect(read.groups.map((group) => group.id).sort()).toEqual(
            groups.groups
              .filter((group) => group.id !== selected.id)
              .map((group) => group.id)
              .sort()
          );
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(id);
        expect(remote?.head.offset).toBe(2);
        const ownCaret = caretAt(a.layout(), point),
          peerCaret = caretAt(b.layout(), point);
        expect(ownCaret).not.toBeNull();
        expect(peerCaret).not.toBeNull();
        expect({ x: ownCaret!.x, y: ownCaret!.y }).toEqual({ x: peerCaret!.x, y: peerCaret!.y });
        expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(
          0
        );
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(pending);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(expected.part)
          );
        const withoutPeer = applyTreeOp(authored, {
          op: 'rejectAllRevisions',
          siteNodeIds: collectRevisionSites(authored)
            .filter((site) =>
              site.node.attributes.some(
                (attribute) => attribute.localName === 'author' && attribute.value === author
              )
            )
            .map((site) => site.node.id),
        });
        if (!withoutPeer.ok) throw new Error(withoutPeer.reason);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(withoutPeer.part)
          );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
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
  test(`bound property rejection obeys a peer content lock without publishing: ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture(
        '<w:sdt><w:sdtPr><w:tag w:val="locked-target"/></w:sdtPr><w:sdtContent>' +
          propertyBoundaryBody +
          '</w:sdtContent></w:sdt>'
      ),
      documentId: context.documentId,
      clientIds: clients,
      ownedAttributes: true,
      experimentalPropertyReview: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      propose(alice!);
      await pair.settle();
      const body = b.session
        .part()
        .root.children.find((node) => node.kind !== 'textValue' && node.localName === 'body')!;
      if (body.kind === 'textValue') throw new Error('Missing body.');
      const control = body.children.find(
        (node) => node.kind !== 'textValue' && node.localName === 'sdt'
      )!;
      expect(
        b.applyAutomationOps(() => [
          { op: 'setContentControlProperties', controlId: control.id, lock: 'contentLocked' },
        ]).committed
      ).toBe(true);
      await pair.settle();
      const groups = readReviewGroups(a.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      const version = a.session.packageRevision();
      const native = canonicalOoxmlFingerprint(a.session.part());
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      let publications = 0;
      const published = () => publications++;
      pair.documents[0]!.on('update', published);
      try {
        for (const group of groups.groups.filter(
          (group) => group.native[0]?.address.author !== 'Peer'
        )) {
          const result = rejectReviewGroup(
            alice!,
            group.id,
            readEditorRevision(alice!, context.documentId, context.epoch)!,
            context
          );
          expect(result.success).toBe(false);
          if (result.success) throw new Error('A locked property group was rejected.');
          expect(result.error).toContain('locked');
        }
      } finally {
        pair.documents[0]!.off('update', published);
      }
      await pair.settle();
      expect(publications).toBe(0);
      expect(a.session.packageRevision()).toBe(version);
      expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(native);
      for (let index = 0; index < pair.documents.length; index++)
        expect(Y.encodeStateAsUpdate(pair.documents[index]!)).toEqual(shared[index]!);
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      const group = groups.groups.find((group) => group.native[0]?.address.author === 'Alice')!;
      const result = rejectReviewGroup(
        alice!,
        group.id,
        readEditorRevision(alice!, context.documentId, context.epoch)!,
        context
      );
      if (!result.success) throw new Error(result.error);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
