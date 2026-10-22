import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { groupToken } from '../../test/support/review-groups';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { captureReviewGroupHistory } from './group-history';
import { groupRemovalBudget } from './group-removal-budget';
import { groupReviewFieldBudget } from './group-field-budget';
import { persistReviewGroups, readReviewGroups } from './groups';
import { withInitializedReviewGroupStore, withReviewGroupRecord } from './group-record-store';
import { withReviewGroupRevisionReservations } from './group-allocation';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

const documentId = 'sciobot-parity';
const date = '2026-10-21T09:00:00.000Z';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
function bytes() {
  const loaded = readOoxmlPackage(
    docxFixture(
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
        '<w:ins w:id="17" w:author="Bot"><w:r><w:t>old</w:t></w:r></w:ins>' +
        '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Foreign</w:t></w:r><w:ins w:id="22" w:author="Peer"><w:r><w:t xml:space="preserve"> pending</w:t></w:r></w:ins></w:p>' +
        p('Tail')
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const initialized = withInitializedReviewGroupStore(loaded.package);
  if (!initialized.ok) throw new Error(initialized.reason);
  const retained = withReviewGroupRecord(initialized.pkg, {
    schemaVersion: 1,
    id: 'rg_stage_old',
    members: [{ partName: '/word/document.xml', revision: { id: '17', author: 'Bot' } }],
  });
  if (!retained.ok) throw new Error(retained.reason);
  return writeOoxmlPackage(retained.pkg);
}
const options = () => ({
  document: bytes(),
  documentId,
  ownedAttributes: true,
  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
const pkg = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();

/** R53 native-stage proof only. R25 still owns the validated application replacement API. */
function stage(editor: DocxEditorInstance, fail = false) {
  const surface = editor.surface!,
    session = surface.session,
    before = pkg(editor),
    paragraphId = session.paragraphIds()[0]!,
    sites = collectRevisionSites(session.part()).filter((site) =>
      site.node.attributes.some((a) => a.localName === 'id' && a.value === '17')
    );
  expect(sites).toHaveLength(1);
  const inverse: TreeDocOp = {
    op: 'rejectAllRevisions',
    siteNodeIds: sites.map((site) => site.node.id),
  };
  const insert: TreeDocOp = {
    op: 'insertText',
    paragraphId,
    offset: 6,
    text: 'new',
    revision: { author: 'Bot', date },
  };
  const complete = (working: OoxmlPackage) => persistReviewGroups(before, working);
  const scope = surface.collaborationSession();
  return withReviewGroupRevisionReservations(before, () => {
    let committed = false;
    const refusal = surface.commitReviewOps((historyToken) => {
      let preview = before,
        removals = 0,
        fields = 0;
      for (const op of [inverse, insert]) {
        const applied = applyTreeOp(preview.parts.get(before.mainDocumentPart)!, op);
        if (!applied.ok) throw new Error(applied.reason);
        const next = withPart(preview, applied.part);
        removals += groupRemovalBudget(preview, next);
        fields += groupReviewFieldBudget(preview, next);
        preview = next;
      }
      const next = complete(preview);
      removals += groupRemovalBudget(preview, next);
      fields += groupReviewFieldBudget(preview, next);
      const result = captureReviewGroupHistory(
        scope,
        () =>
          session.applyTreeOps(
            [inverse],
            undefined,
            undefined,
            { kind: 'body' },
            {
              historyToken,
              stage: (ctx) => {
                expect(
                  paragraphTextOf(
                    ctx.currentPackage().parts.get(before.mainDocumentPart)!,
                    paragraphId
                  )
                ).toStartWith('Alpha  end');
                expect(ctx.apply(insert)).toBe(true);
                if (fail) throw new Error('Injected dependent stage failure.');
              },
              packageEditsAfterOps: [complete],
            }
          ),
        removals,
        fields
      );
      committed = result.committed;
      return result;
    }, 'revision-resolve');
    if (refusal || !committed) throw new Error(String(refusal));
  });
}

function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const offset = paragraphTextOf(editor.surface!.session.part(), paragraphId)!.length;
  const result = editor.surface!.applyAutomationOps(() => [
    { op: 'insertText', paragraphId, offset, text, bias: 'right' },
  ]);
  expect(result).toMatchObject({ committed: true, rejected: false });
}
async function views(editor: DocxEditorInstance) {
  const [before, after] = await Promise.all([
    readCommittedView(editor.surface!.session, groupToken(editor)),
    readWorkingView(editor.surface!.session, groupToken(editor)),
  ]);
  if (!before.success || !after.success) throw new Error('Cannot read native staged views.');
  return [before.view.projection.text, after.view.projection.text];
}
async function geometry(pair: Awaited<ReturnType<typeof createTwoEditorFixture>>, id: string) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
    offset = paragraphTextOf(bob.surface!.session.part(), id)!.length;
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((r) => r.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice.surface!.layout(), { paragraphId: id, offset }),
    b = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Awaited<ReturnType<typeof createTwoEditorFixture>>) {
  const expected = await views(pair.editors[0]!);
  expect(await views(pair.editors[1]!)).toEqual(expected);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const read = readReviewGroups(pkg(pair.editors[0]!));
  if (!read.success) throw new Error(read.error);
  for (const room of [false, true]) {
    const reopened = await createTwoEditorFixture({
      ...options(),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of reopened.editors) {
        expect(await views(editor)).toEqual(expected);
        const groups = readReviewGroups(pkg(editor));
        if (!groups.success) throw new Error(groups.error);
        expect(groups.groups.map((g) => g.id)).toEqual(read.groups.map((g) => g.id));
      }
    } finally {
      reopened.destroy();
    }
  }
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const) {
    test(`native dependent stages preserve concurrent text, history, geometry, and reopened rooms; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        pair.pause();
        stage(alice);
        append(bob, id, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const committed = 'Alpha  end Peer\n\nForeign\n\nTail',
          working = 'Alpha new end Peer\n\nForeign pending\n\nTail',
          prior = 'Alpha old end Peer\n\nForeign pending\n\nTail';
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([committed, working]);
        for (let cycle = 0; cycle < 2; cycle++) {
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([committed, prior]);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([committed, working]);
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            committed.replace(' Peer', ''),
            working.replace(' Peer', ''),
          ]);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await geometry(pair, id);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);

    test(`failed dependent stages preserve canonical references, shared vectors, and prior redo; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        append(alice, id, ' redo');
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const originals = pair.editors.map(pkg),
          revisions = pair.editors.map((e) => e.surface!.session.packageRevision()),
          vectors = pair.documents.map((d) => Y.encodeStateVector(d));
        pair.pause();
        expect(() => stage(alice, true)).toThrow('Injected dependent stage failure.');
        for (const [i, editor] of pair.editors.entries()) {
          expect(pkg(editor)).toBe(originals[i]);
          expect(editor.surface!.session.packageRevision()).toBe(revisions[i]);
          expect(Y.encodeStateVector(pair.documents[i]!)).toEqual(vectors[i]);
        }
        append(bob, bob.surface!.session.paragraphIds().at(-1)!, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha  end redo\n\nForeign\n\nTail Peer',
            'Alpha old end redo\n\nForeign pending\n\nTail Peer',
          ]);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);

    test(`dependent stage history survives concurrent owner deletion; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        pair.pause();
        stage(alice);
        expect(
          bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }])
        ).toMatchObject({ committed: true, rejected: false });
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), id)).toBeNull();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha  end\n\nForeign\n\nTail',
            'Alpha new end\n\nForeign pending\n\nTail',
          ]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([
            'Alpha  end\n\nForeign\n\nTail',
            'Alpha old end\n\nForeign pending\n\nTail',
          ]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await geometry(pair, id);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), id)).toBeNull();
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);
  }
