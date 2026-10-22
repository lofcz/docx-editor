import { expect, spyOn, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { groupToken } from '../../test/support/review-groups';
import { seedReviewCapacity } from '../../test/support/review-capacity';
import { seedReviewFieldCapacity } from '../../test/support/review-field-capacity';
import { REVIEW_FLAG_TARGETS, REVIEW_FLAG_INTENTS } from '../collaboration/review-flag-codec';
import { REVIEW_FIELD_TARGETS, REVIEW_FIELD_INTENTS } from '../collaboration/review-field-codec';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readReviewGroups } from './groups';
import { withInitializedReviewGroupStore, withReviewGroupRecord } from './group-record-store';
import { createEditPlan } from '../diff/plan';
import { replaceOwnedProposals } from './replace-owned-proposals';
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

/** Exercise the validated R25 command through real native collaboration sessions. */
function stage(editor: DocxEditorInstance, fail = false) {
  const session = editor.surface!.session;
  const read = readReviewGroups(pkg(editor));
  if (!read.success) throw new Error(read.error);
  const group = read.groups.find((g) => g.native.some((n) => n.address.id === '17'));
  if (!group) throw new Error('Missing host-owned original proposal.');
  const token = groupToken(editor);
  const candidate = createEditPlan(token, [
    {
      kind: 'text',
      partName: session.part().name,
      paragraphId: session.paragraphIds()[0]!,
      start: 6,
      end: 6,
      before: '',
      after: 'new',
    },
  ]);
  if (!candidate.success) throw new Error(candidate.error);
  const original = session.applyTreeOps.bind(session);
  const injected = fail
    ? spyOn(session, 'applyTreeOps').mockImplementation((ops, a, h, scope, options) =>
        original(ops, a, h, scope, {
          ...options,
          stage: (ctx) => {
            options?.stage?.(ctx);
            throw new Error('Injected dependent stage failure.');
          },
        })
      )
    : null;
  try {
    const result = replaceOwnedProposals(
      editor,
      [
        {
          groupId: group.id,
          native: group.native.map((member) => ({
            partName: member.partName,
            addresses: [member.address],
          })),
          properties: group.properties,
        },
      ],
      candidate.plan,
      { documentId: token.documentId, epoch: token.epoch, author: 'Bot', date }
    );
    if (!result.success) throw new Error(result.error);
    expect(result.committed).toBe(true);
    expect(result.ownership).toHaveLength(1);
  } finally {
    injected?.mockRestore();
  }
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
    test(`owned-proposal replacements preserve concurrent text, history, geometry, and reopened rooms; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        expect(await views(alice)).toEqual([
          'Alpha  end\n\nForeign\n\nTail',
          'Alpha old end\n\nForeign pending\n\nTail',
        ]);
        pair.pause();
        stage(alice);
        expect(await views(alice)).toEqual([
          'Alpha  end\n\nForeign\n\nTail',
          'Alpha new end\n\nForeign pending\n\nTail',
        ]);
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

    test(`failed owned-proposal replacements preserve canonical references, shared vectors, and prior redo; ${clientIds}/${receiver}`, async () => {
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
        expect(await views(alice)).toEqual([
          'Alpha  end\n\nForeign\n\nTail',
          'Alpha old end\n\nForeign pending\n\nTail',
        ]);
        pair.pause();
        stage(alice);
        expect(await views(alice)).toEqual([
          'Alpha  end\n\nForeign\n\nTail',
          'Alpha new end\n\nForeign pending\n\nTail',
        ]);
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

for (const ledger of ['removal', 'field'] as const)
  for (const kind of ['targets', 'intents'] as const)
    test(`owned-proposal replacement ${ledger} ${kind} capacity refusal preserves packages, groups, vectors, and prior redo`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds: [41, 73] });
      try {
        const alice = pair.editors[0]!,
          tail = alice.surface!.session.paragraphIds().at(-1)!;
        append(alice, tail, ' redo');
        await pair.settle();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        (ledger === 'removal' ? seedReviewCapacity : seedReviewFieldCapacity)(
          pair.documents[0]!,
          kind,
          0
        );
        await pair.settle();
        const maps = [
            REVIEW_FLAG_TARGETS,
            REVIEW_FLAG_INTENTS,
            REVIEW_FIELD_TARGETS,
            REVIEW_FIELD_INTENTS,
          ],
          counts = () => pair.documents.map((doc) => maps.map((name) => doc.getMap(name).size)),
          history = () =>
            pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()]),
          before = pair.editors.map((editor) => ({
            pkg: pkg(editor),
            bytes: writeOoxmlPackage(pkg(editor)),
            revision: editor.surface!.session.packageRevision(),
            groups: readReviewGroups(pkg(editor)),
          })),
          vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
          priorCounts = counts(),
          priorHistory = history();
        expect(priorHistory[0]![1]).toBe(true);
        const unchanged = () => {
          for (const [index, editor] of pair.editors.entries()) {
            expect(pkg(editor)).toBe(before[index]!.pkg);
            expect(writeOoxmlPackage(pkg(editor))).toEqual(before[index]!.bytes);
            expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
            expect(readReviewGroups(pkg(editor))).toEqual(before[index]!.groups);
            expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
          }
          expect(counts()).toEqual(priorCounts);
          expect(history()).toEqual(priorHistory);
        };
        pair.pause();
        for (const editor of pair.editors) {
          expect(() => stage(editor)).toThrow('record limit');
          unchanged();
        }
        pair.reconnect(1);
        await pair.settle();
        unchanged();
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail redo');
          expect(readReviewGroups(pkg(editor))).toEqual(before[0]!.groups);
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
      } finally {
        pair.destroy();
      }
    }, 30000);
