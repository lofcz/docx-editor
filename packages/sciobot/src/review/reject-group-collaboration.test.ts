import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  groupDecisionFixture,
  proposeThreeProperties,
  blockLabel,
} from '../../test/support/group-decisions';
import { groups, groupToken } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { readPropertyReviewStore } from './property-record-store';
import { readReviewGroups } from './groups';
import { rejectReviewGroup } from './reject-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function reject(editor: DocxEditorInstance, block = 'middle') {
  const group = groups(editor).find((group) =>
    group.properties.some((member) => member.blockId === block)
  );
  if (!group) throw new Error('Missing block group');
  return rejectReviewGroup(editor, group.id, groupToken(editor), propertyContext);
}
function paragraph(editor: DocxEditorInstance, text: string) {
  const id = editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text);
  if (!id) throw new Error('Missing paragraph');
  return id;
}
function type(editor: DocxEditorInstance, id: string) {
  const surface = editor.surface!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText('Peer: ');
  surface.flushPendingInput();
}
function stored(editor: DocxEditorInstance) {
  const values = readPayloadStore(source(editor)),
    proposals = readPropertyReviewStore(source(editor));
  if (values.status !== 'ready' || proposals.status !== 'ready') throw new Error('Missing stores');
  return {
    payloads: [...values.entries].map(([id, value]) => [id, value.raw, value.label]),
    proposal: proposals.records.get('three-blocks')!.raw,
  };
}
function fixtureDocument() {
  const pkg = groupDecisionFixture(),
    owner = buildBlockIndex(pkg).byId.get('middle')!;
  const attributes = { ...owner.payload.attributes };
  delete attributes.duration;
  const written = withPayloadRecord(pkg, { ...owner.payload, attributes }, 'Stored middle');
  if (!written.ok) throw new Error(written.reason);
  return writeOoxmlPackage(written.pkg);
}
function baseline(editor: DocxEditorInstance) {
  for (const id of ['first', 'middle', 'last']) {
    const attributes = buildBlockIndex(source(editor)).byId.get(id)!.payload.attributes;
    expect(attributes.name).toBe(`Old ${id}`);
    expect(Object.hasOwn(attributes, 'duration')).toBe(id !== 'middle');
    if (id !== 'middle') expect(attributes.duration).toBe(10);
  }
}
async function saved(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  ids: readonly string[]
) {
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const editor of pair.editors) {
    const read = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!read.ok) throw new Error(read.reason);
    const pending = readReviewGroups(read.package);
    if (!pending.success) throw new Error(pending.error);
    expect(pending.groups.map((group) => group.id)).toEqual([...ids]);
    expect(
      Object.hasOwn(
        buildBlockIndex(read.package).byId.get('middle')!.payload.attributes,
        'duration'
      )
    ).toBe(false);
  }
  pair.ownership.forEach((handle) => handle.assertHealthy());
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const metadata of [false, true])
      test(`independent property rejection retains absent values, immutable records and each actor history; ${clients}/${receiver}/metadata=${metadata}`, async () => {
        const pair = await createTwoEditorFixture({
          document: fixtureDocument(),
          clientIds: clients,
          ownedAttributes: true,
          editor: editorOptions,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          proposeThreeProperties(alice, metadata);
          await pair.settle();
          const ids = groups(alice).map((group) => group.id),
            before = stored(alice),
            body = paragraph(bob, 'Body middle');
          expect(ids).toHaveLength(3);
          pair.pause();
          expect(reject(alice, 'first').success).toBe(true);
          type(bob, body);
          expect(reject(bob).success).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          const state = (pending: readonly string[]) => {
            for (const editor of pair.editors) {
              expect(groups(editor).map((group) => group.id)).toEqual([...pending]);
              expect(stored(editor)).toEqual(before);
              baseline(editor);
              expect(paragraphTextOf(editor.surface!.session.part(), body)).toBe(
                'Peer: Body middle'
              );
            }
          };
          state([ids[2]!]);
          for (const editor of pair.editors) {
            expect(blockLabel(editor, 'first')).toBe('Old first');
            expect(blockLabel(editor, 'middle')).toBe('Old middle');
          }
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          state([ids[0]!, ids[2]!]);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          state(ids);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          state([ids[1]!, ids[2]!]);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          state([ids[2]!]);
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (row) => row.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(body);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: body, offset: 6 })
          );
          expect(carets[0]).not.toBeNull();
          expect(carets[1]).not.toBeNull();
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          await saved(pair, [ids[2]!]);
        } finally {
          pair.destroy();
        }
      }, 60000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`concurrent rejection of one property group remains resolved until the final actor undo; ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: fixtureDocument(),
        clientIds: clients,
        ownedAttributes: true,
        editor: editorOptions,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        proposeThreeProperties(alice);
        await pair.settle();
        const ids = groups(alice).map((group) => group.id),
          before = stored(alice);
        pair.pause();
        expect(reject(alice).success).toBe(true);
        expect(reject(bob).success).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        const state = (pending: boolean) => {
          for (const editor of pair.editors) {
            expect(groups(editor).map((group) => group.id)).toEqual(
              pending ? ids : [ids[0]!, ids[2]!]
            );
            expect(stored(editor)).toEqual(before);
            baseline(editor);
            if (!pending) expect(blockLabel(editor, 'middle')).toBe('Old middle');
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
        await saved(pair, [ids[0]!, ids[2]!]);
      } finally {
        pair.destroy();
      }
    }, 60000);

for (const receiver of [0, 1] as const)
  test(`rejection during owner deletion restores the old fallback and peer text when deletion is undone; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: fixtureDocument(),
      clientIds: [41, 73],
      ownedAttributes: true,
      editor: editorOptions,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      proposeThreeProperties(alice);
      await pair.settle();
      const ids = groups(alice).map((group) => group.id),
        owner = buildBlockIndex(source(alice)).byId.get('middle')!,
        body = paragraph(bob, 'Body middle');
      pair.pause();
      expect(
        alice.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId: owner.controlId, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).committed
      ).toBe(true);
      type(bob, body);
      expect(reject(bob).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(buildBlockIndex(source(editor)).byId.get('middle')).toBeUndefined();
        expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        baseline(editor);
        expect(blockLabel(editor, 'middle')).toBe('Old middle');
        expect(paragraphTextOf(editor.surface!.session.part(), body)).toBe('Peer: Body middle');
        expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
        expect(caretAt(editor.surface!.layout(), { paragraphId: body, offset: 0 })).not.toBeNull();
      }
      await saved(pair, [ids[0]!, ids[2]!]);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(buildBlockIndex(source(editor)).byId.get('middle')).toBeUndefined();
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  }, 60000);
