import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  acceptBlockGroup,
  blockLabel,
  groupDecisionFixture,
  proposeThreeProperties,
} from '../../test/support/group-decisions';
import {
  groupFixture,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import {
  atomicFixture,
  executeAtomic,
  proposedVisual,
  proposedListening,
} from '../../test/support/atomic-execution';
import { propertyContext } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { readPropertyReviewStore } from './property-record-store';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptReviewGroup } from './accept-group';
import { readReviewGroups } from './groups';

const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
function paragraph(editor: DocxEditorInstance, text: string) {
  const session = editor.surface!.session,
    id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error('Missing ' + text);
  return id;
}
function type(editor: DocxEditorInstance, id: string, text: string) {
  const surface = editor.surface!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
function proposalRaw(editor: DocxEditorInstance) {
  const store = readPropertyReviewStore(source(editor));
  if (store.status !== 'ready') throw new Error('Missing proposal');
  return store.records.get('three-blocks')!.raw;
}

for (const receiver of [0, 1] as const)
  test(`concurrent acceptance of the same property group keeps one payload value and exact fallback label; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupDecisionFixture()),
      clientIds: [41, 73],
      ownedAttributes: true,
      editor: editorOptions,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      proposeThreeProperties(alice);
      await pair.settle();
      pair.pause();
      expect(acceptBlockGroup(alice).success).toBe(true);
      expect(acceptBlockGroup(bob).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(groups(editor)).toHaveLength(2);
        expect(buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes.name).toBe(
          'New middle'
        );
        const payloads = readPayloadStore(source(editor));
        expect(payloads.status === 'ready' && payloads.entries.get('middle')!.label).toBe(
          'New middle'
        );
        expect(blockLabel(editor, 'middle')).toBe('New middle');
      }
      const remaining = groups(alice).map((group) => group.id);
      for (const [actor, action] of [
        [alice, 'undo'],
        [alice, 'redo'],
        [bob, 'undo'],
        [bob, 'redo'],
      ] as const) {
        expect(actor.exec({ type: action }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const payloads = readPayloadStore(source(editor));
          expect(payloads.status === 'ready' && payloads.entries.get('middle')!.label).toBe(
            'New middle'
          );
          expect(buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes.name).toBe(
            'New middle'
          );
          expect(blockLabel(editor, 'middle')).toBe('New middle');
          expect(groups(editor).map((group) => group.id)).toEqual(remaining);
        }
      }
      await assertSaved(pair, remaining);
    } finally {
      pair.destroy();
    }
  }, 30000);
async function assertSaved(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  expectedIds: readonly string[]
) {
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const editor of pair.editors) {
    const read = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!read.ok) throw new Error(read.reason);
    const result = readReviewGroups(read.package);
    if (!result.success) throw new Error(result.error);
    expect(result.groups.map((group) => group.id)).toEqual([...expectedIds]);
  }
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const metadata of [false, true])
      test(`independent group decisions from one proposal preserve native sites, peer text, both actors' history, geometry, and room/DOCX reopening; ${clients}/${receiver}/metadata=${metadata}`, async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(groupDecisionFixture()),
          clientIds: clients,
          editor: editorOptions,
          ownedAttributes: true,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          proposeThreeProperties(alice, metadata);
          await pair.settle();
          const before = groups(alice).map((group) => group.id),
            raw = proposalRaw(alice),
            id = paragraph(bob, 'Body middle');
          pair.pause();
          expect(acceptBlockGroup(alice).success).toBe(true);
          type(bob, id, 'Peer: ');
          expect(acceptBlockGroup(bob, 'first').success).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(groups(editor).map((group) => group.id)).toEqual([before[2]!]);
            expect(proposalRaw(editor)).toBe(raw);
            const index = buildBlockIndex(source(editor));
            expect(index.byId.get('middle')!.payload.attributes.name).toBe('New middle');
            expect(index.byId.get('first')!.payload.attributes.name).toBe('New first');
            expect(index.byId.get('last')!.payload.attributes.name).toBe('Old last');
            expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body middle');
            const pending = readStoredPropertyProposals(source(editor), groupToken(editor));
            if (!pending.success) throw new Error(pending.error);
            expect(pending.pending[0]!.plan.operations).toHaveLength(2);
          }
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (row) => row.actorId === 'bob'
          );
          expect(remote).toBeDefined();
          expect(
            caretAt(alice.surface!.layout(), {
              paragraphId: remote!.head.nodeId,
              offset: remote!.head.offset,
            })
          ).not.toBeNull();
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(groups(editor).map((group) => group.id)).toEqual(before.slice(1));
            expect(buildBlockIndex(source(editor)).byId.get('first')!.payload.attributes.name).toBe(
              'New first'
            );
            expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body middle');
          }
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(groups(editor).map((group) => group.id)).toEqual([before[0]!, before[2]!]);
            expect(
              buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes.name
            ).toBe('New middle');
            expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body middle');
          }
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await assertSaved(pair, [before[2]!]);
          const snapshot = Y.encodeStateAsUpdate(pair.documents[receiver]!);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(),
            roomSnapshot: snapshot,
            editor: editorOptions,
            ownedAttributes: true,
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) {
              expect(groups(editor).map((group) => group.id)).toEqual([before[2]!]);
              expect(proposalRaw(editor)).toBe(raw);
              expect(
                paragraphTextOf(
                  editor.surface!.session.part(),
                  paragraph(editor, 'Peer: Body middle')
                )
              ).toBe('Peer: Body middle');
            }
          } finally {
            reopened.destroy();
          }
        } finally {
          pair.destroy();
        }
      }, 30000);

for (const receiver of [0, 1] as const)
  test(`native group acceptance preserves concurrent text inside the changed paragraph and other peer decisions; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupFixture()),
      clientIds: [41, 73],
      editor: editorOptions,
      ownedAttributes: true,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      proposeGroupText(alice, 'The cat sat.', 'The car sat.');
      proposeGroupText(alice, 'A dog ran.', 'A dog slept.');
      proposeGroupText(alice, 'Keep tail', 'Keep tail!');
      await pair.settle();
      const ids = groups(alice).map((group) => group.id),
        middle = groups(bob)[1]!.anchor!.paragraphId!;
      expect(ids).toHaveLength(3);
      pair.pause();
      expect(acceptReviewGroup(alice, ids[1]!, groupToken(alice), propertyContext).success).toBe(
        true
      );
      type(bob, middle, 'Peer: ');
      expect(acceptReviewGroup(bob, ids[0]!, groupToken(bob), propertyContext).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(groups(editor).map((group) => group.id)).toEqual([ids[2]!]);
        expect(paragraphTextOf(editor.surface!.session.part(), middle)).toBe('Peer: A dog slept.');
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(groups(editor).map((group) => group.id)).toEqual(ids.slice(1));
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await assertSaved(pair, [ids[2]!]);
    } finally {
      pair.destroy();
    }
  }, 30000);

for (const receiver of [0, 1] as const)
  test(`atomic groups accept independent visual and listening payloads through peer history and reconnect; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(atomicFixture()),
      clientIds: [41, 73],
      editor: editorOptions,
      ownedAttributes: true,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      expect(executeAtomic(alice, [proposedVisual, proposedListening]).success).toBe(true);
      await pair.settle();
      pair.pause();
      expect(acceptBlockGroup(alice, 'visual').success).toBe(true);
      type(bob, paragraph(bob, 'Phase body'), 'Peer: ');
      expect(acceptBlockGroup(bob, 'audio').success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(groups(editor)).toEqual([]);
        const index = buildBlockIndex(source(editor));
        expect(index.byId.get('visual')!.payload).toEqual(proposedVisual);
        expect(index.byId.get('audio')!.payload).toEqual(proposedListening);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(groups(editor)[0]!.properties[0]!.blockId).toBe('visual');
        expect(buildBlockIndex(source(editor)).byId.get('audio')!.payload).toEqual(
          proposedListening
        );
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await assertSaved(pair, []);
    } finally {
      pair.destroy();
    }
  }, 30000);

for (const receiver of [0, 1] as const)
  test(`acceptance during owner deletion survives deletion undo with accepted fallback, payload, and peer text; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(groupDecisionFixture()),
      clientIds: [41, 73],
      editor: editorOptions,
      ownedAttributes: true,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      proposeThreeProperties(alice);
      await pair.settle();
      const ids = groups(alice).map((group) => group.id),
        owner = buildBlockIndex(source(alice)).byId.get('middle')!,
        id = paragraph(bob, 'Body middle');
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
      type(bob, id, 'Peer: ');
      expect(acceptBlockGroup(bob).success).toBe(true);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(buildBlockIndex(source(editor)).byId.get('middle')).toBeUndefined();
        expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes.name).toBe(
          'New middle'
        );
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer: Body middle');
        expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
        expect(caretAt(editor.surface!.layout(), { paragraphId: id, offset: 0 })).not.toBeNull();
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      await assertSaved(pair, [ids[0]!, ids[2]!]);
    } finally {
      pair.destroy();
    }
  }, 30000);
