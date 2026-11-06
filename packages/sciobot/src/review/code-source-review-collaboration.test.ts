import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  findNode,
  isJoinedParagraph,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument, openCode } from '../../test/support/code';
import { buildBlockIndex } from '../blocks/block-index';
import { updateCodeLanguage, updateCodeSource } from '../blocks/code-edit';
import { nativeCodeSource } from '../blocks/code-native-source';
import { codeSourceReviewLocks, stageTrackedCodeSource } from './code-source-review';
import type { TrackedStageWriter } from './tracked-stage-writer';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { propertyToken } from '../../test/support/property-execution';

const ORIGINAL = 'one\nmiddle\nlast';
const PROPOSED = 'one\n\tmid\ndle\nlast';
const attribution = { author: 'Alice', date: '2026-11-05T12:00:00.000Z' };
function transaction(editor: DocxEditorInstance, stage: (writer: TrackedStageWriter) => void) {
  const surface = editor.surface!;
  const result = surface.applyAutomationOps(
    () => [],
    { kind: 'body' },
    undefined,
    false,
    true,
    'before',
    (ctx) =>
      stage({
        currentPackage: () => ctx.currentPackage(),
        edit(change) {
          if (!ctx.applyPackage(change)) throw Error('The code review package stage failed.');
        },
        apply(partName, op) {
          if (!ctx.applyTo(partName, op)) throw Error(`The code review operation failed: ${op.op}`);
        },
      })
  );
  if (result.rejected || !result.committed) throw Error(String(result.reason));
}
function state(editor: DocxEditorInstance) {
  const pkg = editor.surface!.session.currentPackage();
  const block = buildBlockIndex(pkg).byId.get('code')!;
  const part = pkg.parts.get(block.partName)!;
  const control = findNode(part, block.controlId)!;
  return {
    pkg,
    block,
    part,
    committed: nativeCodeSource(part, control, 'committed')!,
    working: nativeCodeSource(part, control, 'working')!,
  };
}

test('code review reads exact native source halves without publishing payload cache changes', async () => {
  const pair = await createTwoEditorFixture({
    document: await codeDocument(),
    ownedAttributes: true,
  });
  try {
    const editor = pair.editors[0]!;
    transaction(editor, (writer) =>
      stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
    );
    await pair.settle();
    for (const peer of pair.editors) {
      const before = peer!.surface!.session.currentPackage();
      const committed = await readCommittedView(peer!.surface!.session, propertyToken(peer!));
      const working = await readWorkingView(peer!.surface!.session, propertyToken(peer!));
      if (!committed.success) throw Error(committed.error);
      if (!working.success) throw Error(working.error);
      expect(committed.view.projection.text).toContain('```python\n' + ORIGINAL + '\n```');
      expect(working.view.projection.text).toContain('```python\n' + PROPOSED + '\n```');
      expect(peer!.surface!.session.currentPackage()).toBe(before);
      expect(state(peer!).block.payload.attributes.value).toBe(ORIGINAL);
    }
  } finally {
    pair.destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`native code source ${decision} retains peer language, exact history, cursors, and reopen: ${clientIds}`, async () => {
      const pair = await createTwoEditorFixture({
        document: await codeDocument(),
        clientIds,
        ownedAttributes: true,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const initial = state(alice);
        const originalXml = serializeOoxmlPart(initial.part);
        const caret = { paragraphId: initial.committed.paragraphs[2]!.id, offset: 2 };
        bob.surface!.setSelection({ anchor: caret, head: caret });
        await pair.settle();
        pair.pause();
        transaction(alice, (writer) =>
          stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
        );
        expect(updateCodeLanguage(bob, { blockId: 'code', language: 'typescript' })).toMatchObject({
          success: true,
        });
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const verify = (committed: string, working: string) => {
          for (const editor of pair.editors) {
            const current = state(editor);
            expect(current.committed.value).toBe(committed);
            expect(current.working.value).toBe(working);
            expect(current.block.payload.attributes.language).toBe('typescript');
          }
        };
        verify(ORIGINAL, PROPOSED);
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote).toBeDefined();
        expect(
          caretAt(alice.surface!.layout(), {
            paragraphId: remote.head.nodeId,
            offset: remote.head.offset,
          })
        ).not.toBeNull();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, ORIGINAL);
        expect(serializeOoxmlPart(state(alice).part)).toBe(originalXml);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, PROPOSED);
        // Metadata stays independently editable while the source proposal is pending.
        expect(updateCodeLanguage(bob, { blockId: 'code', language: 'javascript' })).toMatchObject({
          success: true,
        });
        await pair.settle();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, PROPOSED);
        transaction(bob, (writer) => {
          const current = state(bob);
          const locks = codeSourceReviewLocks(current.pkg, ['code']);
          const ids = collectRevisionSites(current.part).map((site) => site.node.id);
          writer.edit(locks.before);
          writer.apply(current.part.name, {
            op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
            siteNodeIds: ids,
          });
          writer.edit(locks.after);
        });
        await pair.settle();
        const result = decision === 'accept' ? PROPOSED : ORIGINAL;
        verify(result, result);
        if (decision === 'reject') {
          expect(state(alice).committed.paragraphs.map((paragraph) => paragraph.id)).toContain(
            initial.committed.paragraphs[1]!.id
          );
          expect(state(alice).committed.paragraphs.filter(isJoinedParagraph)).toHaveLength(1);
          // Shared decisions retain source identities. Word export performs their physical join.
          const exported = openCode(new Uint8Array(await alice.save()));
          expect(serializeOoxmlPart(exported.parts.get(exported.mainDocumentPart)!)).toBe(
            originalXml
          );
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, PROPOSED);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(result, result);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        for (const editor of pair.editors) {
          const reopened = openCode(new Uint8Array(await editor.save()));
          const payload = buildBlockIndex(reopened).byId.get('code')!.payload;
          expect(payload.attributes).toEqual({ language: 'typescript', value: result });
        }
      } finally {
        pair.destroy();
      }
    }, 30_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const tail of [false, true])
    test(`late peer text survives source rejection, both histories, and later code authoring: ${clientIds}/${tail}`, async () => {
      const pair = await createTwoEditorFixture({
        document: await codeDocument(),
        clientIds,
        ownedAttributes: true,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        transaction(alice, (writer) =>
          stageTrackedCodeSource(writer, 'code', ORIGINAL, PROPOSED, attribution)
        );
        await pair.settle();
        const original = state(alice);
        const target = original.committed.paragraphs[tail ? 2 : 1]!;
        const offset = tail ? 1 : 2; // The first physical fragment also contains the proposed tab.
        pair.pause();
        transaction(alice, (writer) => {
          const current = state(alice);
          const locks = codeSourceReviewLocks(current.pkg, ['code']);
          const ids = collectRevisionSites(current.part).map((site) => site.node.id);
          writer.edit(locks.before);
          writer.apply(current.part.name, { op: 'rejectAllRevisions', siteNodeIds: ids });
          writer.edit(locks.after);
        });
        transaction(bob, (writer) => {
          const current = state(bob);
          const locks = codeSourceReviewLocks(current.pkg, ['code']);
          writer.edit(locks.before);
          writer.apply(current.part.name, {
            op: 'insertText',
            paragraphId: target.id,
            offset,
            text: 'PEER',
          });
          writer.edit(locks.after);
        });
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        const joined = tail ? 'one\nmiddPEERle\nlast' : 'one\nmPEERiddle\nlast';
        const pending = tail ? 'one\n\tmid\ndPEERle\nlast' : 'one\n\tmPEERid\ndle\nlast';
        const verify = (committed: string, working: string) => {
          for (const editor of pair.editors) {
            expect(state(editor).committed.value).toBe(committed);
            expect(state(editor).working.value).toBe(working);
          }
        };
        verify(joined, joined);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, pending);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, joined);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(ORIGINAL, ORIGINAL);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, joined);
        const authored = joined.replace('PEER', 'USER\n\t');
        expect(
          updateCodeSource(bob, { blockId: 'code', expectedSource: joined, value: authored })
        ).toMatchObject({ success: true });
        await pair.settle();
        verify(authored, authored);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        verify(joined, joined);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        verify(authored, authored);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        for (const editor of pair.editors) {
          const reopened = openCode(new Uint8Array(await editor.save()));
          expect(buildBlockIndex(reopened).byId.get('code')!.payload.attributes.value).toBe(
            authored
          );
        }
      } finally {
        pair.destroy();
      }
    }, 30_000);
