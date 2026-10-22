import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  storyParagraphs,
} from '@docx-editor.dev/core/store';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from '../blocks/block-index';
import { removeBlock } from '../blocks/remove';
import { beginStreamInsertion, type GeneratedFragment } from './stream-insert';

const context = { documentId: 'generated', epoch: 'mount' };
const fragment = (text: string): GeneratedFragment => ({
  inserted: [{ type: 'paragraph', children: [{ type: 'text', value: text }] }],
});
const activity = (editor: DocxEditorInstance) =>
  buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('activity-1')!.controlId;
const texts = (editor: DocxEditorInstance) => {
  const part = editor.surface!.session.part();
  const phase = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('phase-1')!;
  return storyParagraphs(contentControlContentNodeOf(findNode(part, phase.controlId)!)!).map((p) =>
    paragraphTextOf(part, p.id)
  );
};
const idOf = (editor: DocxEditorInstance, text: string) =>
  editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text)!;

for (const finish of ['finish', 'cancel'] as const)
  test(`two completed generated fragments share one undo unit after ${finish}`, async () => {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({ container, document: blockClipboardDocument() });
    try {
      const initial = new Uint8Array(await editor.save());
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const started = beginStreamInsertion(editor, context, [activity(editor)]);
      if (!started.success) throw new Error(started.error);
      const stream = started.insertion;
      expect(new Uint8Array(await editor.save())).toEqual(initial);
      expect(stream.append(fragment('Generated first')).success).toBe(true);
      expect(
        stream.append({
          inserted: [
            {
              type: 'paragraph',
              children: [
                { type: 'strong', children: [{ type: 'text', value: 'Generated rich' }] },
                { type: 'text', value: ' with ' },
                {
                  type: 'link',
                  url: 'https://example.com',
                  children: [{ type: 'text', value: 'reference' }],
                },
              ],
            },
          ],
        }).success
      ).toBe(true);
      expect(texts(editor)).toEqual([
        'Phase one',
        'Activity text',
        'Note text',
        'Generated first',
        'Generated rich with reference',
        'Phase end',
      ]);
      const landed = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!landed.ok) throw new Error(landed.reason);
      stream[finish]();
      stream[finish]();
      expect(stream.state).toBe('ended');
      expect(stream.append(fragment('Late')).success).toBe(false);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(new Uint8Array(await editor.save())).toEqual(initial);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const redone = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!redone.ok) throw new Error(redone.reason);
      expect(semanticDigest(redone.package.parts.values())).toEqual(
        semanticDigest(landed.package.parts.values())
      );
    } finally {
      editor.destroy();
      container.remove();
    }
  });

for (const receiver of [0, 1] as const)
  test(`streamed insertion preserves offline peer text, shared history and reopen; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({ document: blockClipboardDocument() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const originalY = caretAt(alice.surface!.layout(), {
        paragraphId: idOf(alice, 'Tail'),
        offset: 2,
      })!.y;
      const started = beginStreamInsertion(alice, context, [activity(alice)]);
      if (!started.success) throw new Error(started.error);
      pair.pause();
      expect(started.insertion.append(fragment('First')).success).toBe(true);
      const peer = idOf(bob, 'Activity text');
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: peer, offset: 13, text: ' Bob' },
        ]).rejected
      ).toBe(false);
      const peerOnly = readOoxmlPackage(new Uint8Array(await bob.save()));
      if (!peerOnly.ok) throw new Error(peerOnly.reason);
      const peerOnlyDigest = semanticDigest(peerOnly.package.parts.values());
      pair.reconnect(receiver);
      await pair.settle();
      expect(started.insertion.append(fragment('Second')).success).toBe(true);
      await pair.settle();
      started.insertion.finish();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual([
          'Phase one',
          'Activity text Bob',
          'Note text',
          'First',
          'Second',
          'Phase end',
        ]);
      const landed = await pair.savedDigests();
      expect(landed[0]).toEqual(landed[1]);
      const tail = idOf(bob, 'Tail');
      bob.surface!.setSelection({
        anchor: { paragraphId: tail, offset: 2 },
        head: { paragraphId: tail, offset: 2 },
      });
      await pair.settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      );
      expect(remote).toBeDefined();
      const ys = pair.editors.map(
        (editor) =>
          caretAt(editor.surface!.layout(), {
            paragraphId: remote!.head.nodeId,
            offset: remote!.head.offset,
          })!.y
      );
      expect(ys[0]).toBeGreaterThan(originalY);
      expect(ys[0]).toBe(ys[1]!);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual(['Phase one', 'Activity text Bob', 'Note text', 'Phase end']);
      expect(await pair.savedDigests()).toEqual([peerOnlyDigest, peerOnlyDigest]);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(landed);
      // Deleting the selected scope ends an active request before any later chunk writes.
      const next = beginStreamInsertion(alice, context, [activity(alice)]);
      if (!next.success) throw new Error(next.error);
      pair.pause();
      expect(removeBlock(bob, { controlId: activity(bob), mode: 'delete' }).success).toBe(true);
      const note = idOf(alice, 'Activity text Bob');
      expect(
        alice.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: note, offset: 17, text: ' local' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      const beforeLateChunk = await pair.savedDigests();
      expect(next.insertion.state).toBe('ended');
      expect(next.insertion.append(fragment('Late chunk')).success).toBe(false);
      expect(await pair.savedDigests()).toEqual(beforeLateChunk);
    } finally {
      pair.destroy();
    }
  });

test('a peer edit inside generated output survives another chunk and shared history', async () => {
  const pair = await createTwoEditorFixture({ document: blockClipboardDocument() });
  try {
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    const started = beginStreamInsertion(alice, context, [activity(alice)]);
    if (!started.success) throw new Error(started.error);
    expect(started.insertion.append(fragment('Output')).success).toBe(true);
    await pair.settle();
    const output = idOf(bob, 'Output');
    expect(
      bob.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: output, offset: 6, text: ' Bob' },
      ]).rejected
    ).toBe(false);
    await pair.settle();
    expect(started.insertion.append(fragment('More')).success).toBe(true);
    await pair.settle();
    const landed = await pair.savedDigests();
    expect(landed[0]).toEqual(landed[1]);
    started.insertion.cancel();
    expect(texts(alice)).toContain('Output Bob');
    expect(texts(alice)).toContain('More');
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(texts(alice)).not.toContain('More');
    expect(alice.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(texts(alice)).toContain('Output Bob');
    expect(await pair.savedDigests()).toEqual(landed);
  } finally {
    pair.destroy();
  }
});

test('an intervening local gesture reports a native history split and ends the stream', async () => {
  const pair = await createTwoEditorFixture({ document: blockClipboardDocument() });
  try {
    const alice = pair.editors[0]!;
    const started = beginStreamInsertion(alice, context, [activity(alice)]);
    if (!started.success) throw new Error(started.error);
    expect(started.insertion.append(fragment('First'))).toMatchObject({
      success: true,
      history: { kind: 'started' },
    });
    const tail = idOf(alice, 'Tail');
    expect(
      alice.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' Teacher' }],
        { kind: 'body' },
        undefined,
        false,
        true
      ).rejected
    ).toBe(false);
    expect(started.insertion.append(fragment('Second'))).toMatchObject({
      success: true,
      history: { kind: 'split' },
    });
    expect(started.insertion.state).toBe('ended');
    expect(started.insertion.append(fragment('Third')).success).toBe(false);
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(alice)).toContain('First');
    expect(texts(alice)).not.toContain('Second');
    expect(idOf(alice, 'Tail Teacher')).toBeDefined();
  } finally {
    pair.destroy();
  }
});
