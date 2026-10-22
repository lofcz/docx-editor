import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  contentControlPropertiesOf,
  findNode,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import {
  fixture,
  readBlocks,
  type DocumentBlock,
} from '../../../../examples/react-blocks/src/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { writeBlockPayload } from './payload-part';
import { readPayloadStore } from './payload-opc-codec';
import type { BlockPayload } from './payload-schema';

function activity(editor: DocxEditorInstance) {
  const flatten = (blocks: DocumentBlock[]): DocumentBlock[] =>
    blocks.flatMap((block) => [block, ...flatten(block.children)]);
  return flatten(readBlocks(editor.surface!.session.part())).find(
    (block) => block.type === 'activity'
  )!;
}
function updateFor(editor: DocxEditorInstance, block: DocumentBlock) {
  const old = paragraphTextOf(editor.surface!.session.part(), block.headerId)!;
  const label = 'Warmup — 15 minutes';
  const payload: BlockPayload = {
    id: 'activity-1',
    type: 'lesson_activity',
    version: 1,
    attributes: { name: 'Warmup', duration: 15 },
  };
  const fallbackOps: TreeDocOp[] = [
    { op: 'deleteText', paragraphId: block.headerId, start: 0, end: old.length },
    { op: 'insertText', paragraphId: block.headerId, offset: 0, text: label },
  ];
  return { controlId: block.id, payload, label, fallbackOps };
}
function payloadOf(editor: DocxEditorInstance) {
  const store = readPayloadStore(editor.surface!.session.currentPackage());
  return store.status === 'ready' ? store.entries.get('activity-1')?.decoded : undefined;
}

test('metadata, readable fallback, and short tag share one undo unit and save/reopen', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: fixture(), author: 'Alice' });
  try {
    const block = activity(editor);
    const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const write = updateFor(editor, block);
    expect(writeBlockPayload(editor, write).success).toBe(true);
    expect(payloadOf(editor)).toEqual({ status: 'supported', value: write.payload });
    expect(paragraphTextOf(editor.surface!.session.part(), block.headerId)).toBe(write.label);
    const control = findNode(editor.surface!.session.part(), block.id)!;
    expect(contentControlPropertiesOf(control).alias).toBe(write.label);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(payloadOf(editor)).toBeUndefined();
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(payloadOf(editor)).toEqual({ status: 'supported', value: write.payload });
    expect(paragraphTextOf(editor.surface!.session.part(), block.headerId)).toBe(write.label);
    const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!opened.ok) throw new Error(opened.reason);
    const saved = readPayloadStore(opened.package);
    expect(saved.status).toBe('ready');
    if (saved.status !== 'ready') throw new Error('No saved payloads');
    expect(saved.entries.get('activity-1')?.decoded).toEqual({
      status: 'supported',
      value: write.payload,
    });
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('failed fallback, viewing mode, and stale reads leave package metadata untouched', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: fixture(), author: 'Alice' });
  try {
    const write = updateFor(editor, activity(editor));
    const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const rejected = writeBlockPayload(editor, {
      ...write,
      fallbackOps: [{ op: 'insertText', paragraphId: 'missing', offset: 0, text: 'Refuse' }],
    });
    expect(rejected.success).toBe(false);
    expect(payloadOf(editor)).toBeUndefined();
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
    editor.surface!.setEditingMode('view');
    expect(writeBlockPayload(editor, write).success).toBe(false);
    expect(payloadOf(editor)).toBeUndefined();
    editor.surface!.setEditingMode('edit');
    expect(writeBlockPayload(editor, { ...write, expectedRevision: -1 })).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(payloadOf(editor)).toBeUndefined();
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('payload writes synchronize with offline peer text, undo/redo, presence, deletion, and reopen', async () => {
  const pair = await createTwoEditorFixture({
    document: fixture(),
    documentId: 'block-payload-write',
  });
  try {
    const first = pair.editors[0]!;
    const second = pair.editors[1]!;
    const block = activity(first);
    const write = updateFor(first, block);
    const bodyId = [...block.paragraphs].find((id) =>
      paragraphTextOf(first.surface!.session.part(), id)?.includes('Plan an experiment.')
    )!;
    pair.pause();
    expect(writeBlockPayload(first, write).success).toBe(true);
    second.surface!.setSelection({
      anchor: { paragraphId: bodyId, offset: 0 },
      head: { paragraphId: bodyId, offset: 0 },
    });
    second.surface!.insertPlainText('Bob: ');
    pair.reconnect();
    await pair.settle();
    expect(payloadOf(first)).toEqual(payloadOf(second));
    expect(paragraphTextOf(second.surface!.session.part(), block.headerId)).toBe(write.label);
    expect(paragraphTextOf(first.surface!.session.part(), bodyId)).toContain('Bob: ');
    expect(pair.rooms[0]!.session.remoteSelections().length).toBeGreaterThan(0);
    expect(first.exec({ type: 'undo' }).ok).toBe(true);
    pair.flush();
    expect(payloadOf(first)).toBeUndefined();
    expect(payloadOf(second)).toBeUndefined();
    expect(paragraphTextOf(second.surface!.session.part(), bodyId)).toContain('Bob: ');
    expect(first.exec({ type: 'redo' }).ok).toBe(true);
    pair.flush();
    expect(payloadOf(first)).toEqual(payloadOf(second));
    pair.pause();
    second.surface!.setSelection({
      anchor: { paragraphId: bodyId, offset: 0 },
      head: { paragraphId: bodyId, offset: 0 },
    });
    second.surface!.insertPlainText('Still editing: ');
    expect(
      first.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: bodyId }]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    expect(paragraphTextOf(first.surface!.session.part(), bodyId)).toBeNull();
    expect(paragraphTextOf(second.surface!.session.part(), bodyId)).toBeNull();
    expect(payloadOf(first)).toEqual(payloadOf(second));
    const [alice, bob] = await pair.savedDigests();
    expect(alice).toEqual(bob);
  } finally {
    pair.destroy();
  }
}, 30000);
