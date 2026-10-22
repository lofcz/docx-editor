import { expect, test } from 'bun:test';
import {
  createBrowserAutomationHost,
  createDocxEditor,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { fixture } from '../../../../examples/react-blocks/src/document';
import { REVIEW, createBlockBridge } from '../../../../examples/react-blocks/src/bridge';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  RevisionProjectionCache,
  readEditorRevision,
  type ProjectionCacheKey,
} from './revision-cache';

const identity = { id: 'lesson-1', path: 'lesson', kind: 'lesson_plan' } as const;
const key: ProjectionCacheKey = {
  identity,
  revision: { documentId: identity.id, epoch: 'mount-1', revision: 0 },
  view: 'working',
};

test('reuse is limited to the same document, load epoch, package revision, and view', () => {
  const cache = new RevisionProjectionCache<object>();
  const first = cache.get(key, () => ({ first: true }));
  expect(cache.get({ ...key }, () => ({ wrong: true }))).toBe(first);
  const committed = cache.get({ ...key, view: 'committed' }, () => ({ committed: true }));
  expect(committed).not.toBe(first);
  for (const next of [
    { ...key, revision: { ...key.revision, revision: 1 } },
    { ...key, revision: { ...key.revision, epoch: 'mount-2' } },
    { ...key, identity: { ...identity, path: 'renamed' } },
    {
      ...key,
      identity: { ...identity, id: 'lesson-2' },
      revision: { ...key.revision, documentId: 'lesson-2' },
    },
  ])
    expect(cache.get(next, () => ({ fresh: true }))).not.toBe(first);
  expect(() =>
    cache.get({ ...key, revision: { ...key.revision, documentId: 'other' } }, () => ({}))
  ).toThrow();
});

test('explicit invalidation and failed builders never return an older projection', () => {
  const cache = new RevisionProjectionCache<object>();
  const first = cache.get(key, () => ({}));
  cache.invalidate();
  expect(() =>
    cache.get(key, () => {
      throw new Error('Projection refused');
    })
  ).toThrow('Projection refused');
  expect(cache.get(key, () => ({}))).not.toBe(first);
});

interface ModelSnapshot {
  readonly text: string | null;
  readonly fingerprint: string;
}
function reader(editor: DocxEditorInstance, epoch: string, target: string) {
  const cache = new RevisionProjectionCache<ModelSnapshot>();
  return () => {
    const revision = readEditorRevision(editor, identity.id, epoch);
    if (!revision) throw new Error('No live revision');
    return cache.get({ identity, revision, view: 'working' }, () => ({
      text: paragraphTextOf(editor.surface!.session.part(), target),
      fingerprint: canonicalOoxmlFingerprint(editor.surface!.session.part()),
    }));
  };
}
function targetOf(editor: DocxEditorInstance) {
  return editor
    .surface!.session.paragraphIds()
    .find((id) =>
      paragraphTextOf(editor.surface!.session.part(), id)?.includes('Plan an experiment.')
    )!;
}
function select(editor: DocxEditorInstance, id: string) {
  editor.surface!.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
}

test('buffered typing, undo/redo, and review acceptance/rejection invalidate real editor reads', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: fixture(),
    author: 'Alice',
    modules: [REVIEW],
  });
  const bridge = createBlockBridge(editor);
  const host = createBrowserAutomationHost(editor);
  const decide = (op: 'acceptAllRevisions' | 'rejectAllRevisions') => {
    const document = host.execute({ operations: [{ op: 'getDocument' }] }).results[0]!;
    if (document.status !== 'ok' || document.value.kind !== 'handle')
      throw new Error('No document');
    const operation =
      op === 'acceptAllRevisions'
        ? { op: 'acceptAllRevisions' as const, document: document.value.handle }
        : { op: 'rejectAllRevisions' as const, document: document.value.handle };
    expect(host.execute({ operations: [operation] }).ok).toBe(true);
  };
  try {
    const id = targetOf(editor);
    const read = reader(editor, 'mount-1', id);
    const original = read();
    expect(read()).toBe(original);
    select(editor, id);
    editor.surface!.enqueueType('User: ');
    const typed = read();
    expect(typed).not.toBe(original);
    expect(typed.text).toContain('User: ');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const undone = read();
    expect(undone).not.toBe(typed);
    expect(undone.text).toBe(original.text);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(read().text).toBe(typed.text);
    bridge.propose();
    const proposed = read();
    decide('acceptAllRevisions');
    const accepted = read();
    expect(accepted).not.toBe(proposed);
    expect(accepted.fingerprint).not.toBe(proposed.fingerprint);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const pendingAgain = read();
    decide('rejectAllRevisions');
    expect(read().fingerprint).not.toBe(pendingAgain.fingerprint);
  } finally {
    host.dispose();
    bridge.destroy();
    editor.destroy();
    container.remove();
  }
});

test('two-session concurrent edits, undo/redo, deletion, and reconnect invalidate peer caches', async () => {
  const pair = await createTwoEditorFixture({
    document: fixture(),
    documentId: 'projection-cache',
  });
  const first = pair.editors[0]!;
  const second = pair.editors[1]!;
  try {
    const id = targetOf(first);
    const readFirst = reader(first, 'alice-mount', id);
    const readSecond = reader(second, 'bob-mount', id);
    const original = readSecond();
    pair.pause();
    select(first, id);
    first.surface!.insertPlainText('Alice: ');
    select(second, id);
    second.surface!.insertPlainText('Bob: ');
    pair.reconnect();
    await pair.settle();
    expect(readSecond()).not.toBe(original);
    expect(readFirst().text).toBe(readSecond().text);
    expect(readSecond().text).toContain('Alice: ');
    expect(readSecond().text).toContain('Bob: ');
    expect(pair.rooms[0]!.session.remoteSelections().length).toBeGreaterThan(0);
    const combined = readSecond();
    expect(first.exec({ type: 'undo' }).ok).toBe(true);
    pair.flush();
    expect(readSecond()).not.toBe(combined);
    expect(first.exec({ type: 'redo' }).ok).toBe(true);
    pair.flush();
    expect(readSecond().text).toBe(combined.text);
    pair.pause();
    second.surface!.insertPlainText('Still editing: ');
    expect(
      first.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    expect(readFirst().text).toBeNull();
    expect(readSecond().text).toBeNull();
    const digests = await pair.savedDigests();
    expect(digests[0]).toEqual(digests[1]);
  } finally {
    pair.destroy();
  }
}, 30000);
