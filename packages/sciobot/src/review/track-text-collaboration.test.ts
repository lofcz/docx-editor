import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  applyTextPlan,
  paragraphPlan,
  TEXT_DOCUMENT_ID,
  TEXT_EPOCH,
} from '../../test/support/text-plan';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from './committed-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTextPlan } from './track-text';
import { readWorkingView } from './working-view';

const context = {
  documentId: TEXT_DOCUMENT_ID,
  epoch: TEXT_EPOCH,
  author: 'ScioBot',
  date: '2026-10-19T09:00:00Z',
};
const documentBytes = () =>
  docxFixture('<w:p><w:r><w:t>The cat sat</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
const fixture = () =>
  createTwoEditorFixture({
    document: documentBytes(),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
async function views(editor: DocxEditorInstance) {
  const revision = readEditorRevision(editor, TEXT_DOCUMENT_ID, TEXT_EPOCH)!;
  const [original, proposed] = await Promise.all([
    readCommittedView(editor.surface!.session, revision),
    readWorkingView(editor.surface!.session, revision),
  ]);
  if (!original.success) throw new Error(original.error);
  if (!proposed.success) throw new Error(proposed.error);
  return [original.view.projection.text, proposed.view.projection.text];
}

for (const action of ['acceptReviewItem', 'rejectReviewItem'] as const)
  test(`tracked replacement and ${action} converge, preserve remote caret geometry, and survive undo/redo and save`, async () => {
    const pair = await fixture();
    try {
      const [alice, bob] = pair.editors;
      const first = alice!.surface!,
        second = bob!.surface!;
      const id = first.session.paragraphIds()[0]!;
      second.setSelection({
        anchor: { paragraphId: id, offset: 11 },
        head: { paragraphId: id, offset: 11 },
      });
      await pair.settle();
      const caretBefore = caretAt(first.layout(), { paragraphId: id, offset: 11 });
      expect(caretBefore !== null).toBe(true);
      const target = 'The caaaaaaaaaaar sat';
      expect(paragraphPlan(alice!, id, target).operations).toHaveLength(1);
      const result = trackTextPlan(alice!, paragraphPlan(alice!, id, target), context);
      if (!result.success) throw new Error(result.error);
      expect(result.proposals.length).toBeGreaterThan(0);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(await views(editor)).toEqual(['The cat sat\n\nTail', `${target}\n\nTail`]);
        expect(editor.getReviewItems().length).toBe(1);
      }
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(Boolean(remote)).toBe(true);
      expect(remote.head.offset).toBe(second.state().selection!.head.offset);
      const caretAfter = caretAt(first.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      });
      expect(caretAfter !== null).toBe(true);
      expect(caretAfter!.x).toBeGreaterThan(caretBefore!.x);
      expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
      const [pendingAlice, pendingBob] = await pair.savedDigests();
      expect(pendingAlice).toEqual(pendingBob);
      expect(bob![action](bob!.getReviewItems()[0]!.key).ok).toBe(true);
      await pair.settle();
      const decided = action === 'acceptReviewItem' ? target : 'The cat sat';
      for (const editor of pair.editors) {
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(decided);
        expect(editor.getReviewItems()).toHaveLength(0);
      }
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(await views(editor)).toEqual(['The cat sat\n\nTail', `${target}\n\nTail`]);
        expect(editor.getReviewItems().length).toBe(1);
      }
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(decided);
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 30_000);

test('offline tracked replacement merges with a peer’s disjoint edit and undo preserves peer content', async () => {
  const pair = await fixture();
  try {
    const [alice, bob] = pair.editors;
    const id = alice!.surface!.session.paragraphIds()[0]!;
    pair.pause();
    expect(trackTextPlan(alice!, paragraphPlan(alice!, id, 'The car sat'), context).success).toBe(
      true
    );
    expect(applyTextPlan(bob!, paragraphPlan(bob!, id, 'The cat sat Bob'))).toBe(true);
    pair.reconnect();
    await pair.settle();
    for (const editor of pair.editors) {
      expect(await views(editor)).toEqual(['The cat sat Bob\n\nTail', 'The car sat Bob\n\nTail']);
      expect(editor.getReviewItems().length).toBe(1);
    }
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors) {
      expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('The cat sat Bob');
      expect(editor.getReviewItems()).toHaveLength(0);
    }
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(await views(editor)).toEqual(['The cat sat Bob\n\nTail', 'The car sat Bob\n\nTail']);
    expect(alice!.rejectReviewItem(alice!.getReviewItems()[0]!.key).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('The cat sat Bob');
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
  } finally {
    pair.destroy();
  }
}, 30_000);

test('paragraph deletion converges during a peer’s tracked replacement and stale plans cannot restore it', async () => {
  const pair = await fixture();
  try {
    const [alice, bob] = pair.editors;
    const id = alice!.surface!.session.paragraphIds()[0]!;
    pair.pause();
    expect(trackTextPlan(bob!, paragraphPlan(bob!, id, 'The car sat'), context).success).toBe(true);
    const stale = paragraphPlan(bob!, id, 'The cargo sat');
    expect(
      alice!.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    for (const editor of pair.editors) {
      expect(paragraphTextOf(editor.surface!.session.part(), id)).toBeNull();
      expect(editor.getReviewItems()).toHaveLength(0);
    }
    const before = canonicalOoxmlFingerprint(bob!.surface!.session.part());
    const result = trackTextPlan(bob!, stale, context);
    expect(result.success).toBe(false);
    expect(!result.success && result.code).toBe('stale-revision');
    expect(canonicalOoxmlFingerprint(bob!.surface!.session.part())).toBe(before);
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
  } finally {
    pair.destroy();
  }
}, 30_000);
