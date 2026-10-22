import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  LINK_CONTEXT,
  linkDocument,
  linkPlan,
  linkTargets,
  linkViews,
} from '../../test/support/link-review';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';

const fixture = (clientIds: readonly [number, number]) =>
  createTwoEditorFixture({
    document: linkDocument(),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    ownedAttributes: true,
    clientIds,
  });
const paragraphStarting = (editor: DocxEditorInstance, prefix: string) => {
  const part = editor.surface!.session.part();
  return editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(part, id)?.startsWith(prefix))!;
};
function propose(editor: DocxEditorInstance) {
  const result = trackHyperlinkPlan(
    editor,
    linkPlan(editor, { url: 'https://example.org/new', tooltip: 'New tip' }),
    LINK_CONTEXT
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
}
async function same(pair: Awaited<ReturnType<typeof fixture>>) {
  const [a, b] = await Promise.all(pair.editors.map((editor) => linkViews(editor)));
  expect(a).toEqual(b);
  expect(linkTargets(pair.editors[0]!)).toEqual(linkTargets(pair.editors[1]!));
  return a!;
}
const ORIGINAL = [
  ['https://example.com', 'Old tip'],
  ['https://example.com', null],
];
const ACCEPTED = [
  ['https://example.org/new', 'New tip'],
  ['https://example.com', null],
];

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`offline retarget keeps peer text and the shared link through history and ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const pair = await fixture(clientIds);
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          propose(alice!);
          const tail = paragraphStarting(bob!, 'Tail');
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: tail, offset: 4, text: ' Bob' },
            ]).rejected
          ).toBe(false);
          pair.reconnect(first);
          await pair.settle();
          const [committed, working] = await same(pair);
          expect(committed).toContain('Tail Bob');
          expect(working).toContain('Tail Bob');
          expect(committed).toContain('https://example.com');
          expect(working).toContain('https://example.org/new');
          // Bob's caret after the pending link pair sits at one native position in both sessions.
          const first_ = paragraphStarting(bob!, 'See');
          bob!.surface!.setSelection({
            anchor: { paragraphId: first_, offset: 2 },
            head: { paragraphId: first_, offset: 2 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(first_);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: first_, offset: 2 })
          );
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, committed]);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const resolved = action === 'accept' ? working : committed;
          expect(await same(pair)).toEqual([resolved, resolved]);
          expect(linkTargets(alice!)).toEqual(action === 'accept' ? ACCEPTED : ORIGINAL);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([committed, working]);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await same(pair)).toEqual([resolved, resolved]);
          const [savedAlice, savedBob] = await pair.savedDigests();
          expect(savedAlice).toEqual(savedBob);
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`peer text typed at the end of the link survives once in both views and both decisions; clients ${clientIds.join('/')}`, async () => {
    const pair = await fixture(clientIds);
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      propose(alice!);
      // "See the site" — offset 12 is the end of the link text.
      const id = paragraphStarting(bob!, 'See');
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: id, offset: 12, text: '!' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(1);
      await pair.settle();
      const [committed, working] = await same(pair);
      const plain = (text: string) => text.replace(/\]\([^)]*\)/g, ']').replace(/[[\]*_\\]/g, '');
      expect(plain(committed!)).toContain('See the site! or twin');
      // Known limit (R43): the transport orders this boundary insertion against the inserted
      // copy by client id, so Proposed reads it on either side of the link. It is never lost.
      expect(['See the site! or twin', 'See !the site or twin']).toContain(
        plain(working!).split('\n')[0]!
      );
      for (const action of ['reject', 'accept'] as const) {
        expect(alice!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        await pair.settle();
        const [resolved] = await same(pair);
        expect(plain(resolved!)).toBe(plain(action === 'reject' ? committed! : working!));
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([committed, working]);
      }
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`a peer deleting the retargeted link while offline converges; clients ${clientIds.join('/')}`, async () => {
    const pair = await fixture(clientIds);
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      propose(alice!);
      const id = paragraphStarting(bob!, 'See');
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'deleteText', paragraphId: id, start: 4, end: 12 },
        ]).rejected
      ).toBe(false);
      pair.reconnect(0);
      await pair.settle();
      const [committed, working] = await same(pair);
      expect(committed).not.toContain('site');
      // The proposal copy was Alice's unseen insertion; it survives as the pending suggestion.
      expect(working).toContain('site');
      expect(alice!.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      await pair.settle();
      const [rejected] = await same(pair);
      expect(rejected).not.toContain('site');
      expect(rejected).toContain('twin');
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 60_000);
