import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint, findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { applyTextPlan, paragraphPlan, textPlanDocument } from '../../test/support/text-plan';

test('planned word edits synchronize with native undo, run identity, remote caret geometry, and save/reopen', async () => {
  const pair = await createTwoEditorFixture({
    document: textPlanDocument(),
    documentId: 'paragraph-plan-presence',
  });
  try {
    const [alice, bob] = pair.editors;
    const first = alice!.surface!,
      second = bob!.surface!;
    const id = first.session.paragraphIds()[0]!;
    const original = paragraphTextOf(first.session.part(), id)!;
    const paragraph = findNode(first.session.part(), id)!;
    if (paragraph.kind === 'textValue') throw new Error('Missing paragraph');
    const runs = paragraph.children.filter((node) => node.kind === 'run');
    const keep = [runs[0]!, runs[2]!].map((node) => ({
      id: node.id,
      fingerprint: canonicalOoxmlFingerprint(node),
    }));
    second.setSelection({
      anchor: { paragraphId: id, offset: original.length },
      head: { paragraphId: id, offset: original.length },
    });
    await pair.settle();
    const remoteBefore = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    )!;
    expect(Boolean(remoteBefore)).toBe(true);
    const beforeCaret = caretAt(first.layout(), {
      paragraphId: remoteBefore.head.nodeId,
      offset: remoteBefore.head.offset,
    });
    expect(beforeCaret !== null).toBe(true);
    const target = 'Before hippopotamus after 😀.';
    const plan = paragraphPlan(alice!, id, target);
    expect(applyTextPlan(alice!, plan)).toBe(true);
    await pair.settle();
    expect(paragraphTextOf(second.session.part(), id)).toBe(target);
    for (const preserved of keep) {
      expect(canonicalOoxmlFingerprint(findNode(first.session.part(), preserved.id)!)).toBe(
        preserved.fingerprint
      );
      expect(canonicalOoxmlFingerprint(findNode(second.session.part(), preserved.id)!)).toBe(
        preserved.fingerprint
      );
    }
    const remoteAfter = pair.rooms[0]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'bob'
    )!;
    expect(remoteAfter.head.offset).toBe(target.length);
    expect(second.state().selection?.head.offset).toBe(target.length);
    const afterCaret = caretAt(first.layout(), {
      paragraphId: remoteAfter.head.nodeId,
      offset: remoteAfter.head.offset,
    });
    expect(afterCaret !== null).toBe(true);
    expect(afterCaret!.x).toBeGreaterThan(beforeCaret!.x);
    expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(paragraphTextOf(first.session.part(), id)).toBe(original);
    expect(paragraphTextOf(second.session.part(), id)).toBe(original);
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(paragraphTextOf(first.session.part(), id)).toBe(target);
    expect(paragraphTextOf(second.session.part(), id)).toBe(target);
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
  } finally {
    pair.destroy();
  }
}, 30_000);

test('disjoint paragraph plans converge after concurrent offline edits and preserve peer text through undo/redo', async () => {
  const pair = await createTwoEditorFixture({
    document: textPlanDocument(),
    documentId: 'paragraph-plan-concurrency',
  });
  try {
    const [alice, bob] = pair.editors;
    const id = alice!.surface!.session.paragraphIds()[0]!;
    const read = (index: number) =>
      paragraphTextOf(pair.editors[index]!.surface!.session.part(), id);
    pair.pause();
    expect(applyTextPlan(alice!, paragraphPlan(alice!, id, 'Before car after 😀.'))).toBe(true);
    expect(applyTextPlan(bob!, paragraphPlan(bob!, id, 'Before cat after 😀. Bob'))).toBe(true);
    pair.reconnect();
    await pair.settle();
    expect(read(0)).toBe('Before car after 😀. Bob');
    expect(read(1)).toBe(read(0));
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(read(0)).toBe('Before cat after 😀. Bob');
    expect(read(1)).toBe(read(0));
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(read(0)).toBe('Before car after 😀. Bob');
    expect(read(1)).toBe(read(0));
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
  } finally {
    pair.destroy();
  }
}, 30_000);

test('paragraph deletion wins over a concurrent peer text plan and stale plans cannot recreate removed content', async () => {
  const pair = await createTwoEditorFixture({
    document: textPlanDocument(),
    documentId: 'paragraph-plan-deletion',
  });
  try {
    const [alice, bob] = pair.editors;
    const id = alice!.surface!.session.paragraphIds()[0]!;
    pair.pause();
    expect(applyTextPlan(bob!, paragraphPlan(bob!, id, 'Before car after 😀.'))).toBe(true);
    const stale = paragraphPlan(bob!, id, 'Before cargo after 😀.');
    expect(
      alice!.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    for (const editor of pair.editors)
      expect(paragraphTextOf(editor.surface!.session.part(), id)).toBeNull();
    const before = canonicalOoxmlFingerprint(bob!.surface!.session.part());
    expect(applyTextPlan(bob!, stale)).toBe(false);
    expect(canonicalOoxmlFingerprint(bob!.surface!.session.part())).toBe(before);
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
  } finally {
    pair.destroy();
  }
}, 30_000);
