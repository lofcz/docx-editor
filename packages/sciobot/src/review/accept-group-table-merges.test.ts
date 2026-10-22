import { expect, test } from 'bun:test';
import { acceptReviewGroup } from './accept-group';
import { groups, proposeGroupText } from '../../test/support/review-groups';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { readReviewGroups } from './groups';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import { planTableMerge, type TableMergeTarget } from '../diff/table-merges';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableMergePlan } from './track-table-merges';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string, c: string) =>
  `<w:tr>${cell(a, 1200)}${cell(b, 2400)}${cell(c, 3600)}</w:tr>`;
const fixture = (clientIds: readonly [number, number], neighbors = false) =>
  createTwoEditorFixture({
    document: docxFixture(
      (neighbors ? p('First old') : '') +
        '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>' +
        row('A', 'B', 'C') +
        row('D', 'E', 'F') +
        row('G', 'H', 'I') +
        '</w:tbl>' +
        p('Tail') +
        (neighbors ? p('Last old') : '')
    ),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    ownedAttributes: true,
    clientIds,
  });
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const ids = (part: OoxmlPart) =>
  children(tableOf(part))
    .filter((r) => r.kind === 'tableRow')
    .map((r) =>
      children(r)
        .filter((n) => n.kind === 'tableCell')
        .map((n) => n.id)
    );
const paragraphNamed = (editor: DocxEditorInstance, text: string) => {
  const part = editor.surface!.session.part();
  return editor.surface!.session.paragraphIds().find((id) => paragraphTextOf(part, id) === text)!;
};
function propose(editor: DocxEditorInstance, target: (cells: string[][]) => TableMergeTarget) {
  const part = editor.surface!.session.part();
  const plan = planTableMerge(
    {
      revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
      part,
      tableId: tableOf(part).id,
    },
    target(ids(part))
  );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTableMergePlan(editor, plan.plan, CELL_CONTEXT);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
}
const append = (editor: DocxEditorInstance, word: string, text: string) => {
  const id = paragraphNamed(editor, word);
  expect(
    editor.surface!.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: id, offset: word.length, text },
    ]).rejected
  ).toBe(false);
};
async function same(pair: Awaited<ReturnType<typeof fixture>>) {
  const [a, b] = await Promise.all(pair.editors.map((editor) => cellViews(editor)));
  expect(a).toEqual(b);
  return a!;
}
function acceptSelected(editor: DocxEditorInstance, groupId?: string) {
  const currentGroups = groups(editor);
  if (groupId === undefined) expect(currentGroups).toHaveLength(1);
  const result = acceptReviewGroup(
    editor,
    groupId ?? currentGroups[0]!.id,
    readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
    CELL_CONTEXT
  );
  if (!result.success) throw new Error(JSON.stringify(result));
  return { ok: true };
}
const rectangle = (cells: string[][]): TableMergeTarget => ({
  kind: 'merge',
  cellIds: [cells[0]![0]!, cells[0]![1]!, cells[1]![0]!, cells[1]![1]!],
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`offline merge proposal keeps owner and outside peer text through history and accept; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await fixture(clientIds);
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        propose(alice!, rectangle);
        append(bob!, 'A', ' Bob');
        append(bob!, 'I', ' Bob');
        pair.reconnect(first);
        await pair.settle();
        const [committed, working] = await same(pair);
        expect(committed).toContain('| A Bob | B | C |');
        expect(working.replace(/<br>/g, ' ')).toContain('A Bob B D E');
        expect(working).toContain('I Bob');
        const owner = paragraphNamed(bob!, 'A Bob');
        bob!.surface!.setSelection({
          anchor: { paragraphId: owner, offset: 5 },
          head: { paragraphId: owner, offset: 5 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote.head.nodeId).toBe(owner);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId: owner, offset: 5 })
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
        expect(acceptSelected(bob!).ok).toBe(true);
        await pair.settle();
        const resolved = working;
        expect(await same(pair)).toEqual([resolved, resolved]);
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
  test(`accepting a merge after peer text enters a source cell keeps the proposal content; clients ${clientIds.join('/')}, accept`, async () => {
    const pair = await fixture(clientIds);
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      propose(alice!, rectangle);
      append(bob!, 'E', ' Bob');
      pair.reconnect(0);
      await pair.settle();
      const [committed] = await same(pair);
      expect(committed).toContain('| D | E Bob | F |');
      expect(acceptSelected(bob!).ok).toBe(true);
      await pair.settle();
      const [resolved] = await same(pair);
      // A merge proposal gathers the content it saw; the source cell itself is deleted on accept.
      expect(resolved).not.toContain('E Bob');
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
  test(`offline split proposal keeps peer text in the split cell through accept; clients ${clientIds.join('/')}`, async () => {
    const pair = await fixture(clientIds);
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      propose(alice!, (cells) => ({ kind: 'split', cellId: cells[1]![1]!, rows: 2, columns: 2 }));
      append(bob!, 'E', ' Bob');
      pair.reconnect(1);
      await pair.settle();
      const [committed, working] = await same(pair);
      expect(committed).toContain('| D | E Bob | F |');
      expect(working).toContain('E Bob');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(await same(pair)).toEqual([committed, committed]);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(acceptSelected(bob!).ok).toBe(true);
      await pair.settle();
      const resolved = working;
      expect(await same(pair)).toEqual([resolved, resolved]);
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
  for (const first of [0, 1] as const)
    test(`accepting a complete merge group keeps both neighbor groups pending; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await fixture(clientIds, true);
      try {
        const [alice, bob] = pair.editors;
        pair.pause();
        proposeGroupText(alice!, 'First old', 'First new');
        propose(alice!, rectangle);
        proposeGroupText(alice!, 'Last old', 'Last new');
        append(bob!, 'I', ' Bob');
        pair.reconnect(first);
        await pair.settle();
        const pending = groups(bob!);
        expect(pending).toHaveLength(3);
        const ids = pending.map((group) => group.id);
        expect(pending[1]!.native.length).toBeGreaterThan(1);
        const [committed, working] = await same(pair);
        expect(committed).toContain('First old');
        expect(committed).toContain('| A | B | C |');
        expect(working).toContain('First new');
        expect(working.replace(/<br>/g, ' ')).toContain('A B D E');
        expect(working).toContain('I Bob');
        expect(acceptSelected(bob!, ids[1]).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
        const [accepted, proposed] = await same(pair);
        expect(accepted).toBe(
          working.replace('First new', 'First old').replace('Last new', 'Last old')
        );
        expect(proposed).toBe(working);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([committed, working]);
        for (const editor of pair.editors)
          expect(groups(editor).map((group) => group.id)).toEqual(ids);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([accepted, working]);
        for (const editor of pair.editors)
          expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
        for (const editor of pair.editors) {
          const current = groups(editor);
          expect(current.map((group) => group.native)).toEqual([
            pending[0]!.native,
            pending[2]!.native,
          ]);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          const reopened = readReviewGroups(saved.package);
          if (!reopened.success) throw new Error(reopened.error);
          expect(reopened.groups.map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
          expect(reopened.groups.every((group) => !group.readOnly)).toBe(true);
        }
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
  for (const first of [0, 1] as const)
    test(`concurrent complete merge acceptance lasts until both actors undo; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await fixture(clientIds);
      try {
        const [alice, bob] = pair.editors;
        propose(alice!, rectangle);
        await pair.settle();
        append(bob!, 'I', ' Bob');
        await pair.settle();
        const pending = groups(alice!).map((group) => group.id);
        expect(pending).toHaveLength(1);
        const [committed, working] = await same(pair);
        expect(working.replace(/<br>/g, ' ')).toContain('A B D E');
        expect(working).toContain('I Bob');
        pair.pause();
        for (const editor of pair.editors) expect(acceptSelected(editor).ok).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        expect(await same(pair)).toEqual([working, working]);
        for (const editor of pair.editors) expect(groups(editor)).toEqual([]);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([working, working]);
        for (const editor of pair.editors) expect(groups(editor)).toEqual([]);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([committed, working]);
        for (const editor of pair.editors)
          expect(groups(editor).map((group) => group.id)).toEqual(pending);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await same(pair)).toEqual([working, working]);
        for (const editor of pair.editors) {
          expect(groups(editor)).toEqual([]);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(readReviewGroups(saved.package)).toEqual({ success: true, groups: [] });
        }
        const [savedAlice, savedBob] = await pair.savedDigests();
        expect(savedAlice).toEqual(savedBob);
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 60_000);
