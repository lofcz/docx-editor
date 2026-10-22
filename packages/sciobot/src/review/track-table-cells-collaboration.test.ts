import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  parentNodeOf,
  propertyContainer,
  runsCovering,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  CELL_CONTEXT,
  CELL_ORIGINAL,
  cellPlan,
  cellReviewDocument,
  cellScopeFor,
  cellViews,
  scopeOnly,
} from '../../test/support/table-review';
import { directParagraphTextOps, planParagraphText } from '../diff/paragraph-text';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableCellPlan } from './track-table-cells';

const fixture = (clientIds?: readonly [number, number]) =>
  createTwoEditorFixture({
    document: cellReviewDocument(),
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    ownedAttributes: true,
    ...(clientIds ? { clientIds } : {}),
  });
const bold = [{ start: 0, end: 3, target: { properties: [{ localName: 'b' }] } }];

for (const action of ['accept', 'reject'] as const)
  test(`cell text/mark proposals and ${action} converge with remote caret geometry, undo/redo, and save/reopen`, async () => {
    const pair = await fixture();
    try {
      const [alice, bob] = pair.editors;
      const scope = cellScopeFor(alice!);
      bob!.surface!.setSelection({
        anchor: { paragraphId: scope.paragraphId, offset: 3 },
        head: { paragraphId: scope.paragraphId, offset: 3 },
      });
      await pair.settle();
      const caretBefore = caretAt(alice!.surface!.layout(), {
        paragraphId: scope.paragraphId,
        offset: 3,
      });
      expect(caretBefore !== null).toBe(true);
      const plan = await cellPlan(alice!, scope, [
        { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
      ]);
      const result = trackTableCellPlan(alice!, plan, scopeOnly(scope), CELL_CONTEXT);
      if (!result.success) throw new Error(result.error);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(await cellViews(editor)).toEqual([
          CELL_ORIGINAL,
          CELL_ORIGINAL.replace('Two', '**Changed**'),
        ]);
        expect(editor.getReviewItems().some((item) => item.kind === 'revision')).toBe(true);
      }
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(Boolean(remote)).toBe(true);
      expect(remote.head.nodeId).toBe(scope.paragraphId);
      expect(remote.head.offset).toBe(bob!.surface!.state().selection!.head.offset);
      const caretAfter = caretAt(alice!.surface!.layout(), {
        paragraphId: scope.paragraphId,
        offset: remote.head.offset,
      });
      expect(caretAfter !== null).toBe(true);
      expect(remote.head.offset).toBe(10); // The old-text endpoint follows the leading seven-character insertion.
      expect(
        paragraphTextOf(alice!.surface!.session.part(), scope.paragraphId)!.slice(
          0,
          remote.head.offset
        )
      ).toBe('ChangedTwo');
      const bobCaret = caretAt(bob!.surface!.layout(), {
        paragraphId: scope.paragraphId,
        offset: remote.head.offset,
      });
      expect(bobCaret !== null).toBe(true);
      expect({ x: caretAfter!.x, y: caretAfter!.y }).toEqual({ x: bobCaret!.x, y: bobCaret!.y });
      expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
      const [pendingAlice, pendingBob] = await pair.savedDigests();
      expect(pendingAlice).toEqual(pendingBob);
      expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      await pair.settle();
      const expected =
        action === 'accept' ? CELL_ORIGINAL.replace('Two', '**Changed**') : CELL_ORIGINAL;
      for (const editor of pair.editors) {
        expect(await cellViews(editor)).toEqual([expected, expected]);
        expect(editor.getReviewItems()).toHaveLength(0);
      }
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await cellViews(editor)).toEqual([
          CELL_ORIGINAL,
          CELL_ORIGINAL.replace('Two', '**Changed**'),
        ]);
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await cellViews(editor)).toEqual([expected, expected]);
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 30_000);

for (const peerCell of ['One', 'Two'])
  for (const marked of peerCell === 'Two' ? [false, true] : [false])
    for (const clientIds of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const first of [0, 1] as const)
        for (const action of ['accept', 'reject'] as const)
          test(`offline cell text/marks retain peer content in ${peerCell} through reconnect, proposal history, and ${action} (${marked ? 'italic' : 'plain'}); clients ${clientIds.join('/')}, receiver ${first}`, async () => {
            const pair = await fixture(clientIds);
            const stops: (() => void)[] = [];
            try {
              const [alice, bob] = pair.editors;
              const scope = cellScopeFor(alice!),
                peer = cellScopeFor(bob!, peerCell);
              const plan = await cellPlan(alice!, scope, [
                { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
              ]);
              bob!.surface!.setSelection({
                anchor: { paragraphId: peer.paragraphId, offset: 3 },
                head: { paragraphId: peer.paragraphId, offset: 3 },
              });
              await pair.settle();
              pair.pause();
              expect(trackTableCellPlan(alice!, plan, scopeOnly(scope), CELL_CONTEXT).success).toBe(
                true
              );
              const part = bob!.surface!.session.part();
              const direct = planParagraphText(
                {
                  revision: readEditorRevision(bob!, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
                  partName: part.name,
                  paragraphId: peer.paragraphId,
                  text: paragraphTextOf(part, peer.paragraphId)!,
                },
                peerCell + ' Bob'
              );
              if (!direct.success) throw new Error(direct.error);
              const ops = direct.plan.operations.flatMap((edit) =>
                edit.kind === 'text' ? directParagraphTextOps(edit) : []
              );
              expect(
                bob!.surface!.applyAutomationOps(() => [
                  ...ops,
                  ...(marked
                    ? [
                        {
                          op: 'setRunProperties' as const,
                          paragraphId: peer.paragraphId,
                          start: 4,
                          end: 7,
                          properties: [{ localName: 'i' }],
                        },
                      ]
                    : []),
                ]).rejected
              ).toBe(false);
              const foreign: string[] = [];
              let publications = 0;
              for (const editor of pair.editors)
                stops.push(
                  editor.on('change', () => {
                    publications++;
                    const current = editor.surface!.session.part(),
                      text = paragraphTextOf(current, peer.paragraphId) ?? '',
                      start = text.lastIndexOf('Bob');
                    if (start < 0) return;
                    for (const run of runsCovering(current, peer.paragraphId, start, start + 3)) {
                      const properties = propertyContainer(run, 'runProperties', 'rPr');
                      if (
                        properties &&
                        properties.kind !== 'textValue' &&
                        properties.children.some(
                          (child) =>
                            child.kind !== 'textValue' &&
                            ['b', 'rPrChange'].includes(child.localName)
                        )
                      )
                        foreign.push(run.id);
                      for (
                        let owner = parentNodeOf(current, run.id);
                        owner && owner.kind !== 'paragraph';
                        owner = parentNodeOf(current, owner.id)
                      )
                        if (owner.kind === 'revisionInsert' || owner.kind === 'revisionDelete')
                          foreign.push(owner.id);
                    }
                  })
                );
              pair.reconnect(first);
              await pair.settle();
              const original = CELL_ORIGINAL.replace(
                peerCell,
                peerCell + (marked ? ' *Bob*' : ' Bob')
              );
              const working = original.replace('Two', '**Changed**');
              for (const editor of pair.editors)
                expect(await cellViews(editor)).toEqual([original, working]);
              const remote = pair.rooms[0]!.session.remoteSelections().find(
                (selection) => selection.actorId === 'bob'
              );
              expect(remote).toBeDefined();
              expect(remote!.head.offset).toBe(bob!.surface!.state().selection!.head.offset);
              const position = { paragraphId: remote!.head.nodeId, offset: remote!.head.offset };
              const leftCaret = caretAt(alice!.surface!.layout(), position),
                rightCaret = caretAt(bob!.surface!.layout(), position);
              expect(leftCaret).not.toBeNull();
              expect(rightCaret).not.toBeNull();
              expect({ x: leftCaret!.x, y: leftCaret!.y }).toEqual({
                x: rightCaret!.x,
                y: rightCaret!.y,
              });
              const pending = await pair.savedDigests();
              expect(pending[0]).toEqual(pending[1]);
              if (peerCell === 'Two') {
                const reopened = await createTwoEditorFixture({
                  document: cellReviewDocument(),
                  ownedAttributes: true,
                  roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
                  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
                });
                try {
                  for (const editor of reopened.editors)
                    expect(await cellViews(editor)).toEqual([original, working]);
                  expect(
                    reopened.editors[0]!.exec({ type: 'resolveAllReviewChanges', action }).ok
                  ).toBe(true);
                  await reopened.settle();
                  const resolved = action === 'accept' ? working : original;
                  for (const editor of reopened.editors)
                    expect(await cellViews(editor)).toEqual([resolved, resolved]);
                } finally {
                  reopened.destroy();
                }
              }
              expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors) {
                expect(await cellViews(editor)).toEqual([original, original]);
                expect(editor.getReviewItems()).toHaveLength(0);
              }
              expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors)
                expect(await cellViews(editor)).toEqual([original, working]);
              const decision = alice!.exec({ type: 'resolveAllReviewChanges', action });
              if (!decision.ok) throw new Error(JSON.stringify(decision));
              expect(decision.ok).toBe(true);
              await pair.settle();
              const resolved = action === 'accept' ? working : original;
              for (const editor of pair.editors)
                expect(await cellViews(editor)).toEqual([resolved, resolved]);
              expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors)
                expect(await cellViews(editor)).toEqual([original, working]);
              expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
              await pair.settle();
              for (const editor of pair.editors)
                expect(await cellViews(editor)).toEqual([resolved, resolved]);
              expect(publications).toBeGreaterThan(0);
              expect(foreign).toEqual([]);
              pair.ownership.forEach((handle) => handle.assertHealthy());
              const [savedAlice, savedBob] = await pair.savedDigests();
              expect(savedAlice).toEqual(savedBob);
            } finally {
              stops.forEach((stop) => stop());
              pair.destroy();
            }
          }, 30_000);

test('row deletion converges while a peer proposes text/marks in that row; stale plans cannot recreate it', async () => {
  const pair = await fixture();
  try {
    const [alice, bob] = pair.editors;
    const scope = cellScopeFor(bob!);
    const plan = await cellPlan(bob!, scope, [
      { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
    ]);
    pair.pause();
    expect(trackTableCellPlan(bob!, plan, scopeOnly(scope), CELL_CONTEXT).success).toBe(true);
    expect(
      alice!.surface!.applyAutomationOps(() => [
        { op: 'deleteTableRow', tableId: scope.tableId, rowId: scope.rowId },
      ]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    for (const editor of pair.editors) {
      expect(paragraphTextOf(editor.surface!.session.part(), scope.paragraphId)).toBeNull();
      expect(editor.getReviewItems()).toHaveLength(0);
    }
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(await cellViews(editor)).toEqual([
        CELL_ORIGINAL,
        CELL_ORIGINAL.replace('Two', '**Changed**'),
      ]);
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(paragraphTextOf(editor.surface!.session.part(), scope.paragraphId)).toBeNull();
    const before = canonicalOoxmlFingerprint(bob!.surface!.session.part());
    const result = trackTableCellPlan(bob!, plan, scopeOnly(scope), CELL_CONTEXT);
    expect(result.success).toBe(false);
    expect(!result.success && result.code).toBe('stale-revision');
    expect(canonicalOoxmlFingerprint(bob!.surface!.session.part())).toBe(before);
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
  } finally {
    pair.destroy();
  }
}, 30_000);
