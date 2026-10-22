import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { groupToken, groups } from '../../test/support/review-groups';
import { targetAstKey } from '../diff/target-key';
import { parseTargetMarkdown } from '../projection/parse-target';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { replaceOwnedProposals } from './replace-owned-proposals';
import { createEditPlan } from '../diff/plan';
import { propertyEditor } from '../../test/support/property-execution';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
type Wave = Extract<Awaited<ReturnType<typeof replaceOwnedProposals>>, { success: true }>;
const context = {
  documentId: 'property-records',
  epoch: 'load-1',
  author: 'Bot',
  date: '2026-10-21T14:00:00.000Z',
};
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const document_ = () =>
  docxFixture(
    p('Alpha end') +
      '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      [
        ['H1', 'H2'],
        ['A', 'B'],
        ['C', 'D'],
      ]
        .map((cells) => '<w:tr>' + cells.map(cell).join('') + '</w:tr>')
        .join('') +
      '</w:tbl>' +
      p('Tail')
  );
function options() {
  return {
    document: document_(),
    documentId: context.documentId,
    ownedAttributes: true,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  };
}
function ast(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
}
async function views(editor: DocxEditorInstance) {
  const [before, after] = await Promise.all([
    readCommittedView(editor.surface!.session, groupToken(editor), { tableSyntax: 'auto' }),
    readWorkingView(editor.surface!.session, groupToken(editor), { tableSyntax: 'auto' }),
  ]);
  if (!before.success || !after.success) throw new Error('Missing mixed shared views.');
  return [before.view.projection.text, after.view.projection.text];
}
async function wave(editor: DocxEditorInstance, anchor: Anchors, kind: Kind, previous?: Wave) {
  const revision = groupToken(editor);
  const publications: { origin: string; fingerprint: string }[] = [];
  const unsubscribe = editor.surface!.session.subscribe((change) => {
    const source = editor.surface!.session.currentPackage();
    publications.push({
      origin: change.origin,
      fingerprint: JSON.stringify(
        [...source.parts]
          .map(([name, part]) => {
            // OPC relationships form a keyed set. Room reflection can change their XML order.
            const canonical =
              name.endsWith('.rels') && part.root.kind === 'generic'
                ? {
                    ...part,
                    root: {
                      ...part.root,
                      children: [...part.root.children].sort((a, b) =>
                        canonicalOoxmlFingerprint(a).localeCompare(canonicalOoxmlFingerprint(b))
                      ),
                    },
                  }
                : part;
            return [name, canonicalOoxmlFingerprint(canonical)];
          })
          .sort(([a], [b]) => a!.localeCompare(b!))
      ),
    });
  });
  try {
    const built = createEditPlan(revision, program(anchor, kind));
    if (!built.success) throw new Error(built.error);
    const result = replaceOwnedProposals(editor, previous?.ownership ?? [], built.plan, {
      ...context,
      epoch: revision.epoch,
    });
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect(result.committed).toBe(true);
    expect(publications.filter((entry) => !entry.origin.endsWith('.remote'))).toHaveLength(1);
    expect(new Set(publications.map((entry) => entry.fingerprint)).size).toBe(1);
    return result;
  } finally {
    unsubscribe();
  }
}
function append(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const offset = paragraphTextOf(editor.surface!.session.part(), paragraphId)!.length;
  expect(
    editor.surface!.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId, offset, text, bias: 'right' }],
      { kind: 'body' },
      [],
      false,
      true
    )
  ).toMatchObject({ committed: true, rejected: false });
}
function unnested(editor: DocxEditorInstance) {
  const visit = (node: OoxmlNode, inside = false) => {
    if (node.kind === 'textValue') return;
    const contentRevision = node.kind === 'revisionInsert' || node.kind === 'revisionDelete';
    if (contentRevision) expect(inside).toBe(false);
    for (const child of node.children) visit(child, inside || contentRevision);
  };
  visit(editor.surface!.session.part().root);
  const sites = collectRevisionSites(editor.surface!.session.part());
  expect(new Set(sites.map((site) => site.node.id)).size).toBe(sites.length);
}
async function check(pair: Pair, original: string, proposed: string) {
  for (const editor of pair.editors) {
    expect((await views(editor)).map(ast)).toEqual([original, proposed].map(ast));
    unnested(editor);
  }
}
async function geometry(pair: Pair, id: string) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
  const offset = paragraphTextOf(bob.surface!.session.part(), id)!.length;
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((entry) => entry.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice.surface!.layout(), { paragraphId: id, offset });
  const b = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Pair) {
  const expected = await views(pair.editors[0]!);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const expectedGroups = groups(pair.editors[0]!).map((group) => group.id);
  for (const room of [false, true]) {
    const reopened = await createTwoEditorFixture({
      ...options(),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of reopened.editors) {
        expect((await views(editor)).map(ast)).toEqual(expected.map(ast));
        expect(groups(editor).map((group) => group.id)).toEqual(expectedGroups);
      }
    } finally {
      reopened.destroy();
    }
  }
}

type Kind = 'row' | 'column' | 'merge';
const kids = (node: OoxmlNode | undefined): readonly OoxmlNode[] =>
  node && node.kind !== 'textValue' ? node.children : [];
function anchors(editor: DocxEditorInstance) {
  const part = editor.surface!.session.part(),
    body = kids(part.root).find((node) => node.kind === 'body')!,
    table = kids(body).find((node) => node.kind === 'table')!,
    rows = kids(table).filter((node) => node.kind === 'tableRow'),
    grid = kids(table).find((node) => node.kind === 'tableGrid')!;
  return {
    partName: part.name,
    tableId: table.id,
    rowId: rows.at(-1)!.id,
    columnId: kids(grid).at(-1)!.id,
    cellIds: kids(rows[0])
      .filter((node) => node.kind === 'tableCell')
      .map((node) => node.id),
  };
}
type Anchors = ReturnType<typeof anchors>;
function program(a: Anchors, kind: Kind) {
  const edits = [
    {
      kind: 'table' as const,
      partName: a.partName,
      operation:
        kind === 'row'
          ? {
              op: 'insertTableRow' as const,
              tableId: a.tableId,
              rowId: a.rowId,
              where: 'below' as const,
            }
          : {
              op: 'insertTableColumn' as const,
              tableId: a.tableId,
              gridColumnId: a.columnId,
              where: 'right' as const,
            },
    },
    {
      kind: 'table' as const,
      partName: a.partName,
      operation: { op: 'mergeTableCells' as const, tableId: a.tableId, cellIds: a.cellIds },
    },
  ];
  return kind === 'merge' ? edits.slice(1) : edits;
}
async function golden(kind: Kind) {
  const read = readOoxmlPackage(document_());
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const ops = program(anchors(editor), kind).map((edit) => edit.operation);
    expect(
      editor.surface!.applyAutomationOps(() => ops, { kind: 'body' }, [], false, true, 'after')
        .committed
    ).toBe(true);
    return (await views(editor))[1]!;
  } finally {
    editor.destroy();
  }
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`native compound table stages preserve both histories, ${decision}, peer geometry, and saved reopening; ${clientIds}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            anchor = anchors(alice),
            original = (await views(alice))[0]!,
            row = await golden('row'),
            column = await golden('column');
          const peer = alice
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'C')!;
          const first = await wave(alice, anchor, 'row');
          await pair.settle();
          await check(pair, original, row);
          pair.pause();
          await wave(alice, anchor, 'column', first);
          append(bob, peer, ' Peer');
          pair.reconnect(receiver);
          await pair.settle();
          const before = original.replace(/\bC\b/, 'C Peer'),
            after = column.replace(/\bC\b/, 'C Peer');
          await check(pair, before, after);
          await geometry(pair, peer);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, before, row.replace(/\bC\b/, 'C Peer'));
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, before, after);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, original, column);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, before, after);
          expect(groups(bob)).toHaveLength(1);
          const selected = groups(bob)[0]!,
            decided = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              bob,
              selected.id,
              groupToken(bob),
              groupToken(bob)
            );
          if (!decided.success) throw new Error(decided.error);
          await pair.settle();
          const expected = decision === 'accept' ? after : before;
          await check(pair, expected, expected);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, before, after);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, expected, expected);
          await saved(pair);
        } finally {
          pair.destroy();
        }
      }, 30000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const owner of ['old-row', 'whole-table'] as const)
      test(`offline ${owner} deletion survives compound replacement and both histories; ${clientIds}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            anchor = anchors(alice),
            original = (await views(alice))[0]!,
            column = await golden('column'),
            merge = await golden('merge');
          const tail = alice
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'Tail')!;
          const first = await wave(alice, anchor, 'row');
          await pair.settle();
          const oldRow = kids(
            kids(bob.surface!.session.part().root).find((node) => node.kind === 'body')!
          ).find((node) => node.kind === 'table')!;
          const oldRowId = kids(oldRow)
            .filter((node) => node.kind === 'tableRow')
            .at(-1)!.id;
          pair.pause();
          await wave(alice, anchor, 'column', first);
          append(bob, tail, ' Peer');
          expect(
            bob.surface!.applyAutomationOps(
              () =>
                owner === 'old-row'
                  ? [{ op: 'deleteTableRow', tableId: anchor.tableId, rowId: oldRowId }]
                  : [{ op: 'deleteBlock', blockId: anchor.tableId }],
              { kind: 'body' },
              [],
              false,
              true,
              'after'
            ).committed
          ).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          const committed =
              owner === 'old-row'
                ? original.replace('Tail', 'Tail Peer')
                : 'Alpha end\n\nTail Peer',
            proposed = owner === 'old-row' ? column.replace('Tail', 'Tail Peer') : committed;
          await check(pair, committed, proposed);
          await geometry(pair, tail);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(
            pair,
            committed,
            owner === 'old-row' ? merge.replace('Tail', 'Tail Peer') : committed
          );
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, proposed);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(
            pair,
            original.replace('Tail', 'Tail Peer'),
            column.replace('Tail', 'Tail Peer')
          );
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, proposed);
          await saved(pair);
        } finally {
          pair.destroy();
        }
      }, 30000);
