import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  parentNodeOf,
  paragraphTextOf,
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
import { recomposeOwnedProposals } from './recompose';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
type Wave = Extract<Awaited<ReturnType<typeof recomposeOwnedProposals>>, { success: true }>;
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
const markdown = (rows: readonly (readonly string[])[]) =>
  '**Alpha XYZ end**\n\n' +
  rows
    .map(
      (row, i) =>
        '| ' +
        row.join(' | ') +
        ' |' +
        (i === 0 ? '\n| ' + row.map(() => '---').join(' | ') + ' |' : '')
    )
    .join('\n') +
  '\n\nTail';
const scenarios: readonly [string, readonly (readonly string[])[], string?, (() => Uint8Array)?][] =
  [
    [
      'insert-row',
      [
        ['H1', 'H2'],
        ['A', 'B'],
        ['New **bold**', '`code` [link](https://example.com/new)'],
        ['C', 'D'],
      ],
    ],
    [
      'delete-row',
      [
        ['H1', 'H2'],
        ['C', 'D'],
      ],
    ],
    [
      'insert-column',
      [
        ['H1', 'New header', 'H2'],
        ['A', 'New **bold**', 'B'],
        ['C', '`code` [link](https://example.com/new)', 'D'],
      ],
    ],
    ['delete-column', [['H1'], ['A'], ['C']]],
    [
      'merge',
      [
        ['H1', 'H2'],
        ['A', 'B'],
        ['C', 'D'],
      ],
      '**Alpha XYZ end**\n\n<table>\n<tr>\n<td colspan="2">\n\nNew title\n\n</td>\n</tr>\n<tr>\n<td>\n\nA\n\n</td>\n<td>\n\nB\n\n</td>\n</tr>\n<tr>\n<td>\n\nC\n\n</td>\n<td>\n\nD\n\n</td>\n</tr>\n</table>\n\nTail',
    ],
    [
      'split',
      [
        ['Head', 'New **bold**'],
        ['A', 'B'],
        ['C', 'D'],
      ],
      undefined,
      () =>
        docxFixture(
          p('Alpha end') +
            '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:tcW w:w="4800" w:type="dxa"/></w:tcPr>' +
            p('Head') +
            '</w:tc></w:tr><w:tr>' +
            cell('A') +
            cell('B') +
            '</w:tr><w:tr>' +
            cell('C') +
            cell('D') +
            '</w:tr></w:tbl>' +
            p('Tail')
        ),
    ],
  ];
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
async function wave(editor: DocxEditorInstance, targetMarkdown: string, previous?: Wave) {
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
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision,
        targetMarkdown,
        ownership: previous?.ownership ?? [],
        contentProofs: previous?.contentProofs ?? [],
      },
      { ...context, epoch: revision.epoch }
    );
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

function tableGroup(editor: DocxEditorInstance) {
  const part = editor.surface!.session.part();
  return groups(editor).find((group) =>
    group.native.some((member) =>
      member.siteNodeIds.some((id) => {
        let node = findNode(part, id);
        for (let depth = 0; node && depth < 64; depth++) {
          if (node.kind === 'table') return true;
          node = parentNodeOf(part, node.id);
        }
        return false;
      })
    )
  );
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const [kind, rows, targetOverride, sourceOverride] of scenarios)
      for (const decision of ['accept', 'reject'] as const) {
        test(`mixed ${kind} waves preserve peer cells, both histories, ${decision}, geometry, and reopening; ${clientIds}/${receiver}`, async () => {
          const pair = await createTwoEditorFixture({
            ...options(),
            clientIds,
            ...(sourceOverride ? { document: sourceOverride() } : {}),
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
            const peer = alice
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(alice.surface!.session.part(), id) === 'C')!;
            const original = (await views(alice))[0]!,
              target = targetOverride ?? markdown(rows),
              first = await wave(alice, target);
            await pair.settle();
            pair.pause();
            const next = target
              .replace('**Alpha XYZ end**', '*Alpha Q end*')
              .replace(/New/g, 'Final');
            const second = await wave(alice, next, first);
            append(bob, peer, ' Peer');
            pair.reconnect(receiver);
            await pair.settle();
            const before = original.replace(/\bC\b/, 'C Peer'),
              working = next.replace(/\bC\b/, 'C Peer');
            await check(pair, before, working);
            await geometry(pair, peer);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, target.replace(/\bC\b/, 'C Peer'));
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, working);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, original, next);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, working);
            const selected = tableGroup(bob);
            expect(selected).toBeDefined();
            const decided = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              bob,
              selected!.id,
              groupToken(bob),
              groupToken(bob)
            );
            if (!decided.success) throw new Error(decided.error);
            await pair.settle();
            const committed =
              decision === 'accept' ? working.replace('*Alpha Q end*', 'Alpha end') : before;
            const proposed =
              decision === 'accept' ? working : before.replace('Alpha end', '*Alpha Q end*');
            await check(pair, committed, proposed);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, working);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, committed, proposed);
            const finalTarget = working.replace('Alpha Q', 'Alpha R');
            await wave(alice, finalTarget, second);
            await pair.settle();
            await check(pair, committed, finalTarget);
            await saved(pair);
          } finally {
            pair.destroy();
          }
        }, 30_000);
      }

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const owner of ['row', 'table'] as const) {
      test(`deleting the old ${owner} owner offline preserves a later wave and peer history; ${clientIds}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const rows = scenarios.find(
            ([kind]) => kind === (owner === 'row' ? 'insert-row' : 'insert-column')
          )![1];
          const original = (await views(alice))[0]!,
            target = markdown(rows),
            first = await wave(alice, target);
          await pair.settle();
          const session = bob.surface!.session,
            inside = session
              .paragraphIds()
              .find((id) => paragraphTextOf(session.part(), id) === 'New bold')!;
          let row = findNode(session.part(), inside);
          while (row && row.kind !== 'tableRow') row = parentNodeOf(session.part(), row.id);
          let table = findNode(session.part(), inside);
          while (table && table.kind !== 'table') table = parentNodeOf(session.part(), table.id);
          expect(row?.kind).toBe('tableRow');
          expect(table?.kind).toBe('table');
          const peer = session
            .paragraphIds()
            .find(
              (id) => paragraphTextOf(session.part(), id) === (owner === 'row' ? 'C' : 'Tail')
            )!;
          pair.pause();
          const next = target
            .replace('**Alpha XYZ end**', '*Alpha Q end*')
            .replace(/New/g, 'Final');
          await wave(alice, next, first);
          expect(
            bob.surface!.applyAutomationOps(
              () => [
                owner === 'row'
                  ? { op: 'deleteTableRow', tableId: table!.id, rowId: row!.id }
                  : { op: 'deleteBlock', blockId: table!.id },
              ],
              { kind: 'body' },
              [],
              false,
              true
            ).committed
          ).toBe(true);
          append(bob, peer, ' Peer');
          pair.reconnect(receiver);
          await pair.settle();
          const committed =
            owner === 'row' ? original.replace(/\bC\b/, 'C Peer') : 'Alpha end\n\nTail Peer';
          const proposed =
            owner === 'row' ? next.replace(/\bC\b/, 'C Peer') : '*Alpha Q end*\n\nTail Peer';
          await check(pair, committed, proposed);
          await geometry(pair, peer);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, committed.replace('Alpha end', '**Alpha XYZ end**'));
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed, proposed);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed.replace(' Peer', ''), proposed.replace(' Peer', ''));
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, original, next);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, committed.replace(' Peer', ''), proposed.replace(' Peer', ''));
          await saved(pair);
        } finally {
          pair.destroy();
        }
      }, 30_000);
    }
