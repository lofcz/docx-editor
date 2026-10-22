import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  paragraphTextOf,
  storyParagraphs,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { nodeChildren, rowCells, tableRows } from '../../test/support/native-table-oracle';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';

const p = (value: string) => `<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;
const BLANK = '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p/></w:tc>';
const tc = (text: string, extra = '', width = 2400) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const grid = (rows: string) =>
  `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${rows}</w:tbl>`;
const context = { documentId: 'merge-collaboration', epoch: 'mount' };

async function plan(editor: DocxEditorInstance, target: string) {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  const planned = planTargetDocument(
    {
      revision: token,
      pkg: editor.surface!.session.currentPackage(),
      projection: view.view.projection,
    },
    target
  );
  if (!planned.success) throw new Error(planned.error);
  return planned.plan;
}
/** Visual cell texts with their native span/merge shape. */
function layout(part: OoxmlPart) {
  const table = nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
  return tableRows(part, table.id).map((row) =>
    rowCells(row).map((cell) => storyParagraphs(cell).map((node) => paragraphTextOf(part, node.id)))
  );
}
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));

interface Case {
  readonly name: string;
  readonly body: string;
  readonly target: string;
  readonly peerText: string;
  readonly before: string[][][];
  readonly after: string[][][];
}
const CASES: readonly Case[] = [
  {
    name: 'merge keeps a peer edit inside the moved covered paragraph',
    body:
      grid(
        `<w:tr>${tc('A')}${tc('B')}${tc('C')}</w:tr><w:tr>${BLANK}${tc('Two')}${tc('Three')}</w:tr>`
      ) + p('Peer.'),
    target:
      '<table>\n<tr>\n<td>\n\nA\n\n</td>\n<td>\n\nB\n\n</td>\n<td>\n\nC\n\n</td>\n</tr>\n<tr>\n<td colspan="2">\n\nTwo\n\n</td>\n<td>\n\nThree\n\n</td>\n</tr>\n</table>\n\nPeer.',
    peerText: 'Two',
    before: [
      [['A'], ['B'], ['C']],
      [[''], ['Two Bob'], ['Three']],
    ],
    after: [
      [['A'], ['B'], ['C']],
      [['Two Bob'], ['Three']],
    ],
  },
  {
    name: 'split fills new cells beside a peer edit in a neighbor cell',
    body:
      grid(
        `<w:tr>${tc('AB', '<w:gridSpan w:val="2"/>', 4800)}${tc('C1')}</w:tr><w:tr>${tc('A2')}${tc('B2')}${tc('C2')}</w:tr>`
      ) + p('Peer.'),
    target: '| A | B | C1 |\n| --- | --- | --- |\n| A2 | B2 | C2 |\n\nPeer.',
    peerText: 'C1',
    before: [
      [['AB'], ['C1 Bob']],
      [['A2'], ['B2'], ['C2']],
    ],
    after: [
      [['A'], ['B'], ['C1 Bob']],
      [['A2'], ['B2'], ['C2']],
    ],
  },
];

for (const item of CASES)
  for (const first of [0, 1] as const)
    test(`two sessions converge when a direct ${item.name}; receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(item.body),
        documentId: `merge-collaboration-${first}`,
        clientIds: [73, 41],
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          planned = await plan(alice, item.target),
          part = bob.surface!.session.part(),
          edited = bob
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(part, id) === item.peerText)!,
          peerId = bob.surface!.session.paragraphIds().at(-1)!;
        pair.pause();
        const result = applyDirectPlan(alice, planned, context);
        if (!result.success) throw new Error(result.error);
        expect(
          bob.surface!.applyAutomationOps(() => [
            {
              op: 'insertText',
              paragraphId: edited,
              offset: item.peerText.length,
              text: ' Bob',
            },
          ]).committed
        ).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        for (const editor of pair.editors)
          expect(layout(editor.surface!.session.part())).toEqual(item.after);

        // Remote caret geometry after the topology change.
        bob.surface!.setSelection({
          anchor: { paragraphId: peerId, offset: 4 },
          head: { paragraphId: peerId, offset: 4 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote.head.nodeId).toBe(peerId);
        const caret = caretAt(alice.surface!.layout(), { paragraphId: peerId, offset: 4 }),
          start = caretAt(alice.surface!.layout(), { paragraphId: peerId, offset: 0 });
        expect(caret!.x).toBeGreaterThan(start!.x);

        // Undo restores the original topology and keeps Bob's text; redo reapplies it.
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(layout(editor.surface!.session.part())).toEqual(item.before);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(layout(editor.surface!.session.part())).toEqual(item.after);

        // Peer deletion during concurrent editing still converges.
        pair.pause();
        expect(
          bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: peerId }]).committed
        ).toBe(true);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        expect(texts(alice)).toEqual(texts(bob));
        expect(texts(alice)).not.toContain('Peer.');
        expect(layout(alice.surface!.session.part())).toEqual(item.before);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    }, 60_000);
