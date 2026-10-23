// Investigation fixtures for birth retention across merged geometry and later native journals.
import { expect, test } from 'bun:test';
import { createTwoEditorFixture } from '../../../../../packages/sciobot/test/support/collaboration';
import { docxFixture } from '../../../../../packages/sciobot/test/support/document';
import { table } from '../../../../../packages/sciobot/test/support/target-table-fixture';
import {
  nodeChildren,
  namedChild,
  rowCells,
  tableRows,
} from '../../../../../packages/sciobot/test/support/native-table-oracle';
import { bodyStoryRoot, paragraphTextOf } from '@docx-editor.dev/core/store';
import { prepareNativeDocument } from '../../../../../packages/sciobot/src/native/prepare-document';
import { readEditorRevision } from '../../../../../packages/sciobot/src/projection/revision-cache';
import { readWorkingView } from '../../../../../packages/sciobot/src/review/working-view';
import { planTargetDocument } from '../../../../../packages/sciobot/src/diff/plan-target';
import { applyDirectPlan } from '../../../../../packages/sciobot/src/ai/apply-direct';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
const context = { documentId: 'birth-geometry', epoch: 'native' };
async function apply(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    target
  );
  if (!planned.success) throw Error(planned.error);
  const applied = applyDirectPlan(editor, planned.plan, context);
  if (!applied.success) throw Error(applied.error);
}
const fixture = (body: string) => {
  const prepared = prepareNativeDocument(docxFixture(body));
  if (!prepared.success) throw Error(prepared.error);
  return createTwoEditorFixture({ document: prepared.document, clientIds: [41, 73] });
};
const typePeer = (editor: DocxEditorInstance) => {
  const p = editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id)?.startsWith('Tail'))!;
  const offset = paragraphTextOf(editor.surface!.session.part(), p)!.length;
  expect(
    editor.surface!.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId: p, offset, text: ' peerborn' }],
      { kind: 'body' },
      [],
      false,
      true
    ).rejected
  ).toBe(false);
};

test('a recovered middle column retains the widened original merged header', async () => {
  const body =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Peer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const pair = await fixture(body);
  try {
    const [alice, bob] = pair.editors;
    const td = (s: string, props = '') => `<td${props}>\n\n${s}\n\n</td>`;
    await apply(
      alice!,
      `<table>\n<tr>\n${td('Merged', ' colspan="3"')}\n</tr>\n<tr>\n${td('Keep')}\n${td('Tail $x$')}\n${td('Peer')}\n</tr>\n</table>`
    );
    await pair.settle();
    typePeer(bob!);
    await pair.settle();
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors) {
      const part = editor.surface!.session.part();
      const tbl = nodeChildren(bodyStoryRoot(part)!).find((n) => n.kind === 'table')!;
      const rows = tableRows(part, tbl.id);
      expect(rows).toHaveLength(2);
      expect(
        nodeChildren(namedChild(tbl, 'tblGrid')!).filter((n) => n.localName === 'gridCol')
      ).toHaveLength(3);
      const header = rowCells(rows[0]!);
      expect(header).toHaveLength(1);
      const span = namedChild(namedChild(header[0]!, 'tcPr')!, 'gridSpan')!;
      expect(
        span.kind === 'textValue' ? null : span.attributes.find((a) => a.localName === 'val')?.value
      ).toBe('3');
      expect(rowCells(rows[1]!)).toHaveLength(3);
    }
  } finally {
    pair.destroy();
  }
});

test('a later native row deletion cannot unlink a hidden birth awaiting peer text', async () => {
  const pair = await fixture(
    table([
      ['A', 'B'],
      ['Keep', 'Peer'],
    ])
  );
  try {
    const [alice, bob] = pair.editors;
    await apply(alice!, '| New | Tail $x$ |\n| - | - |\n| A | B |\n| Keep | Peer |');
    await pair.settle();
    pair.pause();
    typePeer(bob!);
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await apply(alice!, '| Keep | Peer |\n| - | - |');
    pair.reconnect(0);
    await pair.settle();
    for (const editor of pair.editors) {
      const part = editor.surface!.session.part();
      const tbl = nodeChildren(bodyStoryRoot(part)!).find((n) => n.kind === 'table')!;
      const rows = tableRows(part, tbl.id);
      expect(rows).toHaveLength(2);
      expect(
        rows.map((row) =>
          rowCells(row).map((cell) =>
            nodeChildren(cell)
              .filter((n) => n.kind === 'paragraph')
              .map((p) => paragraphTextOf(part, p.id))
              .join('\n')
          )
        )
      ).toEqual([
        ['', ' peerborn'],
        ['Keep', 'Peer'],
      ]);
    }
  } finally {
    pair.destroy();
  }
});
