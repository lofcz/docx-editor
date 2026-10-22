import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  parentNodeOf,
  type OoxmlNode,
  readOoxmlPackage,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import { groupToken, groups } from '../../test/support/review-groups';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { recomposeOwnedProposals } from './recompose';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import {
  replaceOwnedProposals,
  type OwnedProposalReplacementResult,
} from './replace-owned-proposals';
import { createEditPlan } from '../diff/plan';
import { applyDirectPlan } from '../ai/apply-direct';

const children = (node: OoxmlNode | null | undefined): readonly OoxmlNode[] =>
  node && node.kind !== 'textValue' ? node.children : [];
const context = { ...propertyContext, date: '2026-10-21T18:00:00.000Z' };
type Success = Extract<OwnedProposalReplacementResult, { success: true }>;
type Editor = ReturnType<typeof propertyEditor>;
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const table_ = () =>
  '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  [
    ['H1', 'H2'],
    ['A', 'B'],
    ['C', 'D'],
  ]
    .map((cells) => '<w:tr>' + cells.map(cell).join('') + '</w:tr>')
    .join('') +
  '</w:tbl>';
const document_ = () => docxFixture(p('Alpha end') + table_() + p('Tail'));
const pkg = (editor: Editor) => editor.surface!.session.currentPackage();
const fingerprints = (editor: Editor) =>
  [...pkg(editor).parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);
function ast(markdown: string) {
  const read = parseTargetMarkdown(markdown);
  if (!read.success) throw new Error(read.error);
  return targetAstKey(read.target.normalizedAst);
}
async function views(editor: Editor) {
  const token = groupToken(editor),
    session = editor.surface!.session;
  const original = await readCommittedView(session, token, { tableSyntax: 'auto' }),
    proposed = await readWorkingView(session, token, { tableSyntax: 'auto' });
  if (!original.success || !proposed.success) throw new Error('Missing table review views.');
  return [original.view.projection.text, proposed.view.projection.text];
}
async function wave(editor: Editor, targetMarkdown: string, previous?: Success) {
  let publications = 0;
  const stop = editor.surface!.session.subscribe(() => publications++);
  try {
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: groupToken(editor),
        targetMarkdown,
        ownership: previous?.ownership ?? [],
        contentProofs: previous?.contentProofs ?? [],
      },
      context
    );
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect(result.committed).toBe(true);
    expect(publications).toBe(1);
    return result;
  } finally {
    stop();
  }
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
const scenarios = [
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
] as const;
for (const [kind, rows] of scenarios)
  for (const decision of ['accept', 'reject'] as const) {
    test(`mixed ${kind} waves preserve native history, resources, reopening, and ${decision}`, async () => {
      const loaded = readOoxmlPackage(document_());
      if (!loaded.ok) throw new Error(loaded.reason);
      const editor = propertyEditor(loaded.package);
      try {
        const original = (await views(editor))[0]!,
          before = fingerprints(editor),
          target = markdown(rows);
        const first = await wave(editor, target);
        expect((await views(editor)).map(ast)).toEqual([original, target].map(ast));
        const after = fingerprints(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(after);
        const next = target.replace('Alpha XYZ', 'Alpha Q').replace(/New/g, 'Final');
        await wave(editor, next, first);
        expect((await views(editor)).map(ast)).toEqual([original, next].map(ast));
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprints(editor)).toEqual(after);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        const read = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!read.ok) throw new Error(read.reason);
        const fresh = propertyEditor(read.package);
        try {
          expect((await views(fresh)).map(ast)).toEqual([original, next].map(ast));
          expect(groups(fresh).map((g) => g.id)).toEqual(groups(editor).map((g) => g.id));
        } finally {
          fresh.destroy();
        }
        for (const group of [...groups(editor)]) {
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            group.id,
            groupToken(editor),
            groupToken(editor)
          );
          if (!result.success) throw new Error(result.error);
        }
        const expected = decision === 'accept' ? next : original;
        expect((await views(editor)).map(ast)).toEqual([expected, expected].map(ast));
      } finally {
        editor.destroy();
      }
    }, 30_000);
  }

for (const kind of ['insert-row', 'insert-column'] as const)
  for (const manual of ['text', 'paragraph'] as const) {
    test(`manual ${manual} inside a new ${kind} invalidates ownership before any inverse`, async () => {
      const loaded = readOoxmlPackage(document_());
      if (!loaded.ok) throw new Error(loaded.reason);
      const editor = propertyEditor(loaded.package);
      try {
        const scenario = scenarios.find(([name]) => name === kind)!;
        const target = markdown(scenario[1]),
          first = await wave(editor, target),
          session = editor.surface!.session;
        const paragraphId = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === 'New bold')!;
        expect(paragraphId).toBeDefined();
        expect(
          editor.surface!.applyAutomationOps(
            () => [
              manual === 'text'
                ? { op: 'insertText', paragraphId, offset: 8, text: ' manual', bias: 'right' }
                : {
                    op: 'setParagraphProperties',
                    paragraphId,
                    properties: [{ localName: 'jc', attributes: { val: 'center' } }],
                  },
            ],
            { kind: 'body' },
            [],
            false,
            true
          ).committed
        ).toBe(true);
        const before = pkg(editor),
          token = groupToken(editor),
          prior = groups(editor);
        let publications = 0;
        const stop = session.subscribe(() => publications++);
        try {
          expect(
            await recomposeOwnedProposals(
              editor,
              {
                revision: token,
                targetMarkdown: target.replace(/New/g, 'Final'),
                ownership: first.ownership,
                contentProofs: first.contentProofs,
              },
              context
            )
          ).toMatchObject({ success: false, code: 'ambiguous-target' });
          expect(pkg(editor)).toBe(before);
          expect(groupToken(editor)).toEqual(token);
          expect(groups(editor)).toEqual(prior);
          expect(publications).toBe(0);
        } finally {
          stop();
        }
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        await wave(editor, target.replace(/New/g, 'Final'), first);
      } finally {
        editor.destroy();
      }
    }, 30_000);
  }

const html = (rows: readonly (readonly string[])[]) =>
  '<table>\n' +
  rows.map((cells) => '<tr>\n' + cells.join('\n') + '\n</tr>').join('\n') +
  '\n</table>';
const td = (text: string, attributes = '') => `<td${attributes}>\n\n${text}\n\n</td>`;
for (const decision of ['accept', 'reject'] as const) {
  test(`mixed merge with edited head and covered text preserves both views and ${decision}`, async () => {
    const loaded = readOoxmlPackage(document_());
    if (!loaded.ok) throw new Error(loaded.reason);
    const editor = propertyEditor(loaded.package);
    try {
      const original = (await views(editor))[0]!,
        target =
          '**Alpha XYZ end**\n\n' +
          html([[td('Title', ' colspan="2"')], [td('A'), td('B')], [td('C'), td('D')]]) +
          '\n\nTail';
      const first = await wave(editor, target);
      expect((await views(editor)).map(ast)).toEqual([original, target].map(ast));
      const next = target.replace('Title', 'Final').replace('Alpha XYZ', 'Alpha Q');
      await wave(editor, next, first);
      expect((await views(editor)).map(ast)).toEqual([original, next].map(ast));
      for (const group of [...groups(editor)]) {
        const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
          editor,
          group.id,
          groupToken(editor),
          groupToken(editor)
        );
        if (!result.success) throw new Error(result.error);
      }
      const expected = decision === 'accept' ? next : original;
      expect((await views(editor)).map(ast)).toEqual([expected, expected].map(ast));
    } finally {
      editor.destroy();
    }
  }, 30_000);
}

for (const decision of ['accept', 'reject'] as const) {
  test(`a mixed split fills actual native cells and supports repeated waves and ${decision}`, async () => {
    const input = docxFixture(
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
    );
    const loaded = readOoxmlPackage(input);
    if (!loaded.ok) throw new Error(loaded.reason);
    const editor = propertyEditor(loaded.package);
    try {
      const original = (await views(editor))[0]!,
        before = fingerprints(editor),
        target = markdown([
          ['Head', 'Split **bold**'],
          ['A', 'B'],
          ['C', 'D'],
        ]);
      const first = await wave(editor, target);
      expect((await views(editor)).map(ast)).toEqual([original, target].map(ast));
      const after = fingerprints(editor);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprints(editor)).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprints(editor)).toEqual(after);
      const next = target.replace('Split', 'Final').replace('Alpha XYZ', 'Alpha Q');
      await wave(editor, next, first);
      expect((await views(editor)).map(ast)).toEqual([original, next].map(ast));
      for (const group of [...groups(editor)]) {
        const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
          editor,
          group.id,
          groupToken(editor),
          groupToken(editor)
        );
        if (!result.success) throw new Error(result.error);
      }
      const expected = decision === 'accept' ? next : original;
      expect((await views(editor)).map(ast)).toEqual([expected, expected].map(ast));
    } finally {
      editor.destroy();
    }
  }, 30_000);
}

for (const command of [
  'fill',
  'width',
  'borders',
  'alignment',
  'row-height',
  'header',
  'divider',
  'right-edge',
] as const) {
  test(`mixed ${command} intentions retain independent table decisions and match direct results`, async () => {
    const read = readOoxmlPackage(docxFixture(p('Alpha end') + table_() + table_() + p('Tail')));
    if (!read.ok) throw new Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const part = editor.surface!.session.part(),
        body = part.root.children.find(
          (node) => node.kind !== 'textValue' && node.localName === 'body'
        )!;
      if (body.kind === 'textValue') throw new Error('Missing native body.');
      const tables = children(body).filter((node) => node.kind === 'table');
      expect(tables).toHaveLength(2);
      const operations = [
        {
          kind: 'text' as const,
          partName: part.name,
          paragraphId: editor.surface!.session.paragraphIds()[0]!,
          start: 6,
          end: 6,
          before: '',
          after: 'Q ',
        },
        ...tables.map((table) => {
          const row = children(table).find((node) => node.kind === 'tableRow')!;
          if (row.kind !== 'tableRow') throw new Error('Missing native row.');
          const cell = children(row).find((node) => node.kind === 'tableCell')!;
          const grid = children(table).find(
            (node) => node.kind !== 'textValue' && node.localName === 'tblGrid'
          );
          const columns = children(grid);
          return {
            kind: 'table' as const,
            partName: part.name,
            operation:
              command === 'fill'
                ? {
                    op: 'setTableCellFill' as const,
                    tableId: table.id,
                    cellIds: [cell.id],
                    color: { kind: 'hex' as const, value: '00FF00' },
                  }
                : command === 'width'
                  ? {
                      op: 'setTableProperties' as const,
                      tableId: table.id,
                      columnWidthsTwips: [1800, 3000],
                    }
                  : command === 'borders'
                    ? {
                        op: 'setTableCellBorders' as const,
                        tableId: table.id,
                        cellIds: [cell.id],
                        scope: 'all' as const,
                        spec: {
                          style: 'double' as const,
                          size: 12,
                          color: { kind: 'hex' as const, value: '0000FF' },
                        },
                      }
                    : command === 'alignment'
                      ? {
                          op: 'setTableCellVerticalAlignment' as const,
                          tableId: table.id,
                          cellIds: [cell.id],
                          alignment: 'center' as const,
                        }
                      : command === 'row-height'
                        ? {
                            op: 'setTableRowHeight' as const,
                            tableId: table.id,
                            rowId: row.id,
                            heightTwips: 900,
                          }
                        : command === 'header'
                          ? {
                              op: 'setTableProperties' as const,
                              tableId: table.id,
                              headerRowCount: 1,
                            }
                          : command === 'divider'
                            ? {
                                op: 'setTableColumnWidths' as const,
                                tableId: table.id,
                                leftGridColumnId: columns[0]!.id,
                                rightGridColumnId: columns[1]!.id,
                                leftWidthTwips: 1800,
                                rightWidthTwips: 3000,
                              }
                            : {
                                op: 'setTableRightEdgeWidth' as const,
                                tableId: table.id,
                                gridColumnId: columns[1]!.id,
                                columnWidthTwips: 3000,
                                tableWidthTwips: 5400,
                              },
          };
        }),
      ];
      const build = () => {
        const result = createEditPlan(groupToken(editor), operations);
        if (!result.success) throw new Error(result.error);
        return result.plan;
      };
      const before = fingerprints(editor);
      const direct = applyDirectPlan(editor, build(), context);
      if (!direct.success) throw new Error(direct.error);
      const expectedView = (await views(editor))[0]!;
      const expected = tables.map((table) =>
        canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), table.id)!)
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprints(editor)).toEqual(before);
      const proposed = replaceOwnedProposals(editor, [], build(), context);
      if (!proposed.success) throw new Error(proposed.error);
      const owner = (nodeId: string) => {
        let node = findNode(editor.surface!.session.part(), nodeId);
        for (let depth = 0; node && depth < 64; depth++) {
          if (node.kind === 'table') return node.id;
          node = parentNodeOf(editor.surface!.session.part(), node.id);
        }
        return null;
      };
      const own = tables.map((table) =>
        groups(editor).filter((group) =>
          group.native.some((member) => member.siteNodeIds.some((id) => owner(id) === table.id))
        )
      );
      expect(own.map((groups) => groups.length)).toEqual([1, 1]);
      expect(own[0]![0]!.id).not.toBe(own[1]![0]!.id);
      const accepted = acceptReviewGroup(
        editor,
        own[0]![0]!.id,
        groupToken(editor),
        groupToken(editor)
      );
      if (!accepted.success) throw new Error(accepted.error);
      expect(groups(editor).some((group) => group.id === own[1]![0]!.id)).toBe(true);
      for (const group of [...groups(editor)]) {
        const result = acceptReviewGroup(editor, group.id, groupToken(editor), groupToken(editor));
        if (!result.success) throw new Error(result.error);
      }
      expect(
        tables.map((table) =>
          canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), table.id)!)
        )
      ).toEqual(expected);
      expect((await views(editor)).map(ast)).toEqual([expectedView, expectedView].map(ast));
    } finally {
      editor.destroy();
    }
  }, 30_000);
}
