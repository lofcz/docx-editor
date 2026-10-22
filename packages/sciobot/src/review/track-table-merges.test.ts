import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import { applyDirectPlan } from '../ai/apply-direct';
import type { NativeEditFailure } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import { planTableMerge, type TableMergeTarget } from '../diff/table-merges';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableMergePlan } from './track-table-merges';
import type { TrackTextContext } from './tracked-inline';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${text.split('|').map(p).join('')}</w:tc>`;
const document_ = (rows = 3) =>
  docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>' +
      `<w:tr>${cell('A', 1200)}${cell('B', 2400)}${cell('C', 3600)}</w:tr>` +
      `<w:tr>${cell('D|D2', 1200)}${cell('E', 2400)}${cell('F', 3600)}</w:tr>` +
      (rows === 3 ? `<w:tr>${cell('G', 1200)}${cell('H', 2400)}${cell('I', 3600)}</w:tr>` : '') +
      '</w:tbl>' +
      p('Tail')
  );
const mount = (bytes = document_(), review = true): DocxEditorInstance =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const cellIds = (part: OoxmlPart) =>
  children(tableOf(part))
    .filter((row) => row.kind === 'tableRow')
    .map((row) =>
      children(row)
        .filter((n) => n.kind === 'tableCell')
        .map((n) => n.id)
    );
function mergePlan(editor: DocxEditorInstance, target: (ids: string[][]) => TableMergeTarget) {
  const part = editor.surface!.session.part();
  const result = planTableMerge(
    {
      revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
      part,
      tableId: tableOf(part).id,
    },
    target(cellIds(part))
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function direct(plan: CanonicalEditPlan, rows = 3): OoxmlPart {
  const editor = mount(document_(rows), false);
  try {
    const result = applyDirectPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    return editor.surface!.session.part();
  } finally {
    editor.destroy();
  }
}
/** Fresh Word paragraph IDs differ; decisions spell the default `continue` explicitly. */
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part)
    .replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '')
    .replace(/<w:vMerge w:val="continue"\/>/g, '<w:vMerge/>');

const targets: [string, (ids: string[][]) => TableMergeTarget, string][] = [
  [
    'a rectangle merge',
    (ids) => ({ kind: 'merge', cellIds: [ids[0]![0]!, ids[0]![1]!, ids[1]![0]!, ids[1]![1]!] }),
    '| A B D D2 E |',
  ],
  ['a vertical merge', (ids) => ({ kind: 'merge', cellIds: [ids[0]![2]!, ids[1]![2]!] }), 'C F'],
  [
    'a split into two rows and two columns',
    (ids) => ({ kind: 'split', cellId: ids[0]![1]!, rows: 2, columns: 2 }),
    '',
  ],
];
for (const [name, target, proposedFragment] of targets)
  for (const action of ['accept', 'reject'] as const)
    test(`${name} ${action}s to the direct output or the exact original with one undo entry`, async () => {
      const editor = mount();
      try {
        const session = editor.surface!.session,
          originalPrint = canonicalOoxmlFingerprint(session.part());
        const [committed] = await cellViews(editor);
        const plan = mergePlan(editor, target);
        const expected = direct(plan);
        const result = trackTableMergePlan(editor, plan, CELL_CONTEXT);
        if (!result.success) throw new Error(`${result.code}: ${result.error}`);
        expect(result.proposals).toHaveLength(1);
        const views = await cellViews(editor);
        expect(views[0]).toBe(committed);
        if (proposedFragment)
          expect(views[1]!.replace(/<br>|\s+/g, ' ')).toContain(proposedFragment);
        const pending = canonicalOoxmlFingerprint(session.part());
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(editor.getReviewItems()).toHaveLength(0);
        if (action === 'accept') expect(comparable(session.part())).toBe(comparable(expected));
        else expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        expect(session.canUndo()).toBe(false);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
      } finally {
        editor.destroy();
      }
    }, 30_000);

test('a merge covering every row accepts to the direct geometry with Word grid compaction', () => {
  const editor = mount(document_(2));
  try {
    const session = editor.surface!.session;
    const plan = mergePlan(editor, targets[0]![1]);
    const expected = direct(plan, 2);
    const result = trackTableMergePlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    // Word drops a grid boundary no row uses after accepting cell deletions; the direct merge
    // keeps it as a span. Content, cell widths, and merges are equal.
    const geometry = (part: OoxmlPart) => {
      const xml = comparable(part);
      return (xml.match(/<w:tc>.*?<\/w:tc>/g) ?? []).map((tc) => [
        /w:w="(\d+)"/.exec(tc)![1],
        /<w:vMerge( w:val="restart")?\/>/.exec(tc)?.[0] ?? '',
        [...tc.matchAll(/<w:t>([^<]*)</g)].map((m) => m[1]).join('|'),
      ]);
    };
    expect(geometry(session.part())).toEqual(geometry(expected));
    expect(comparable(session.part()).match(/<w:gridCol /g)).toHaveLength(2);
  } finally {
    editor.destroy();
  }
});

test('a pending merge survives save/reopen and rejects exactly', async () => {
  const editor = mount();
  try {
    const originalPrint = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const result = trackTableMergePlan(editor, mergePlan(editor, targets[0]![1]), CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      const ids = reopened
        .surface!.session.reviewItems()
        .flatMap((item) => (item.kind === 'revision' ? item.addresses : []))
        .map((entry) => entry.id)
        .sort();
      expect(ids).toEqual(result.proposals[0]!.addresses.map((entry) => entry.id).sort());
      expect(reopened.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(originalPrint);
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('invalid and stacked merge plans refuse atomically', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      revision = session.packageRevision(),
      print = canonicalOoxmlFingerprint(session.part());
    const refuse = (
      plan: CanonicalEditPlan,
      code: NativeEditFailure['code'],
      context: TrackTextContext = CELL_CONTEXT
    ) => {
      const result = trackTableMergePlan(editor, plan, context);
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe(code);
      expect(session.packageRevision()).toBe(revision);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(print);
    };
    const plan = mergePlan(editor, targets[1]![1]);
    refuse(plan, 'invalid-input', { ...CELL_CONTEXT, author: '' });
    refuse({ ...plan, revision: { ...plan.revision, revision: 999 } }, 'stale-revision');
    expect(trackTableMergePlan(editor, plan, CELL_CONTEXT).success).toBe(true);
    const again = trackTableMergePlan(editor, mergePlan(editor, targets[2]![1]), CELL_CONTEXT);
    expect(again.success).toBe(false);
    if (!again.success) expect(again.code).toBe('unsupported-operation');
  } finally {
    editor.destroy();
  }
});

for (const mode of ['editing', 'suggesting'] as const)
  test(`merge and split proposals preserve ${mode} mode and AI attribution`, () => {
    for (const [, target] of targets) {
      const editor = mount();
      try {
        expect(editor.setEditingMode(mode).ok).toBe(true);
        const result = trackTableMergePlan(editor, mergePlan(editor, target), CELL_CONTEXT);
        if (!result.success) throw new Error(result.error);
        expect(
          result.proposals[0]!.addresses.every((entry) => ['ScioBot', ''].includes(entry.author))
        ).toBe(true);
        expect(editor.surface!.editingMode()).toBe(mode === 'editing' ? 'edit' : 'suggest');
      } finally {
        editor.destroy();
      }
    }
  });
