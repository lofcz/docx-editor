import { nodeChildren } from '../../test/support/native-table-oracle';
import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { readTableStructure } from '@docx-editor.dev/core/layout';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import { groupToken, groups } from '../../test/support/review-groups';
import { cases, document_, target } from '../../test/support/compound-target-table';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { decodeEditPlan, serializeEditPlan } from '../diff/plan';
import { targetAstKey } from '../diff/target-key';
import { parseTargetMarkdown } from '../projection/parse-target';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
type Editor = ReturnType<typeof propertyEditor>;
const context = { ...propertyContext, date: '2026-10-21T21:00:00.000Z' };
const mount = (merged = false) => {
  const opened = readOoxmlPackage(document_(merged));
  if (!opened.ok) throw new Error(opened.reason);
  return propertyEditor(opened.package);
};
async function view(editor: Editor, committed = false) {
  const result = await (committed ? readCommittedView : readWorkingView)(
    editor.surface!.session,
    groupToken(editor),
    { tableSyntax: 'auto' }
  );
  if (!result.success) throw new Error(result.error);
  return result.view.projection;
}
function ast(text: string) {
  const result = parseTargetMarkdown(text);
  if (!result.success) throw new Error(result.error);
  return targetAstKey(result.target.normalizedAst);
}
const fingerprint = (editor: Editor) => canonicalOoxmlFingerprint(editor.surface!.session.part());
const source = async (editor: Editor) => ({
  revision: groupToken(editor),
  pkg: editor.surface!.session.currentPackage(),
  projection: await view(editor),
});
for (const fixture of cases)
  for (const decision of ['accept', 'reject'] as const)
    test(`complete compound target ${fixture.name}: ${decision}`, async () => {
      const direct = mount(fixture.source === 'merged'),
        editor = mount(fixture.source === 'merged');
      try {
        const initial = ast((await view(editor)).text),
          before = fingerprint(editor),
          markdown = target(fixture.rows);
        const planned = planTargetDocument(await source(direct), markdown);
        if (!planned.success) throw new Error(`direct plan: ${planned.error}`);
        expect(decodeEditPlan(serializeEditPlan(planned.plan)).success).toBe(true);
        expect(planned.plan.operations.some((edit) => edit.kind === 'table-geometry')).toBe(true);
        const applied = applyDirectPlan(direct, planned.plan, context);
        if (!applied.success) throw new Error(`direct apply: ${applied.error}`);
        const expected = ast((await view(direct)).text),
          replan = planTargetDocument(await source(direct), markdown);
        if (!replan.success) throw new Error(replan.error);
        expect(replan.plan.operations).toEqual([]);
        const body = bodyStoryRoot(direct.surface!.session.part())!;
        const table =
          body.kind !== 'textValue'
            ? body.children.find((node) => node.kind === 'table')
            : undefined;
        expect(table).toBeDefined();
        const shape = readTableStructure(table!, 500, 0, undefined, 'proposed')!;
        expect(shape.rows.length).toBe(fixture.rows.length);
        const occupied = new Set<string>();
        const expectedHeads: { row: number; column: number; span: number }[] = [];
        for (const [r, row] of fixture.rows.entries()) {
          let column = 0;
          for (const cell of row) {
            while (occupied.has(`${r}/${column}`)) column++;
            expectedHeads.push({ row: r, column, span: cell.columns ?? 1 });
            for (let rr = r; rr < r + (cell.rows ?? 1); rr++)
              for (let c = column; c < column + (cell.columns ?? 1); c++)
                occupied.add(`${rr}/${c}`);
            column += cell.columns ?? 1;
          }
        }
        const actualHeads = shape.rows.flatMap((row, r) =>
          row.cells
            .filter((cell) => !cell.vMergeContinue)
            .map((cell) => ({
              row: r,
              column: cell.resolvedGridColumn ?? cell.gridColumn,
              span: cell.resolvedGridSpan ?? cell.gridSpan,
            }))
        );
        expect(actualHeads).toEqual(expectedHeads);
        const directAfter = fingerprint(direct);
        expect(direct.exec({ type: 'undo' }).ok).toBe(true);
        expect(ast((await view(direct)).text)).toBe(initial);
        expect(direct.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(direct)).toBe(directAfter);
        let publications = 0;
        const stop = editor.surface!.session.subscribe(() => publications++);
        const wave = await recomposeOwnedProposals(
          editor,
          {
            revision: groupToken(editor),
            targetMarkdown: markdown,
            ownership: [],
            contentProofs: [],
          },
          context
        );
        stop();
        if (!wave.success) throw new Error(`tracked apply: ${wave.error}`);
        expect(publications).toBe(1);
        expect(ast((await view(editor)).text)).toBe(expected);
        expect(ast((await view(editor, true)).text)).toBe(initial);
        expect(groups(editor)).toHaveLength(1);
        const pending = fingerprint(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(pending);
        const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!reopened.ok) throw new Error(reopened.reason);
        const again = propertyEditor(reopened.package);
        try {
          expect(ast((await view(again)).text)).toBe(expected);
          expect(ast((await view(again, true)).text)).toBe(initial);
          for (const group of groups(again)) {
            const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              again,
              group.id,
              groupToken(again),
              context
            );
            if (!result.success) throw new Error(result.error);
          }
          expect(ast((await view(again)).text)).toBe(decision === 'accept' ? expected : initial);
          expect(groups(again)).toEqual([]);
        } finally {
          again.destroy();
        }
        for (const group of groups(editor)) {
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            group.id,
            groupToken(editor),
            context
          );
          if (!result.success) throw new Error(result.error);
        }
        expect(ast((await view(editor)).text)).toBe(decision === 'accept' ? expected : initial);
        expect(groups(editor)).toEqual([]);
      } finally {
        direct.destroy();
        editor.destroy();
      }
    }, 30000);

for (const change of ['unchanged', 'body'] as const)
  test(`default GFM ${change} after a compound target preserves native cell paragraphs`, async () => {
    const editor = mount();
    try {
      const fixture = cases.find(
          (entry) => entry.name === 'new column 3 merges allocated and retained cells'
        )!,
        markdown = target(fixture.rows);
      const plan = planTargetDocument(await source(editor), markdown);
      if (!plan.success) throw new Error(plan.error);
      const applied = applyDirectPlan(editor, plan.plan, context);
      if (!applied.success) throw new Error(applied.error);
      const original = (await view(editor)).text,
        table = nodeChildren(bodyStoryRoot(editor.surface!.session.part())!).find(
          (node) => node.kind === 'table'
        )!,
        saved = fingerprint(editor);
      const raw = await readWorkingView(editor.surface!.session, groupToken(editor));
      if (!raw.success) throw new Error(raw.error);
      const wave = await recomposeOwnedProposals(
        editor,
        {
          revision: groupToken(editor),
          targetMarkdown:
            change === 'body'
              ? raw.view.projection.text.replace('Lead', 'Latest Lead')
              : raw.view.projection.text,
          ownership: [],
          contentProofs: [],
        },
        context
      );
      if (!wave.success) throw new Error(wave.error);
      expect(wave.committed).toBe(change === 'body');
      expect(ast((await view(editor)).text)).toBe(
        ast(change === 'body' ? original.replace('Lead', 'Latest Lead') : original)
      );
      const retained = nodeChildren(bodyStoryRoot(editor.surface!.session.part())!).find(
        (node) => node.kind === 'table'
      );
      expect(retained).toBe(table);
      if (change === 'unchanged') expect(fingerprint(editor)).toBe(saved);
    } finally {
      editor.destroy();
    }
  }, 30000);
