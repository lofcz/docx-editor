import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  withPart,
  storyParagraphs,
  paragraphTextOf,
  propertyContainer,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { readNumPr } from '@docx-editor.dev/core/layout';
import { targetDocumentInput, targetParagraph } from '../../test/support/target-document';
import { table } from '../../test/support/target-table-fixture';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { planTargetDocument } from './plan-target';
import { createEditPlan, type CanonicalEdit } from './plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { stageTrackedPlan } from '../review/tracked-stage';
import { trackTableRowPlan } from '../review/track-table-rows';
import { trackTableColumnPlan } from '../review/track-table-columns';

const context = { documentId: 'shared-new-cell-list', epoch: '1' };
const attribution = { ...context, author: 'ScioBot', date: '2026-10-22T08:00:00Z' };
const targetList = {
  kind: 'ordered' as const,
  level: 0,
  instance: { kind: 'new' as const, key: 'shared-across-intents', start: 3 },
};
const html = (anchor: string, kind: 'row' | 'column', text: string) =>
  `<table>\n<tr>\n<td>\n\n${anchor}\n\n</td>\n${kind === 'row' ? '</tr>\n<tr>\n' : ''}<td>\n\n3. ${text}\n\n</td>\n</tr>\n</table>`;

for (const mode of ['direct', 'staged', 'tracked-rows', 'tracked-columns'] as const) {
  test(`${mode} shares one authored-start list instance across separate table intents`, async () => {
    const input = await targetDocumentInput(
      table([['A']]) + targetParagraph('Between') + table([['B']]) + targetParagraph('Tail')
    );
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(input.pkg),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      const session = editor.surface!.session;
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const initial = session.currentPackage();
      const projection = { ...(await projectSourcePackage(initial)), revision };
      const kinds =
        mode === 'tracked-rows'
          ? (['row', 'row'] as const)
          : mode === 'tracked-columns'
            ? (['column', 'column'] as const)
            : (['row', 'column'] as const);
      const target =
        html('A', kinds[0], 'First') +
        '\n\nBetween\n\n' +
        html('B', kinds[1], 'Second') +
        (mode === 'direct' ? '\n\n3. Tail' : '\n\nTail');
      const planned = planTargetDocument({ pkg: initial, revision, projection }, target);
      if (!planned.success) throw new Error(planned.error);
      const edits: CanonicalEdit[] = planned.plan.operations.map((edit) => {
        if (edit.kind === 'insert-row')
          return {
            ...edit,
            intent: {
              ...edit.intent,
              lists: edit.intent.lists!.map((binding) => ({ ...binding, target: targetList })),
            },
          };
        if (edit.kind === 'insert-column')
          return {
            ...edit,
            intent: {
              ...edit.intent,
              lists: edit.intent.lists!.map((binding) => ({ ...binding, target: targetList })),
            },
          };
        if (edit.kind === 'list') return { ...edit, after: targetList };
        throw new Error(`Unexpected ${edit.kind}`);
      });
      expect(edits).toHaveLength(mode === 'direct' ? 3 : 2);
      const made = createEditPlan(revision, edits);
      if (!made.success) throw new Error(made.error);
      let pkg = initial;
      if (mode === 'staged') {
        stageTrackedPlan(
          {
            currentPackage: () => pkg,
            apply: (name, op) => {
              const result = applyTreeOp(pkg.parts.get(name)!, op);
              if (!result.ok) throw new Error(result.reason);
              pkg = withPart(pkg, result.part);
            },
            edit: (change) => {
              pkg = change(pkg);
            },
          },
          made.plan,
          attribution,
          true
        );
        expect(session.currentPackage()).toBe(initial);
      } else {
        const applied =
          mode === 'direct'
            ? applyDirectPlan(editor, made.plan, context)
            : (mode === 'tracked-rows' ? trackTableRowPlan : trackTableColumnPlan)(
                editor,
                made.plan,
                attribution
              );
        if (!applied.success) throw new Error(applied.error);
        pkg = session.currentPackage();
      }
      const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
      if (!reopened.ok) throw new Error(reopened.reason);
      for (const current of [pkg, reopened.package]) {
        const part = current.parts.get(current.mainDocumentPart)!;
        const paragraphs = storyParagraphs(bodyStoryRoot(part)!).filter((node) =>
          ['First', 'Second', ...(mode === 'direct' ? ['Tail'] : [])].includes(
            paragraphTextOf(part, node.id)!
          )
        );
        expect(paragraphs).toHaveLength(mode === 'direct' ? 3 : 2);
        const ids = paragraphs.map(
          (node) => readNumPr([propertyContainer(node, 'paragraphProperties', 'pPr')!])!.numId
        );
        expect(new Set(ids).size).toBe(1);
        const numbering = current.parts.get('/word/numbering.xml')!;
        const instance = numbering.root.children.find(
          (node) =>
            node.kind !== 'textValue' &&
            node.localName === 'num' &&
            node.attributes.some(
              (attribute) => attribute.localName === 'numId' && attribute.value === ids[0]
            )
        );
        if (!instance || instance.kind === 'textValue') throw new Error('Missing instance');
        const overrides = instance.children.filter(
          (node) => node.kind !== 'textValue' && node.localName === 'lvlOverride'
        );
        expect(overrides).toHaveLength(1);
        const override = overrides[0]!;
        if (override.kind === 'textValue') throw new Error('Missing override');
        const start = override.children.find(
          (node) => node.kind !== 'textValue' && node.localName === 'startOverride'
        );
        expect(
          start &&
            start.kind !== 'textValue' &&
            start.attributes.find((attribute) => attribute.localName === 'val')?.value
        ).toBe('3');
      }
    } finally {
      editor.destroy();
    }
  });
}
