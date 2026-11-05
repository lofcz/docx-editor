import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import { propertyEditor } from '../../test/support/property-execution';
import { listDocument, checklistParagraph } from '../../test/support/list-document';
import { readEditorRevision } from '../projection/revision-cache';
import { planListMembership } from '../diff/lists';
import { applyDirectPlan } from '../ai/apply-direct';
import { trackChecklistPlan } from './track-checklists';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

const context = {
  documentId: 'task-decision-order',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-22T09:00:00Z',
};
for (const firstKind of ['nest', 'check'] as const)
  for (const firstAction of ['accept', 'reject'] as const)
    for (const secondAction of ['accept', 'reject'] as const)
      test(`reverse decisions preserve completion and nesting: ${firstKind} ${firstAction}/${secondAction}`, async () => {
        const read = readOoxmlPackage(listDocument(checklistParagraph(false)));
        if (!read.ok) throw new Error(read.reason);
        const editor = propertyEditor(read.package),
          direct = propertyEditor(read.package);
        const token = (value: typeof editor) =>
          readEditorRevision(value, context.documentId, context.epoch)!;
        const plan = (value: typeof editor, level: number, checked: boolean) => {
          const session = value.surface!.session;
          const result = planListMembership(
            {
              pkg: session.currentPackage(),
              partName: session.part().name,
              revision: token(value),
            },
            [
              {
                paragraphId: session.paragraphIds()[0]!,
                target: { kind: 'checklist', level, checked },
              },
            ]
          );
          if (!result.success) throw new Error(result.error);
          return result.plan;
        };
        try {
          const first = trackChecklistPlan(
            editor,
            plan(editor, firstKind === 'nest' ? 2 : 0, firstKind === 'check'),
            context
          );
          if (!first.success) throw new Error(first.error);
          const second = trackChecklistPlan(editor, plan(editor, 2, true), context);
          if (!second.success) throw new Error(second.error);
          const checked = (firstKind === 'check' ? firstAction : secondAction) === 'accept';
          const level = (firstKind === 'nest' ? firstAction : secondAction) === 'accept' ? 2 : 0;
          if (checked || level) {
            const applied = applyDirectPlan(direct, plan(direct, level, checked), context);
            if (!applied.success) throw new Error(applied.error);
          }
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          const reopened = propertyEditor(saved.package);
          try {
            for (const [proposal, action] of [
              [second, secondAction],
              [first, firstAction],
            ] as const) {
              const decided = (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
                reopened,
                proposal.ownership[0]!.groupId,
                token(reopened),
                context
              );
              if (!decided.success) throw new Error(decided.error);
            }
            expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(direct.surface!.session.part())
            );
            const task = reopened
              .surface!.layout()
              .pages.flatMap((page) => paragraphFragmentsOf(page))
              .flatMap((paragraph) => paragraph.lines.flatMap((line) => line.spans))
              .find((span) => span.text.includes('Task'));
            expect(task!.style.strike).toBe(checked);
            const finalBytes = new Uint8Array(await reopened.save()),
              final = readOoxmlPackage(finalBytes);
            if (!final.ok) throw new Error(final.reason);
            const again = propertyEditor(final.package);
            try {
              expect(canonicalOoxmlFingerprint(again.surface!.session.part())).toBe(
                canonicalOoxmlFingerprint(reopened.surface!.session.part())
              );
            } finally {
              again.destroy();
            }
          } finally {
            reopened.destroy();
          }
        } finally {
          editor.destroy();
          direct.destroy();
        }
      });
