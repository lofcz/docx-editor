import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { wordPart, paraIds, bornWordBody } from '../../test/support/born-word-body';
import { docxFixture } from '../../test/support/document';
import {
  propertyContext,
  propertyEditor,
  propertyToken,
} from '../../test/support/property-execution';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { planTargetDocument } from '../diff/plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';

const table = '| A | **B** |\n| :--- | ---: |\n| [C](https://example.com/) | `D` |';
for (const placement of ['middle', 'end', 'two-tables', 'lists', 'quotes'] as const)
  for (const action of ['accept', 'reject'] as const)
    test(`new ${placement} native tables ${action} with exact Word structure, history, and later waves`, async () => {
      const loaded = readOoxmlPackage(
        docxFixture(
          '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>',
          placement === 'quotes'
            ? {
                styles:
                  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style></w:styles>',
              }
            : {}
        )
      );
      if (!loaded.ok) throw Error(loaded.reason);
      const editor = propertyEditor(loaded.package),
        direct = propertyEditor(loaded.package);
      const view = async (value: typeof editor, original = false) => {
        const result = await (original ? readCommittedView : readWorkingView)(
          value.surface!.session,
          propertyToken(value)
        );
        if (!result.success) throw Error(result.error);
        return result.view.projection;
      };
      try {
        const before = canonicalOoxmlFingerprint(
          wordPart(editor.surface!.session.currentPackage())
        );
        const liveBefore = canonicalOoxmlFingerprint(editor.surface!.session.part());
        const initial = (await view(editor)).text;
        const imported = new Set(paraIds(wordPart(editor.surface!.session.currentPackage()).root));
        const body =
          placement === 'two-tables'
            ? `${table}\n\n${table.replace('**B**', 'Second')}`
            : placement === 'lists'
              ? `- One\n\n${table}\n\n- Two`
              : placement === 'quotes'
                ? `> **Quoted** one\n>\n> Quoted two\n\n${table}\n\n> Last quote`
                : table;
        const target =
          placement === 'end' ? `Before\n\nAfter\n\n${body}` : `Before\n\n${body}\n\nAfter`;
        const first = await recomposeOwnedProposals(
          editor,
          {
            revision: propertyToken(editor),
            targetMarkdown: target,
            ownership: [],
            contentProofs: [],
          },
          propertyContext
        );
        if (!first.success) throw Error(first.error);
        expect((await view(editor, true)).text).toBe(initial);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(liveBefore);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        const next = target.replace('| A |', '| Revised |');
        const second = await recomposeOwnedProposals(
          editor,
          {
            revision: propertyToken(editor),
            targetMarkdown: next,
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          propertyContext
        );
        if (!second.success) throw Error(second.error);
        const projected = await view(direct);
        const plan = planTargetDocument(
          {
            pkg: direct.surface!.session.currentPackage(),
            revision: propertyToken(direct),
            projection: projected,
          },
          next
        );
        if (!plan.success) throw Error(plan.error);
        const applied = applyDirectPlan(direct, plan.plan, propertyContext);
        if (!applied.success) throw Error(applied.error);
        expect((await view(editor)).text).toBe((await view(direct)).text);
        const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
        const result = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          editor,
          propertyToken(editor),
          propertyContext
        );
        if (!result.success) throw Error(result.error);
        if (action === 'accept')
          expect(bornWordBody(editor.surface!.session.currentPackage(), imported)).toBe(
            bornWordBody(direct.surface!.session.currentPackage(), imported)
          );
        else
          expect(
            canonicalOoxmlFingerprint(wordPart(editor.surface!.session.currentPackage()))
          ).toBe(before);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
      } finally {
        editor.destroy();
        direct.destroy();
      }
    }, 30000);
