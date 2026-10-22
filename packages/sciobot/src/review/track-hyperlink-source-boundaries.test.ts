// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { findNode } from '@docx-editor.dev/core/store';
import {
  mount,
  paragraph,
  plan,
  context,
  fingerprint,
} from '../../test/support/link-structure-execution';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { applyDirectPlan } from '../ai/apply-direct';
for (const count of [1, 2])
  for (const face of ['', '<w:rPr/>'])
    for (const action of ['accept', 'reject'] as const)
      test(`tracked links preserve ${count} original runs, properties ${face || 'absent'}, ${action}`, async () => {
        const run = (text: string) => `<w:r>${face}<w:t>${text}</w:t></w:r>`;
        const body = paragraph(count === 1 ? run('AB') : run('A') + run('B'));
        const editor = mount(body),
          direct = mount(body);
        try {
          const p = findNode(
            editor.surface!.session.part(),
            editor.surface!.session.paragraphIds()[0]!
          )!;
          expect(
            (p.kind === 'textValue' ? [] : p.children).filter((n) => n.kind === 'run')
          ).toHaveLength(count);
          const original = fingerprint(editor),
            target = '[AXB](https://example.org/new)\n\nTail';
          const a = applyDirectPlan(direct, await plan(direct, target), context);
          if (!a.success) throw Error(a.error);
          const result = trackHyperlinkPlan(editor, await plan(editor, target), context);
          if (!result.success) throw Error(result.error);
          expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          expect(fingerprint(editor)).toBe(action === 'accept' ? fingerprint(direct) : original);
        } finally {
          editor.destroy();
          direct.destroy();
        }
      });
