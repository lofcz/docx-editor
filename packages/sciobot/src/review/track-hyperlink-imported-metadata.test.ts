import { expect, test } from 'bun:test';
import {
  mount,
  paragraph,
  plan,
  context,
  fingerprint,
} from '../../test/support/link-structure-execution';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { applyDirectPlan } from '../ai/apply-direct';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { propertyEditor } from '../../test/support/property-execution';
async function reopen(editor: ReturnType<typeof mount>) {
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw Error(read.reason);
  const next = propertyEditor(read.package);
  editor.destroy();
  return next;
}
for (const saved of [false, true])
  for (const kind of ['run', 'text', 'both'])
    for (const action of ['accept', 'reject'] as const)
      test(`imported ${kind} shell fidelity ${action}${saved ? ' after reopen' : ''}`, async () => {
        const body = paragraph(
          `<w:r xmlns:keep="urn:keep"${kind !== 'text' ? ' keep:run="original"' : ''}><w:rPr/><w:t${kind !== 'run' ? ' keep:text="original"' : ''}>Alpha end</w:t></w:r>`
        );
        let editor = mount(body);
        const direct = mount(body);
        try {
          const original = fingerprint(editor),
            target = 'Alpha [new](https://example.org/new) end\n\nTail';
          const a = applyDirectPlan(direct, await plan(direct, target), context);
          if (!a.success) throw Error(a.error);
          const b = trackHyperlinkPlan(editor, await plan(editor, target), context);
          if (!b.success) throw Error(b.error);
          if (saved) editor = await reopen(editor);
          expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          expect(fingerprint(editor)).toBe(action === 'accept' ? fingerprint(direct) : original);
        } finally {
          editor.destroy();
          direct.destroy();
        }
      });

for (const variant of ['distinct', 'same', 'mixed-space'] as const)
  for (const saved of [false, true])
    for (const action of ['accept', 'reject'] as const)
      test(`${variant} imported text shells survive ${action}${saved ? ' after reopen' : ''}`, async () => {
        const body = paragraph(
          '<w:r xmlns:keep="urn:keep" keep:run="original"><w:rPr/>' +
            `<w:t keep:text="first"${variant === 'mixed-space' ? ' xml:space="preserve"' : ''}>Alpha </w:t><w:t keep:text="${variant === 'same' ? 'first' : 'second'}">end</w:t></w:r>`
        );
        let editor = mount(body);
        const direct = mount(body);
        try {
          const original = fingerprint(editor);
          const target = 'Alpha [new](https://example.org/new) end\n\nTail';
          const a = applyDirectPlan(direct, await plan(direct, target), context);
          if (!a.success) throw Error(a.error);
          const b = trackHyperlinkPlan(editor, await plan(editor, target), context);
          if (!b.success) throw Error(b.error);
          if (saved) editor = await reopen(editor);
          expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          expect(fingerprint(editor)).toBe(action === 'accept' ? fingerprint(direct) : original);
        } finally {
          editor.destroy();
          direct.destroy();
        }
      });
