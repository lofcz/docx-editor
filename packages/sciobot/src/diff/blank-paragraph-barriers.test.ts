import { expect, test } from 'bun:test';
import { paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import { readEditorRevision } from '../projection/revision-cache';
import { DirectTargetBaseline } from '../ai/direct-baseline';
import { applyDirectPlan } from '../ai/apply-direct';
import { recomposeOwnedProposals } from '../review/recompose';
import { readCommittedView } from '../review/committed-view';
import { readWorkingView } from '../review/working-view';
import { planTargetDocument } from './plan-target';

// Word's blank lines are empty paragraphs. Markdown cannot address them, so edits plan around
// them and leave each one in place.
const p = (text: string) => (text ? `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` : '<w:p/>');
const CASES = [
  { name: 'trailing', body: ['One.', ''], target: 'One changed.', after: ['One changed.', ''] },
  { name: 'leading', body: ['', 'One.'], target: 'One changed.', after: ['', 'One changed.'] },
  {
    name: 'between',
    body: ['One.', '', 'Two.'],
    target: 'One changed.\n\nTwo.\n\nThree.',
    after: ['One changed.', '', 'Two.', 'Three.'],
  },
  {
    name: 'several',
    body: ['', 'One.', '', '', 'Two.', ''],
    target: 'One.\n\nNew.\n\nTwo changed.',
    // A new block between two barrier windows joins the later window when nothing anchors it.
    after: ['', 'One.', '', '', 'New.', 'Two changed.', ''],
  },
] as const;

function editorFor(body: readonly string[]) {
  const prepared = prepareNativeDocument(docxFixture(body.map(p).join('')));
  if (!prepared.success) throw Error(prepared.error);
  const read = readOoxmlPackage(prepared.document);
  if (!read.ok) throw Error(read.reason);
  return propertyEditor(read.package);
}
const texts = (editor: ReturnType<typeof editorFor>) => {
  const session = editor.surface!.session;
  return [...session.paragraphIds()].map((id) => paragraphTextOf(session.part(), id) ?? '');
};
const token = (editor: ReturnType<typeof editorFor>) =>
  readEditorRevision(editor, propertyContext.documentId, propertyContext.epoch)!;

for (const { name, body, target, after } of CASES) {
  test(`a direct edit plans around ${name} blank paragraphs`, async () => {
    const editor = editorFor(body);
    const baseline = new DirectTargetBaseline();
    const before = await baseline.prepare(editor, propertyContext);
    if (!before.success) throw Error(before.error);
    const planned = planTargetDocument(before.input, target);
    if (!planned.success) throw Error(planned.error);
    const applied = applyDirectPlan(editor, planned.plan, propertyContext);
    if (!applied.success) throw Error(applied.error);
    expect(texts(editor)).toEqual([...after]);
    const landed = await baseline.prepare(editor, propertyContext);
    if (!landed.success) throw Error(landed.error);
    const again = planTargetDocument(landed.input, target);
    if (!again.success) throw Error(again.error);
    expect(again.plan.operations).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(editor)).toEqual([...body]);
  });

  test(`a tracked wave plans around ${name} blank paragraphs`, async () => {
    const editor = editorFor(body);
    const result = await recomposeOwnedProposals(
      editor,
      { revision: token(editor), targetMarkdown: target, ownership: [] },
      { ...propertyContext, author: 'AI', date: '2026-10-31T00:00:00.000Z' }
    );
    if (!result.success) throw Error(result.error);
    const [committed, working] = await Promise.all([
      readCommittedView(editor.surface!.session, token(editor)),
      readWorkingView(editor.surface!.session, token(editor)),
    ]);
    if (!committed.success || !working.success) throw Error('unreadable views');
    // Blank paragraphs project as extra blank lines; the content is the same.
    const content = (markdown: string) => markdown.replace(/\n{3,}/g, '\n\n').trim();
    expect(content(committed.view.projection.text)).toBe(
      body
        .filter(Boolean)
        .map((text) => text.replace('.', '\\.'))
        .join('\n\n')
    );
    expect(content(working.view.projection.text)).toBe(target.replaceAll('.', '\\.'));
    expect(editor.getReviewItems().length).toBeGreaterThan(0);
  });
}
