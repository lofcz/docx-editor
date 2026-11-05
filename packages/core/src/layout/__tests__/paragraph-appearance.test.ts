import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, readOoxmlPart } from '@docx-editor.dev/core/store';
import { resolveEditorModules } from '../../contracts/modules.ts';
import type { ParagraphAppearanceContribution } from '../../contracts/paragraph-appearance.ts';
import { withParagraphAppearance } from '../paragraph-appearance.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import { linesOf } from '../semantic-records.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const marker = (checked: boolean, tag = 'test:task') =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>${checked ? '☒' : '☐'}</w:t></w:r></w:sdtContent></w:sdt>`;
function part(body: string, root = 'document') {
  const parsed = readOoxmlPart(
    `<w:${root} xmlns:w="${W}" xmlns:w14="${W14}">${root === 'document' ? `<w:body>${body}</w:body>` : body}</w:${root}>`,
    { name: `/word/${root}.xml`, contentType: 'app/xml' }
  );
  if (!parsed.ok) throw new Error(parsed.reason);
  return parsed.part;
}
const contribution: ParagraphAppearanceContribution = {
  id: 'task-test',
  resolve: (context) =>
    context.controls.some((control) => control.tag === 'test:task' && control.checked)
      ? { strike: true, color: { val: '606060' } }
      : null,
};
const spans = (layout: ReturnType<typeof layoutSemanticDocument>) =>
  [...linesOf(layout)].flatMap((line) => line.spans);

test('defaults preserve explicit run overrides, authored character styles, and canonical bytes', () => {
  const styles = part(
    '<w:style w:type="character" w:styleId="Authored"><w:rPr><w:strike/><w:color w:val="004488"/></w:rPr></w:style>',
    'styles'
  );
  const document = part(
    `<w:p>${marker(true)}<w:r><w:t>Plain</w:t></w:r><w:r><w:rPr><w:strike w:val="0"/><w:color w:val="CC0000"/></w:rPr><w:t>Direct</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Authored"/></w:rPr><w:t>Character</w:t></w:r></w:p>`
  );
  const before = canonicalOoxmlFingerprint(document.root);
  const layout = layoutSemanticDocument(document, 0, {
    measurer: createFixedMeasurer(),
    styleCascade: withParagraphAppearance(
      buildStyleCascadeTable(styles.root),
      contribution,
      'proposed'
    ),
  });
  const values = spans(layout);
  expect(values.find((span) => span.text === 'Plain')!.style).toMatchObject({
    strike: true,
    color: '606060',
  });
  expect(values.find((span) => span.text === 'Direct')!.style).toMatchObject({
    strike: false,
    color: 'CC0000',
  });
  expect(values.find((span) => span.text === 'Character')!.style).toMatchObject({
    strike: true,
    color: '004488',
  });
  expect(canonicalOoxmlFingerprint(document.root)).toBe(before);
});

for (const mode of ['original', 'proposed', 'all-markup'] as const)
  test(`native revision visibility supplies effective checkbox state: ${mode}`, () => {
    const document = part(
      `<w:p><w:del w:id="1" w:author="Alice">${marker(false)}</w:del><w:ins w:id="2" w:author="Alice">${marker(true)}</w:ins><w:r><w:t>Task</w:t></w:r></w:p>`
    );
    const layout = layoutSemanticDocument(document, 0, {
      measurer: createFixedMeasurer(),
      displayMode: mode,
      styleCascade: withParagraphAppearance(buildStyleCascadeTable(null), contribution, mode),
    });
    expect(spans(layout).find((span) => span.text === 'Task')!.style.strike).toBe(
      mode !== 'original'
    );
  });

test('ordinary Word controls stay untouched, and fact inputs are immutable', () => {
  let calls = 0;
  const inspect: ParagraphAppearanceContribution = {
    id: 'inspect',
    resolve(context) {
      calls++;
      expect(Object.isFrozen(context)).toBe(true);
      expect(Object.isFrozen(context.controls)).toBe(true);
      expect(Object.isFrozen(context.controls[0])).toBe(true);
      return contribution.resolve(context);
    },
  };
  const document = part(`<w:p>${marker(true, 'unrelated')}<w:r><w:t>Task</w:t></w:r></w:p>`);
  const layout = layoutSemanticDocument(document, 0, {
    measurer: createFixedMeasurer(),
    styleCascade: withParagraphAppearance(buildStyleCascadeTable(null), inspect, 'proposed'),
  });
  expect(spans(layout).find((span) => span.text === 'Task')!.style.strike).toBe(false);
  expect(calls).toBe(1);
});

test('reused IDs with changed callbacks invalidate layout caches and conflicting registrations refuse', () => {
  const document = part(`<w:p>${marker(true)}<w:r><w:t>Task</w:t></w:r></w:p>`);
  const base = buildStyleCascadeTable(null),
    cache = createParagraphLayoutCache();
  const changed: ParagraphAppearanceContribution = { id: contribution.id, resolve: () => null };
  const render = (value: ParagraphAppearanceContribution) =>
    layoutSemanticDocument(document, 0, {
      measurer: createFixedMeasurer(),
      cache,
      styleCascade: withParagraphAppearance(base, value, 'proposed'),
    });
  expect(spans(render(contribution)).find((span) => span.text === 'Task')!.style.strike).toBe(true);
  expect(spans(render(changed)).find((span) => span.text === 'Task')!.style.strike).toBe(false);
  expect(() =>
    resolveEditorModules([
      { id: 'one', paragraphAppearance: contribution },
      { id: 'two', paragraphAppearance: changed },
    ])
  ).toThrow('Conflicting paragraph appearance contributions.');
  expect(
    resolveEditorModules([
      { id: 'one', paragraphAppearance: contribution },
      { id: 'two', paragraphAppearance: contribution },
    ]).paragraphAppearance?.resolve
  ).toBe(contribution.resolve);
});

test('invalid appearance output refuses before layout publication', () => {
  const document = part(`<w:p>${marker(true)}<w:r><w:t>Task</w:t></w:r></w:p>`);
  const invalid: ParagraphAppearanceContribution = {
    id: 'invalid',
    resolve: () => ({ color: { val: 'url(https://invalid)' } }),
  };
  expect(() =>
    layoutSemanticDocument(document, 0, {
      measurer: createFixedMeasurer(),
      styleCascade: withParagraphAppearance(buildStyleCascadeTable(null), invalid, 'proposed'),
    })
  ).toThrow('bounded Word color');
});
