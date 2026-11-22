import { afterEach, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { runToolbarCommand, toolbarCommandState } from '../toolbar-commands.ts';
import { docx } from './paginated-surface-fixtures.ts';
import { openReviewEditor } from './replace-accepted-text-fixture.ts';

const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup();
});
function open(body: string) {
  const { editor, destroy } = openReviewEditor(docx(body));
  cleanups.push(destroy);
  return editor;
}
const p = (s: string) => '<w:p><w:r><w:t>' + s + '</w:t></w:r></w:p>';
test.each([
  ['lowercase', 'a b! straße ος'],
  ['uppercase', 'A B! STRASSE ΟΣ'],
  ['capitalizeWords', 'A B! Straße Ος'],
  ['sentence', 'A b! Straße ος'],
  ['toggle', 'a B! StRASSE ος'],
] as const)('changes selected text using %s', (mode, expected) => {
  const editor = open(p('A b! sTraße ΟΣ'));
  const surface = editor.surface!;
  const id = editor.getParagraphs()[0]!.blockId;
  expect(toolbarCommandState(editor, 'font.changeCase').enabled).toBe(false);
  surface.setSelection({
    anchor: { paragraphId: id, offset: 14 },
    head: { paragraphId: id, offset: 0 },
  });
  const result = runToolbarCommand(editor, 'font.changeCase', mode);
  expect(result).toMatchObject({ ok: true });
  const actual = editor.getParagraphText(id)!.text;
  expect(actual).toBe(expected);
  expect(surface.state().selection).toEqual({
    anchor: { paragraphId: id, offset: expected.length },
    head: { paragraphId: id, offset: 0 },
  });
});
test('preserves mixed runs and hyperlinks across paragraphs in one undo step', async () => {
  const xml =
    '<w:p><w:r><w:t>~abc</w:t></w:r><w:hyperlink w:anchor="spot"><w:r><w:rPr><w:b/></w:rPr><w:t>straße</w:t></w:r></w:hyperlink></w:p>' +
    p('xyz!');
  const editor = open(xml),
    surface = editor.surface!;
  const [a, b] = editor.getParagraphs();
  surface.setSelection({
    anchor: { paragraphId: a!.blockId, offset: 1 },
    head: { paragraphId: b!.blockId, offset: 3 },
  });
  const before = await editor.save();
  expect(editor.exec({ type: 'changeCase', mode: 'uppercase' })).toMatchObject({ ok: true });
  expect(serializeOoxmlPart(surface.session.part())).toMatch(
    /<w:hyperlink[^>]*><w:r><w:rPr><w:b\s*\/><\/w:rPr><w:t[^>]*>STRASSE<\/w:t>/
  );
  expect(editor.getParagraphText(b!.blockId)!.text).toBe('XYZ!');
  expect(editor.exec({ type: 'undo' })).toMatchObject({ ok: true });
  expect(new Uint8Array(await editor.save())).toEqual(new Uint8Array(before));
});
test('suggests the casing and refuses a selection crossing a content lock atomically', () => {
  const editor = open(
    p('abc') +
      '<w:p><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>xyz</w:t></w:r></w:sdtContent></w:sdt></w:p>'
  );
  const [a, b] = editor.getParagraphs();
  const s = editor.surface!;
  const anchor = { paragraphId: a!.blockId, offset: 0 };
  s.setSelection({ anchor, head: { paragraphId: b!.blockId, offset: 3 } });
  const cmd = { type: 'changeCase', mode: 'uppercase' } as const;
  expect(editor.can(cmd)).toMatchObject({ ok: false, code: 'locked' });
  expect(editor.exec(cmd)).toMatchObject({ ok: false, code: 'locked' });
  expect(editor.getParagraphText(a!.blockId)!.text).toBe('abc');
  s.setSelection({ anchor, head: { paragraphId: a!.blockId, offset: 3 } });
  expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' })).toMatchObject({ ok: true });
  expect(editor.exec(cmd)).toMatchObject({ ok: true });
  expect(editor.getParagraphText(a!.blockId)!.acceptedText).toBe('ABC');
  expect(serializeOoxmlPart(s.session.part())).toContain('<w:del ');
  expect(serializeOoxmlPart(s.session.part())).toContain('<w:ins ');
});

test('retains contextual casing across a spellcheck exclusion', () => {
  const editor = open(
    '<w:p><w:r><w:t>Ο</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>Σ</w:t></w:r></w:p>'
  );
  const id = editor.getParagraphs()[0]!.blockId;
  editor.surface!.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 2 },
  });
  expect(editor.exec({ type: 'changeCase', mode: 'lowercase' }).ok).toBe(true);
  expect(editor.getParagraphText(id)!.text).toBe('ος');
});

test('uses authored language, preserves inline objects and refuses split characters', () => {
  const editor = open(
    '<w:p><w:r><w:rPr><w:lang w:val="tr-TR"/></w:rPr><w:t>Iİ</w:t><w:tab/><w:t>Iİ</w:t></w:r></w:p>' +
      p('e\u0301x')
  );
  const [a, b] = editor.getParagraphs();
  const surface = editor.surface!;
  surface.setSelection({
    anchor: { paragraphId: a!.blockId, offset: 0 },
    head: { paragraphId: a!.blockId, offset: 5 },
  });
  expect(editor.exec({ type: 'changeCase', mode: 'lowercase' }).ok).toBe(true);
  expect(editor.getParagraphText(a!.blockId)!.text).toBe('ıi\tıi');
  surface.setSelection({
    anchor: { paragraphId: b!.blockId, offset: 1 },
    head: { paragraphId: b!.blockId, offset: 3 },
  });
  expect(editor.exec({ type: 'changeCase', mode: 'uppercase' }).ok).toBe(false);
  expect(editor.getParagraphText(b!.blockId)!.text).toBe('e\u0301x');
});
