import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createFixedMeasurer, layoutSemanticDocument } from '../index.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const sdt = (alias: string, content: string, lock?: string) =>
  `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/>${lock ? `<w:lock w:val="${lock}"/>` : ''}</w:sdtPr>` +
  `<w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;
function load(xml: string) {
  const read = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'app/xml' });
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}
function layout(part: ReturnType<typeof load>) {
  return layoutSemanticDocument(part, 1, {
    measurer: createFixedMeasurer(6, 14),
    geometry: { width: 300, height: 300, margin: { top: 10, right: 10, bottom: 10, left: 10 } },
  });
}
const document = (rows: string) =>
  `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/>` +
  '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  rows +
  '</w:tbl></w:body></w:document>';

test('row controls retain cell and nested block boundaries, geometry, locks, and saved roster', () => {
  const wrapped = load(
    document(
      sdt(
        'row',
        '<w:tr>' +
          sdt('cell', cell(sdt('nested', p('One'), 'contentLocked'))) +
          cell(p('Two')) +
          '</w:tr>',
        'sdtLocked'
      )
    )
  );
  const bare = load(document('<w:tr>' + cell(p('One')) + cell(p('Two')) + '</w:tr>'));
  const result = layout(wrapped),
    plain = layout(bare);
  const controls = result.contentControls!;
  expect(
    controls.map(({ alias, level, nestingDepth, effectiveLock }) => ({
      alias,
      level,
      nestingDepth,
      effectiveLock,
    }))
  ).toEqual([
    { alias: 'row', level: 'row', nestingDepth: 0, effectiveLock: 'sdtLocked' },
    { alias: 'cell', level: 'cell', nestingDepth: 1, effectiveLock: 'sdtLocked' },
    { alias: 'nested', level: 'block', nestingDepth: 2, effectiveLock: 'sdtContentLocked' },
  ]);
  const wrappedTable = result.pages[0]!.fragments[0]!,
    plainTable = plain.pages[0]!.fragments[0]!;
  expect(wrappedTable.kind).toBe('table');
  expect(wrappedTable.box).toEqual(plainTable.box);
  expect(controls.every((control) => control.fragments.length === 1)).toBe(true);
  expect(controls[0]!.fragments[0]!.box.width).toBeGreaterThan(
    controls[1]!.fragments[0]!.box.width
  );
  expect(controls[1]!.fragments).toEqual(controls[2]!.fragments);
  expect(result.pages[0]!.contentControls!.map((control) => control.id)).toEqual(
    controls.map((control) => control.id)
  );
  expect(layout(load(serializeOoxmlPart(wrapped))).contentControls).toEqual(controls);
});
