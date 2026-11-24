import { expect, test } from 'bun:test';
import { createFixedMeasurer, layoutSemanticDocument } from '@docx-editor.dev/core/layout';
import { readOoxmlPart } from '@docx-editor.dev/core/store';
import { elevenPointDefaults } from '../../layout/__tests__/fixtures/eleven-point-defaults.ts';
import { plainLineForPaint } from '../plain-line-paint.ts';

test('ordinary words share a painted range before a separately clipped space', () => {
  const read = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      '<w:p><w:r><w:t>abc def hij klm nopq next</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="2400" w:h="8000"/><w:pgMar w:left="0" w:right="0" w:top="0" w:bottom="0"/></w:sectPr>' +
      '</w:body></w:document>',
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw Error(read.reason);
  const layout = layoutSemanticDocument(read.part, 0, {
    measurer: createFixedMeasurer(6, 14),
    styleCascade: elevenPointDefaults(),
  });
  const paragraph = layout.pages[0]!.fragments[0]!;
  if (paragraph.kind !== 'paragraph') throw Error('Expected a paragraph');
  const line = paragraph.lines[0]!;
  const space = line.spans.at(-1)!;
  expect(space.lineEndWhitespace).toBe(true);
  expect(space.text).toBe(' ');
  const painted = plainLineForPaint(line);
  expect(painted.spans).toHaveLength(2);
  expect(painted.spans[0]!.text).toBe('abc def hij klm nopq');
  expect(painted.spans[0]!.range.end).toBe(space.range.start);
  expect(painted.spans[1]).toBe(space);
  expect(line.spans.length).toBeGreaterThan(2);
});
