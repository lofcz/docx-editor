import type { ProtectedTargetOptions } from './target-document';
import { targetParagraph as p } from './target-document';
import { sourceControl } from './target-source-fixture';
import { listNumbering, listParagraph } from './list-document';
import { table } from './target-table-fixture';

/** Native Word bodies whose projection is each captured C01 before document. */
const lesson = (): [string, ProtectedTargetOptions] => [
  sourceControl('title', p('Lesson')) +
    sourceControl(
      'phase',
      sourceControl(
        'activity',
        p('First question.') +
          `<w:p>${sourceControl('source', '<w:r><w:t>Source</w:t></w:r>')}</w:p>` +
          p('Last question.')
      )
    ),
  {
    payloads: [
      { version: 1, id: 'title', type: 'lesson_title', attributes: { grade: 7 } },
      { version: 1, id: 'phase', type: 'lesson_phase', attributes: { phaseType: 'evocation' } },
      {
        version: 1,
        id: 'activity',
        type: 'lesson_activity',
        attributes: { name: 'Warmup', duration: 10, privateData: { preserve: true } },
      },
      {
        version: 1,
        id: 'source',
        type: 'source',
        attributes: {
          title: 'Source',
          url: 'https://example.com',
          trust: 'verified',
          privateData: { preserve: true },
        },
      },
    ],
  },
];
export const C01_SOURCES: Record<string, () => [string, ProtectedTargetOptions]> = {
  'formatting and lists': () => [
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>' +
      listParagraph('First', '1') +
      listParagraph('Second', '1') +
      listParagraph('Apple', '2') +
      listParagraph('Pear', '2'),
    { numbering: listNumbering },
  ],
  'tables and links': () => [
    table([
      ['A', 'B'],
      ['One', 'Two'],
    ]) + '<w:p><w:hyperlink r:id="link"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
    {
      relationships:
        '<Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>',
    },
  ],
  'section insertion': () => [p('Opening.') + p('Closing.'), {}],
  'nested MDX and attributes': lesson,
  'source content': lesson,
  deletion: () => [p('Keep.') + p('Remove.') + p('Tail.'), {}],
  'complete rewrite': () => [p('Old document.') + p('Old end.'), {}],
};
