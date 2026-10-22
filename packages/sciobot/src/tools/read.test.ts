import { expect, test } from 'bun:test';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import reference from '../__tests__/fixtures/tool-contracts.json';
import { docxFixture } from '../../test/support/document';
import { projectMarkdown } from '../projection/markdown';
import { findAndReplace } from './find-replace';
import { readDocument, READ_DOCUMENT_DEFAULT_LIMIT } from './read';

for (const captured of reference.readCases) {
  test(`C02 read: ${captured.name}`, () => {
    expect(readDocument({ text: captured.projection }, captured.input)).toEqual(captured.expected);
  });
}

test('invalid and fractional paging inputs preserve the effective client defaults', () => {
  const projection = { text: 'One\nTwo\nThree' };
  for (const value of [NaN, Infinity, -1, 0]) {
    expect(readDocument(projection, { offset: value, limit: value })).toEqual({
      content: projection.text,
      startLine: 1,
      endLine: 3,
      totalLines: 3,
      truncated: false,
    });
  }
  expect(readDocument(projection, { offset: 2.9, limit: 1.9 })).toEqual({
    content: 'Two',
    startLine: 2,
    endLine: 2,
    totalLines: 3,
    truncated: true,
  });
  expect(READ_DOCUMENT_DEFAULT_LIMIT).toBe(reference.defaults.clientReadLimit);
});

test('empty documents and final blank lines retain the source line count', () => {
  expect(readDocument({ text: '' })).toEqual({
    content: '',
    startLine: 1,
    endLine: 1,
    totalLines: 1,
    truncated: false,
  });
  expect(readDocument({ text: 'A\n' }, { offset: 2 })).toEqual({
    content: '',
    startLine: 2,
    endLine: 2,
    totalLines: 2,
    truncated: false,
  });
});

test('a native projection page stays copyable and leaves the full projection intact', async () => {
  const opened = openDocumentForExport(
    docxFixture(
      '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Žák 😀</w:t></w:r></w:p>'
    )
  );
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const projection = projectMarkdown(await opened.session.layout(), {
      documentId: 'doc',
      epoch: '1',
      revision: 0,
    });
    const before = JSON.stringify(projection);
    const result = readDocument(projection, { offset: 3, limit: 1 });
    expect(result).toEqual({
      content: '**Žák 😀**',
      startLine: 3,
      endLine: 3,
      totalLines: 3,
      truncated: false,
    });
    expect(findAndReplace(projection.text, result.content, 'Updated', false)).toEqual({
      success: true,
      newContent: '# Title\n\nUpdated',
    });
    expect(JSON.stringify(projection)).toBe(before);
    result.content = 'Caller-owned result';
    expect(readDocument(projection, { offset: 3, limit: 1 }).content).toBe('**Žák 😀**');
  } finally {
    opened.session.dispose();
  }
});
