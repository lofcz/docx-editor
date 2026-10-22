import { expect, test } from 'bun:test';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import reference from '../__tests__/fixtures/tool-contracts.json';
import { docxFixture } from '../../test/support/document';
import { projectMarkdown } from '../projection/markdown';
import { searchDocument } from './search';

for (const captured of reference.searchCases) {
  test(`C02 search: ${captured.name}`, () => {
    const projection = Object.freeze({ text: captured.projection });
    expect<unknown>(searchDocument(projection, captured.input)).toEqual(captured.expected);
    expect(projection.text).toBe(captured.projection);
  });
}

test('search counts matching lines and resets global regex state between lines', () => {
  expect(
    searchDocument(
      { text: 'Alpha Alpha\nAlpha\nBeta\nALPHA' },
      { pattern: '^alpha', isRegex: true, maxResults: 2 }
    )
  ).toEqual({
    matches: [
      { lineNumber: 1, content: 'Alpha Alpha', contextBefore: null, contextAfter: 'Alpha' },
      { lineNumber: 2, content: 'Alpha', contextBefore: 'Alpha Alpha', contextAfter: 'Beta' },
    ],
    totalMatches: 3,
    truncated: true,
  });
});

test('literal regex punctuation and cross-line patterns retain production semantics', () => {
  expect(
    searchDocument(
      { text: '.*(x)[y]\nAnything' },
      { pattern: '.*(x)[y]', isRegex: false, maxResults: 20 }
    )
  ).toEqual({
    matches: [
      { lineNumber: 1, content: '.*(x)[y]', contextBefore: null, contextAfter: 'Anything' },
    ],
    totalMatches: 1,
    truncated: false,
  });
  expect(
    searchDocument(
      { text: 'Alpha\nBeta' },
      { pattern: 'Alpha\\nBeta', isRegex: true, maxResults: 20 }
    )
  ).toEqual({ matches: [], totalMatches: 0, truncated: false });
});

test('zero-width regex matches terminate and retain empty-line context', () => {
  expect(
    searchDocument({ text: 'A\n\nB\n' }, { pattern: '^$', isRegex: true, maxResults: 20 })
  ).toEqual({
    matches: [
      { lineNumber: 2, content: '', contextBefore: 'A', contextAfter: 'B' },
      { lineNumber: 4, content: '', contextBefore: 'B', contextAfter: null },
    ],
    totalMatches: 2,
    truncated: false,
  });
  expect(searchDocument({ text: '' }, { pattern: '^$', isRegex: true, maxResults: 20 })).toEqual({
    matches: [{ lineNumber: 1, content: '', contextBefore: null, contextAfter: null }],
    totalMatches: 1,
    truncated: false,
  });
});

test('native Markdown search returns copyable lines and leaves its projection intact', async () => {
  const opened = openDocumentForExport(
    docxFixture('<w:p><w:r><w:t>Žák 😀</w:t></w:r></w:p><w:p><w:r><w:t>Další žák</w:t></w:r></w:p>')
  );
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const projection = projectMarkdown(await opened.session.layout(), {
      documentId: 'doc',
      epoch: '1',
      revision: 0,
    });
    const before = JSON.stringify(projection);
    expect(searchDocument(projection, { pattern: 'žák', isRegex: false, maxResults: 20 })).toEqual({
      matches: [
        { lineNumber: 1, content: 'Žák 😀', contextBefore: null, contextAfter: '' },
        { lineNumber: 3, content: 'Další žák', contextBefore: '', contextAfter: null },
      ],
      totalMatches: 2,
      truncated: false,
    });
    expect(
      searchDocument(projection, { pattern: '[', isRegex: true, maxResults: 20 })
    ).toMatchObject({ totalMatches: 0, matches: [] });
    expect(JSON.stringify(projection)).toBe(before);
  } finally {
    opened.session.dispose();
  }
});
