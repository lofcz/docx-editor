import { expect, test } from 'bun:test';
import type { DocumentToolRequest, EditResult, ReadResult } from './executor';
import fixture from '../__tests__/fixtures/tool-contracts.json';

// Values captured from the original schema factory. Typecheck verifies the port shape.
const requests = [
  {
    tool: 'read_document',
    input: {
      path: 'lesson',
    },
  },
  {
    tool: 'search_document',
    input: {
      path: 'document',
      pattern: 'Alpha',
      is_regex: false,
      max_results: 20,
    },
  },
  {
    tool: 'edit_document',
    input: {
      path: 'document',
      is_suggesion: false,
      old_string: 'Alpha',
      new_string: 'Beta',
      replace_all: false,
    },
  },
  {
    tool: 'edit_document',
    input: {
      path: 'document',
      is_suggesion: true,
      old_string: 'Alpha',
      new_string: 'Beta',
      replace_all: false,
    },
  },
  {
    tool: 'set_document',
    input: {
      is_suggesion: false,
      path: 'document',
      markdown: '# New',
    },
  },
  {
    tool: 'set_document',
    input: {
      is_suggesion: true,
      path: 'lesson',
      markdown: '# New',
    },
  },
] as const satisfies readonly DocumentToolRequest[];
const results = [
  {
    success: true,
    message:
      'The document has been updated successfully.\n\nUpdated document excerpt around your change (plain text — you may copy old_string for a follow-up edit directly from it):\n<document_excerpt>\n**Beta.**\n\nTail.\n</document_excerpt>',
  },
  {
    success: true,
    message:
      'The document has been updated successfully.\n\nUpdated document excerpt around your change (plain text — you may copy old_string for a follow-up edit directly from it):\n<document_excerpt>\nBeta.\n\n</document_excerpt>',
  },
  {
    success: true,
    message:
      'The document has been updated successfully.\n\nUpdated document excerpt around your change (plain text — you may copy old_string for a follow-up edit directly from it):\n<document_excerpt>\nBeta.\n\n</document_excerpt>',
  },
  {
    success: false,
    error:
      'Found 2 matches for old_string. Set replace_all to true to replace all occurrences, or provide a more unique old_string.',
  },
  {
    success: false,
    error:
      'Expected 1 occurrence(s) of old_string but found 2. Re-read the document to confirm the current count, then adjust old_string (more/less context) or expected_replacements.',
  },
  {
    success: false,
    error:
      'Edit REJECTED — the resulting document is not valid MDX and was NOT applied.\nThe most common cause: old_string and new_string disagree about a tag pair (e.g. old_string swallowed a closing </activity> that new_string does not re-emit, or new_string left an opening tag unclosed). Keep BOTH halves of every tag pair consistently kept or removed.\n\n1. [end-tag-mismatch] line 1:1: Expected a closing tag for `<activity>` (1:1-1:11) before the end of `paragraph`\n\nHow to fix your range:\n- Your new_string opens <activity> but never closes it, and old_string did not leave a matching </activity> in place. Close the tag inside new_string.\n\nContext of the rejected result around the first issue (these lines were NOT applied):\n>>>    1 | <activity>Broken\n       2 | ',
  },
  {
    success: true,
    message:
      'The document has been updated successfully.\n\nWARNING: the document now violates the lesson-plan schema. You MUST fix every issue below with further edits before ending your turn:\n1. [missing-lesson-title]: Document must start with a single <lesson_title grade="…">…</lesson_title> element (not a bare # H1).\n2. [missing-preparation-recap]: Document must contain exactly one <preparation_recap> block immediately after <lesson_title>.\n3. [phase-count]: Document must contain <phase> (expected exactly 3).\n\nUpdated document excerpt around your change (plain text — you may copy old_string for a follow-up edit directly from it):\n<document_excerpt>\nBeta.\n\n</document_excerpt>',
  },
  {
    success: true,
    message: 'The document has been replaced successfully.',
  },
  {
    success: true,
    message: 'The document has been replaced successfully.',
  },
  {
    success: false,
    error:
      'Edit REJECTED — the resulting document is not valid MDX and was NOT applied.\nThe most common cause: old_string and new_string disagree about a tag pair (e.g. old_string swallowed a closing </activity> that new_string does not re-emit, or new_string left an opening tag unclosed). Keep BOTH halves of every tag pair consistently kept or removed.\n\n1. [end-tag-mismatch] line 1:1: Expected a closing tag for `<activity>` (1:1-1:11) before the end of `paragraph`\n\nContext of the rejected result around the first issue (these lines were NOT applied):\n>>>    1 | <activity>Broken',
  },
  {
    success: false,
    error:
      'Edit REJECTED — the document parses correctly but does not satisfy the lesson-plan structure. It was NOT applied. Fix the listed schema requirements.\n\n1. [missing-lesson-title]: Document must start with a single <lesson_title grade="…">…</lesson_title> element (not a bare # H1).\n2. [missing-preparation-recap]: Document must contain exactly one <preparation_recap> block immediately after <lesson_title>.\n3. [phase-count]: Document must contain <phase> (expected exactly 3).',
  },
] as const satisfies readonly EditResult[];
const reads = fixture.readCases.map((item) => item.expected) satisfies ReadResult[];

test('captured successful requests satisfy the private protocol types', () => {
  const expected = fixture.schemaCases
    .filter((item) => item.expected.success)
    .map((item) => ({ tool: item.tool, input: item.expected.data }));
  expect<unknown>(requests).toEqual(expected);
  expect<unknown>(results).toEqual(fixture.executorCases.map((item) => item.expected.result));
  expect(reads).toHaveLength(fixture.readCases.length);
});
