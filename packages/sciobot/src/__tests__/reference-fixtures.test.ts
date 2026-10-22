import { expect, test } from 'bun:test';
import reference from './fixtures/ai-pairs.json';
import matcher from './fixtures/find-replace.json';
import tools from './fixtures/tool-contracts.json';

test('the Plate oracle contains all seven named cases and content hashes', () => {
  expect(reference.schemaVersion).toBe(1);
  expect(reference.cases.map((item) => item.name)).toEqual([
    'formatting and lists',
    'tables and links',
    'section insertion',
    'nested MDX and attributes',
    'source content',
    'deletion',
    'complete rewrite',
  ]);
  for (const checkout of Object.values(reference.sources)) {
    expect(checkout.revision).toMatch(/^[a-f0-9]{40}$/);
    expect(Object.keys(checkout.files).length).toBeGreaterThan(0);
    for (const hash of Object.values(checkout.files)) expect(hash).toMatch(/^[a-f0-9]{64}$/);
  }
  for (const item of reference.cases) {
    expect(item.beforeMarkdown).not.toBe(item.afterMarkdown);
    for (const tree of [item.original, item.canonicalBefore, item.expected]) {
      expect(Array.isArray(tree)).toBe(true);
      expect(tree.length).toBeGreaterThan(0);
      expect(JSON.stringify(tree)).not.toMatch(/"id":/);
    }
  }
});

test('the oracle includes nested property edits and structural table changes', () => {
  const nested = reference.cases.find((item) => item.name === 'nested MDX and attributes')!;
  expect(JSON.stringify(nested.original)).toContain('"duration":10');
  expect(JSON.stringify(nested.expected)).toContain('"duration":15');
  expect(JSON.stringify(nested.expected)).toContain('A better question.');
  const table = reference.cases.find((item) => item.name === 'tables and links')!;
  expect(JSON.stringify(table.expected)).toContain('Three');
  expect(JSON.stringify(table.expected)).toContain('https://example.org');
});

test('matcher records cover every original regression with explicit results', () => {
  expect(matcher.schemaVersion).toBe(1);
  expect(matcher.cases).toHaveLength(20);
  expect(new Set(matcher.cases.map((item) => item.sourceTest))).toEqual(
    new Set(matcher.source.testNames)
  );
  expect(matcher.source.testNames).toHaveLength(17);
  for (const item of matcher.cases) {
    expect(item.input.content).toBeTypeOf('string');
    expect(item.input.oldString).toBeTypeOf('string');
    expect(item.input.newString).toBeTypeOf('string');
    if (item.expected.success) {
      expect(item.expected.newContent).toBeTypeOf('string');
    } else {
      expect(item.expected.error).toBeTypeOf('string');
    }
  }
});

test('tool contracts cover successful and failing requests for all four tools', () => {
  for (const name of ['read_document', 'search_document', 'edit_document', 'set_document']) {
    const examples = tools.schemaCases.filter((item) => item.tool === name);
    expect(examples.some((item) => item.expected.success)).toBe(true);
    expect(examples.some((item) => !item.expected.success)).toBe(true);
  }
  expect(tools.defaults).toEqual({
    declaredReadLimit: 500,
    clientReadLimit: 1000,
    maxSearchResults: 30,
  });
  const editDefault = tools.schemaCases.find(
    (item) => item.tool === 'edit_document' && item.input.is_suggesion === undefined
  )!;
  expect(editDefault.expected.data?.is_suggesion).toBe(false);
  expect(tools.schemaCases.find((item) => item.input.path === 'missing')?.expected.success).toBe(
    false
  );
  expect(
    tools.readCases.find((item) => item.name === 'effective client default')?.expected.endLine
  ).toBe(1000);
  expect(tools.searchCases.find((item) => item.name === 'invalid regex')?.expected.error).toContain(
    'Invalid regex'
  );
});

test('executor captures retain direct defaults, review views, and atomic errors', () => {
  const direct = tools.executorCases.find((item) => item.name === 'direct default')!;
  expect(direct.expected.effects).toEqual(['direct']);
  expect(direct.expected.workingMarkdown).toBe(direct.expected.committedMarkdown);
  expect(direct.expected.result.message).toContain('Beta.');
  const suggestion = tools.executorCases.find((item) => item.name === 'explicit suggestion')!;
  expect(suggestion.expected.effects).toEqual(['suggestion']);
  expect(suggestion.expected.workingMarkdown).toContain('Beta.');
  expect(suggestion.expected.committedMarkdown).toContain('Alpha.');
  for (const item of tools.executorCases.filter((item) => !item.expected.result.success)) {
    expect(item.expected.effects).toEqual([]);
    expect(item.expected.workingMarkdown.trim()).toBe(item.beforeMarkdown.trim());
  }
});
