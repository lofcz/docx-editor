import { expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import fixture from './fixtures/native-commands.json';

const root = resolve(import.meta.dirname, '../../../..');
const groups = [
  'BasicBlocks',
  'BasicMarks',
  'Font',
  'List',
  'Align',
  'LineHeight',
  'Link',
  'Table',
];
const leaves = new Set(
  Array.from({ length: 13 }, (_, index) => `N${String(index + 2).padStart(2, '0')}`)
);
leaves.add('N22');
const cases = new Map(fixture.cases.map((entry) => [entry.id, entry]));

function nodes(value: unknown): void {
  expect(Array.isArray(value)).toBe(true);
  for (const node of value as Record<string, unknown>[]) {
    expect(node).toBeObject();
    if ('text' in node) expect(typeof node.text).toBe('string');
    else {
      expect(typeof node.type).toBe('string');
      expect((node.children as unknown[]).length).toBeGreaterThan(0);
      nodes(node.children);
    }
  }
}
function selection(value: unknown): void {
  if (value === null) return;
  expect(value).toBeObject();
  for (const point of Object.values(value as Record<string, { path: number[]; offset: number }>)) {
    expect(Array.isArray(point.path)).toBe(true);
    expect(point.path.length).toBeGreaterThan(0);
    for (const index of point.path) expect(Number.isInteger(index) && index >= 0).toBe(true);
    expect(Number.isInteger(point.offset) && point.offset >= 0).toBe(true);
  }
}

test('native command fixture has captured document states and explicit implementation mappings', () => {
  expect(fixture.schemaVersion).toBe(1);
  expect(fixture.groups).toEqual(groups);
  expect(cases.size).toBe(fixture.cases.length);
  for (const entry of fixture.cases) {
    expect(groups).toContain(entry.group);
    expect(entry.operation).toBeString();
    if (!('command' in entry) || !('arguments' in entry)) throw new Error('Missing command inputs');
    expect(entry.command).toBe(entry.id);
    expect(entry.arguments).toBeObject();
    if ('leaf' in entry.mapping && typeof entry.mapping.leaf === 'string') {
      expect(entry.mapping.kind).toBe('implementation-leaf');
      expect(leaves.has(entry.mapping.leaf)).toBe(true);
    } else if ('command' in entry.mapping) {
      expect(entry.mapping.kind).toBe('native-command');
      expect(entry.mapping.command).toBe('deleteTable');
    } else throw new Error(`Missing native mapping: ${entry.id}`);
    nodes(entry.before.value);
    nodes(entry.after.value);
    selection(entry.before.selection);
    selection(entry.after.selection);
    expect(entry.before.value.length).toBeGreaterThan(0);
    expect(entry.after.value.length).toBeGreaterThan(0);
    expect(entry.before.selection).toHaveProperty('anchor');
    expect(entry.before.selection).toHaveProperty('focus');
  }
  expect(fixture.source.scope).toContain('Actual installed Plate transforms');
  expect(fixture.source.revision).toMatch(/^[0-9a-f]{40}$/);
  for (const hash of Object.values(fixture.source.files)) expect(hash).toMatch(/^[0-9a-f]{64}$/);
});

test('every required command family retains explicit action cases', () => {
  const required = [
    ...['p', 'h1', 'h2', 'h3', 'blockquote'].flatMap((key) => [
      `block.${key}`,
      `block.${key}.from-list`,
      `block.insert.${key}`,
    ]),
    'block.hr',
    'block.insert.hr',
    ...['bold', 'italic', 'underline', 'code', 'strikethrough', 'subscript', 'superscript'].flatMap(
      (key) => ['on', 'off', 'caret'].map((state) => `mark.${key}.${state}`)
    ),
    ...['fontFamily', 'fontSize', 'color', 'backgroundColor'].flatMap((key) => [
      `font.${key}.set`,
      `font.${key}.clear`,
    ]),
    'font.color.inherit',
    'font.background.transparent',
    ...[
      'disc',
      'circle',
      'square',
      'decimal',
      'lower-alpha',
      'upper-alpha',
      'lower-roman',
      'upper-roman',
      'todo',
    ].flatMap((key) => [`list.${key}.on`, `list.${key}.off`]),
    'list.check',
    'list.uncheck',
    'indent.increase',
    'indent.decrease',
    ...['start', 'left', 'center', 'right', 'end', 'justify'].map((key) => `align.${key}`),
    ...[1, 1.2, 1.5, 2, 3].map((key) => `lineHeight.${key}`),
    ...['insert', 'insert-internal', 'edit-url', 'edit-title', 'remove'].map(
      (key) => `link.${key}`
    ),
    ...[
      'insert',
      'insert-default',
      'row.before',
      'row.after',
      'row.delete',
      'column.before',
      'column.after',
      'column.delete',
      'delete',
      'merge',
      'split',
      'column.width',
      'row.height',
      'margin',
      'background',
      'borders',
      'append-row',
      'append-column',
      'append-row-column',
    ].map((key) => `table.${key}`),
  ];
  for (const id of required) expect(cases.has(id)).toBe(true);
  for (const key of ['h1', 'h2', 'h3']) expect(cases.has(`block.${key}.empty-break`)).toBe(true);
  for (const key of [
    'heading1',
    'heading2',
    'heading3',
    'quote',
    'rule-dash',
    'rule-underscore',
    'bullet-dash',
    'bullet-star',
    'number-dot',
    'number-paren',
    'todo-open',
    'todo-checked',
    'bold-star',
    'bold-underscore',
    'italic-star',
    'italic-underscore',
    'underline',
    'code',
    'strike',
    'subscript',
    'superscript',
    'bold-italic',
    'bold-underline',
    'bold-italic-underline',
    'italic-underline',
  ])
    expect(cases.has(`input.${key}`)).toBe(true);
});

test('production kits, consumers, configured input rules, and resolved shortcuts retain provenance', () => {
  for (const kit of [
    'basic-blocks',
    'basic-marks',
    'font',
    'list',
    'indent',
    'align',
    'line-height',
    'link',
    'table',
  ])
    expect(Object.hasOwn(fixture.source.files, `src/components/${kit}-kit.tsx`)).toBe(true);
  for (const source of [
    'src/components/transforms.ts',
    'src/components/workspace/WorkspaceEditorKit.ts',
    'src/components/ui/table-node.tsx',
    'src/components/ui/table-toolbar-button.tsx',
    'src/components/ui/slash-catalog.ts',
    'src/components/ui/turn-into-toolbar-config.tsx',
    'src/components/ui/font-color-toolbar-button.tsx',
    'src/components/ui/link-toolbar.tsx',
  ])
    expect(Object.hasOwn(fixture.source.files, source)).toBe(true);
  for (const entry of fixture.inventory) {
    expect(Object.hasOwn(fixture.source.files, entry.source)).toBe(true);
    expect(entry.cases.length).toBeGreaterThan(0);
    for (const id of entry.cases) expect(cases.has(id)).toBe(true);
  }
  expect(
    fixture.inventory.filter((entry) => entry.kind === 'inputRules').length
  ).toBeGreaterThanOrEqual(10);
  for (const key of [
    'p',
    'h1',
    'h2',
    'h3',
    'bold',
    'italic',
    'underline',
    'code',
    'strikethrough',
    'subscript',
    'superscript',
    'a',
  ]) {
    const entry = fixture.keyboard.find((entry) => entry.key === key);
    expect(entry).toBeDefined();
    expect(Object.keys(entry!.shortcuts).length).toBeGreaterThan(0);
  }
  const hash = createHash('sha256')
    .update(readFileSync(resolve(root, 'scripts/sciobot-port/native-command-runner.mjs')))
    .digest('hex');
  expect(fixture.source.runnerHash).toBe(hash);
  expect(fixture.source.captureHash).toBe(
    createHash('sha256')
      .update(
        readFileSync(resolve(root, 'scripts/sciobot-port/capture-native-commands-reference.mjs'))
      )
      .digest('hex')
  );
});
