// This runner is copied to an isolated directory with source-derived kit modules.
import { writeFileSync } from 'node:fs';
import { createPlateEditor } from 'platejs/react';
import { KEYS } from 'platejs';
import { toggleList, indentTodo } from '@platejs/list';
import { indent, outdent } from '@platejs/indent';
import { upsertLink, unwrapLink } from '@platejs/link';
import * as table from '@platejs/table';
import { BasicBlocksKit } from './basic-blocks.tsx';
import { BasicMarksKit } from './basic-marks.tsx';
import { FontKit } from './font.tsx';
import { ListKit } from './list.tsx';
import { AlignKit } from './align.tsx';
import { LineHeightKit } from './line-height.tsx';
import { TableKit } from './table.tsx';
import { LinkKit } from './link.tsx';
import { setBlockType, insertBlock } from './block.ts';
const plugins = [
  ...BasicBlocksKit,
  ...BasicMarksKit,
  ...FontKit,
  ...ListKit,
  ...AlignKit,
  ...LineHeightKit,
  ...TableKit,
  ...LinkKit,
];
const p = (text = 'Alpha beta') => ({ type: 'p', children: [{ text }] });
const range = (path = [0, 0], end = 5) => ({
  anchor: { path, offset: 0 },
  focus: { path, offset: end },
});
const clone = (value) => JSON.parse(JSON.stringify(value));
// IDs are runtime identity, not command semantics; deterministic labels retain aliasing.
const clean = (value, ids = new Map()) => {
  const visit = (value) =>
    Array.isArray(value)
      ? value.map(visit)
      : value && typeof value === 'object'
        ? Object.fromEntries(
            Object.entries(value).map(([key, value]) => [
              key,
              key === 'id' && typeof value === 'string'
                ? ids.has(value)
                  ? ids.get(value)
                  : (ids.set(value, `id-${ids.size + 1}`), ids.get(value))
                : visit(value),
            ])
          )
        : value;
  return visit(value);
};
const cases = [];
function capture(id, group, leaf, action, value = [p()], selection = range(), setup) {
  const editor = createPlateEditor({ plugins, value: clone(value) });
  editor.tf.select(selection);
  if (setup) setup(editor);
  const ids = new Map();
  const before = clean(
    {
      value: editor.children,
      selection: editor.selection,
      marks: editor.api.marks(),
    },
    ids
  );
  action(editor);
  const after = clean(
    {
      value: editor.children,
      selection: editor.selection,
      marks: editor.api.marks(),
    },
    ids
  );
  const record = {
    id,
    group,
    command: id,
    operation: action.toString(),
    mapping: { kind: 'implementation-leaf', leaf },
    before,
    after,
  };
  cases.push(record);
  return record;
}
for (const type of ['p', 'h1', 'h2', 'h3', 'blockquote']) {
  capture(`block.${type}`, 'BasicBlocks', 'N02', (e) => setBlockType(e, type), [
    { type: type === 'p' ? 'h1' : 'p', children: [{ text: 'Alpha beta' }] },
  ]);
  capture(`block.${type}.from-list`, 'BasicBlocks', 'N02', (e) => setBlockType(e, type), [
    { ...p(), indent: 1, listStyleType: 'disc' },
  ]);
}
for (const type of ['p', 'h1', 'h2', 'h3', 'blockquote', 'hr'])
  capture(`block.insert.${type}`, 'BasicBlocks', 'N02', (e) => insertBlock(e, type));
for (const type of ['disc', 'decimal', 'todo'])
  capture(`list.insert.${type}`, 'List', 'N08', (e) => insertBlock(e, type));
capture('table.insert-default', 'Table', 'N09', (e) => insertBlock(e, 'table'));
for (const type of ['h1', 'h2', 'h3'])
  capture(
    `block.${type}.empty-break`,
    'BasicBlocks',
    'N02',
    (e) => e.tf.insertBreak(),
    [{ type, children: [{ text: '' }] }],
    range([0, 0], 0)
  );
capture('block.hr', 'BasicBlocks', 'N02', (e) =>
  e.tf.insertNodes({ type: 'hr', children: [{ text: '' }] }, { at: [1], select: true })
);
for (const mark of [
  'bold',
  'italic',
  'underline',
  'code',
  'strikethrough',
  'subscript',
  'superscript',
]) {
  capture(`mark.${mark}.on`, 'BasicMarks', 'N03', (e) => e.tf.toggleMark(mark));
  capture(`mark.${mark}.off`, 'BasicMarks', 'N03', (e) => e.tf.toggleMark(mark), [
    { type: 'p', children: [{ text: 'Alpha beta', [mark]: true }] },
  ]);
  capture(
    `mark.${mark}.caret`,
    'BasicMarks',
    'N03',
    (e) => {
      e.tf.toggleMark(mark);
      e.tf.insertText('X');
    },
    [p()],
    range([0, 0], 0)
  );
}
for (const [key, value, leaf] of [
  ['fontFamily', 'Arial', 'N04'],
  ['fontSize', '24px', 'N04'],
  ['color', '#cc0000', 'N05'],
  ['backgroundColor', '#ffff00', 'N05'],
]) {
  capture(`font.${key}.set`, 'Font', leaf, (e) => e.getTransforms(key)[key].addMark(value));
  capture(`font.${key}.clear`, 'Font', leaf, (e) => e.tf.removeMarks(key), [
    { type: 'p', children: [{ text: 'Alpha beta', [key]: value }] },
  ]);
}
capture('font.color.inherit', 'Font', 'N05', (e) => e.tf.addMarks({ [KEYS.color]: undefined }), [
  { type: 'p', children: [{ text: 'Alpha beta', color: '#cc0000' }] },
]);
capture(
  'font.background.transparent',
  'Font',
  'N05',
  (e) => e.tf.addMarks({ [KEYS.backgroundColor]: undefined }),
  [{ type: 'p', children: [{ text: 'Alpha beta', backgroundColor: '#ffff00' }] }]
);
capture('font.size.increase', 'Font', 'N04', (e) =>
  e.getTransforms('fontSize').fontSize.addMark(`${16 + 1}px`)
);
capture('font.size.decrease', 'Font', 'N04', (e) =>
  e.getTransforms('fontSize').fontSize.addMark(`${16 - 1}px`)
);
for (const size of [
  '8',
  '9',
  '10',
  '12',
  '14',
  '16',
  '18',
  '24',
  '30',
  '36',
  '48',
  '60',
  '72',
  '96',
  '15',
  '17',
])
  capture(`font.size.${size}`, 'Font', 'N04', (e) =>
    e.getTransforms('fontSize').fontSize.addMark(`${size}px`)
  );
for (const style of [
  'disc',
  'circle',
  'square',
  'decimal',
  'lower-alpha',
  'upper-alpha',
  'lower-roman',
  'upper-roman',
  'todo',
]) {
  capture(`list.${style}.on`, 'List', 'N08', (e) =>
    style === 'todo' ? indentTodo(e) : toggleList(e, { listStyleType: style })
  );
  capture(
    `list.${style}.off`,
    'List',
    'N08',
    (e) => (style === 'todo' ? indentTodo(e) : toggleList(e, { listStyleType: style })),
    [{ ...p(), indent: 1, listStyleType: style, ...(style === 'todo' ? { checked: false } : {}) }]
  );
}
for (const type of ['disc', 'decimal', 'todo'])
  capture(`list.turnInto.${type}`, 'List', 'N08', (e) => setBlockType(e, type));
capture('list.check', 'List', 'N08', (e) => e.tf.setNodes({ checked: true }, { at: [0] }), [
  { ...p(), indent: 1, listStyleType: 'todo', checked: false },
]);
capture('list.uncheck', 'List', 'N08', (e) => e.tf.setNodes({ checked: false }, { at: [0] }), [
  { ...p(), indent: 1, listStyleType: 'todo', checked: true },
]);
capture('indent.increase', 'List', 'N06', (e) => indent(e));
capture('indent.decrease', 'List', 'N06', (e) => outdent(e), [{ ...p(), indent: 2 }]);
for (const align of ['start', 'left', 'center', 'right', 'end', 'justify'])
  capture(`align.${align}`, 'Align', 'N06', (e) =>
    e.getTransforms('align').textAlign.setNodes(align)
  );
for (const height of [1, 1.2, 1.5, 2, 3])
  capture(`lineHeight.${height}`, 'LineHeight', 'N07', (e) =>
    e.getTransforms('lineHeight').lineHeight.setNodes(height)
  );
capture('link.insert', 'Link', 'N10', (e) =>
  upsertLink(e, { url: 'https://example.com', skipValidation: true })
);
capture('link.insert-internal', 'Link', 'N10', (e) =>
  upsertLink(e, { url: '/document-123', skipValidation: true })
);
const linked = [
  {
    type: 'p',
    children: [
      { text: '' },
      { type: 'a', url: 'https://old.example', children: [{ text: 'Alpha' }] },
      { text: ' beta' },
    ],
  },
];
capture(
  'link.edit-url',
  'Link',
  'N10',
  (e) => upsertLink(e, { url: 'https://new.example', skipValidation: true }),
  linked,
  range([0, 1, 0])
);
capture(
  'link.edit-title',
  'Link',
  'N10',
  (e) => e.tf.replaceNodes({ text: 'New title' }, { at: [0, 1], children: true }),
  linked,
  range([0, 1, 0])
);
capture('link.remove', 'Link', 'N10', (e) => unwrapLink(e), linked, range([0, 1, 0]));
const cell = (text) => ({ type: 'td', children: [p(text)] });
const grid = () => [
  {
    type: 'table',
    children: [
      { type: 'tr', children: [cell('A'), cell('B')] },
      { type: 'tr', children: [cell('C'), cell('D')] },
    ],
  },
  p('After'),
];
const trange = range([0, 0, 0, 0, 0], 0);
capture(
  'table.insert',
  'Table',
  'N09',
  (e) => e.tf.insert.table({ rowCount: 2, colCount: 3 }, { select: true }),
  [p('')],
  range([0, 0], 0)
);
for (const [name, fn, leaf] of [
  ['row.before', (e) => e.tf.insert.tableRow({ before: true }), 'N11'],
  ['row.after', (e) => e.tf.insert.tableRow(), 'N11'],
  ['row.delete', (e) => e.tf.remove.tableRow(), 'N11'],
  ['column.before', (e) => e.tf.insert.tableColumn({ before: true }), 'N12'],
  ['column.after', (e) => e.tf.insert.tableColumn(), 'N12'],
  ['column.delete', (e) => e.tf.remove.tableColumn(), 'N12'],
  ['delete', (e) => e.tf.remove.table(), 'N09'],
  ['column.width', (e) => table.setTableColSize(e, { colIndex: 0, width: 180 }), 'N14'],
  ['row.height', (e) => table.setTableRowSize(e, { rowIndex: 0, height: 60 }), 'N14'],
  ['margin', (e) => table.setTableMarginLeft(e, { marginLeft: 24 }), 'N14'],
  ['background', (e) => table.setCellBackground(e, { color: '#ffff00' }), 'N14'],
  ['borders', (e) => table.setBorderSize(e, 2, { border: 'all' }), 'N14'],
  ['append-row', (e) => e.tf.insert.tableRow({ at: [0] }), 'N11'],
  ['append-column', (e) => e.tf.insert.tableColumn({ at: [0] }), 'N12'],
  [
    'append-row-column',
    (e) =>
      e.tf.withoutNormalizing(() => {
        e.tf.insert.tableRow({ at: [0] });
        e.tf.insert.tableColumn({ at: [0] });
      }),
    'N09',
  ],
])
  capture(`table.${name}`, 'Table', leaf, fn, grid(), trange);
const mergedSelection = {
  anchor: { path: [0, 0, 0, 0, 0], offset: 0 },
  focus: { path: [0, 0, 1, 0, 0], offset: 1 },
};
capture('table.merge', 'Table', 'N13', (e) => e.tf.table.merge(), grid(), mergedSelection);
capture(
  'table.split',
  'Table',
  'N13',
  (e) => e.tf.table.split(),
  grid(),
  mergedSelection,
  (e) => e.tf.table.merge()
);
// Execute configured input rules through the real insertText pipeline.
for (const [id, text, group, leaf] of [
  ['heading1', '# ', 'BasicBlocks', 'N02'],
  ['heading2', '## ', 'BasicBlocks', 'N02'],
  ['heading3', '### ', 'BasicBlocks', 'N02'],
  ['quote', '> ', 'BasicBlocks', 'N02'],
  ['rule-dash', '--- ', 'BasicBlocks', 'N02'],
  ['rule-underscore', '___ ', 'BasicBlocks', 'N02'],
  ['bullet-dash', '- ', 'List', 'N08'],
  ['bullet-star', '* ', 'List', 'N08'],
  ['number-dot', '1. ', 'List', 'N08'],
  ['number-paren', '1) ', 'List', 'N08'],
  ['todo-open', '[] ', 'List', 'N08'],
  ['todo-checked', '[x] ', 'List', 'N08'],
  ['bold-star', '**word**', 'BasicMarks', 'N03'],
  ['bold-underscore', '__word__', 'BasicMarks', 'N03'],
  ['italic-star', '*word*', 'BasicMarks', 'N03'],
  ['italic-underscore', '_word_', 'BasicMarks', 'N03'],
  ['underline', '__word__', 'BasicMarks', 'N03'],
  ['code', '`word`', 'BasicMarks', 'N03'],
  ['strike', '~~word~~', 'BasicMarks', 'N03'],
  ['subscript', '~word~', 'BasicMarks', 'N03'],
  ['superscript', '^word^', 'BasicMarks', 'N03'],
  ['bold-italic', '**word**', 'BasicMarks', 'N03'],
  ['bold-italic-triple', '***word***', 'BasicMarks', 'N03'],
  ['bold-underline', '__word***', 'BasicMarks', 'N03'],
  ['bold-italic-underline', '___word****', 'BasicMarks', 'N03'],
  ['italic-underline', '__word**', 'BasicMarks', 'N03'],
])
  capture(
    `input.${id}`,
    group,
    leaf,
    (e) => {
      for (const character of text) e.tf.insertText(character);
    },
    [p('')],
    range([0, 0], 0)
  ).arguments = { text };
const argumentsById = {
  'font.fontFamily.set': { value: 'Arial' },
  'font.fontSize.set': { value: '24px' },
  'font.color.set': { value: '#cc0000' },
  'font.backgroundColor.set': { value: '#ffff00' },
  'font.size.increase': { current: 16, delta: 1, unit: 'px' },
  'font.size.decrease': { current: 16, delta: -1, unit: 'px' },
  'link.insert': { url: 'https://example.com', skipValidation: true },
  'link.insert-internal': { url: '/document-123', skipValidation: true },
  'link.edit-url': { url: 'https://new.example', skipValidation: true },
  'link.edit-title': { text: 'New title' },
  'table.insert': { rowCount: 2, colCount: 3, select: true },
  'table.column.width': { colIndex: 0, width: 180 },
  'table.row.height': { rowIndex: 0, height: 60 },
  'table.margin': { marginLeft: 24 },
  'table.background': { color: '#ffff00' },
  'table.borders': { size: 2, border: 'all' },
};
for (const entry of cases) {
  entry.arguments ??= argumentsById[entry.id] ?? { variant: entry.id };
  if (entry.id === 'table.margin') entry.mapping = { kind: 'implementation-leaf', leaf: 'N22' };
  if (entry.id === 'table.delete')
    entry.mapping = {
      kind: 'native-command',
      command: 'deleteTable',
      source: 'packages/core/src/contracts/editor.ts',
    };
  if (entry.id === 'table.append-row-column')
    entry.mapping = { kind: 'implementation-leaf', leaf: 'N11', additionalLeaves: ['N12'] };
  if (entry.id === 'align.start' || entry.id === 'align.end')
    entry.mapping.note =
      'N06 must resolve logical alignment per paragraph direction before calling physical setAlignment.';
  if (entry.id === 'link.insert-internal')
    entry.mapping.note =
      'N10 must preserve application document routing and apply native URL policy; this is not a claim of native relative-target support.';
}
const probe = createPlateEditor({ plugins, value: [p()] });
const keyboard = Object.entries(probe.plugins)
  .filter(([key]) =>
    [
      'p',
      'h1',
      'h2',
      'h3',
      'blockquote',
      'hr',
      'bold',
      'italic',
      'underline',
      'code',
      'strikethrough',
      'subscript',
      'superscript',
      'indent',
      'list',
      'table',
      'a',
    ].includes(key)
  )
  .map(([key, plugin]) => ({
    key,
    handlers: Object.keys(plugin.handlers ?? {}),
    shortcuts: clean(
      Object.fromEntries(
        Object.entries(plugin.shortcuts ?? {}).map(([name, value]) => [
          name,
          { keys: value.keys, hasHandler: typeof value.handler === 'function' },
        ])
      )
    ),
  }));
writeFileSync(
  'results.json',
  JSON.stringify({
    groups: ['BasicBlocks', 'BasicMarks', 'Font', 'List', 'Align', 'LineHeight', 'Link', 'Table'],
    keyboard,
    cases,
  })
);
