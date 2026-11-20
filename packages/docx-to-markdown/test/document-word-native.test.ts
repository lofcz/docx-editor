import { expect, test } from 'bun:test';
import {
  TreeDocumentStore,
  contentControlPropertiesOf,
  paragraphTextOf,
  readOoxmlPackage,
  validatePackageInvariants,
  withPart,
  writeOoxmlPackage,
  writeZip,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { readWordDocument, writeWordDocument } from '../src/mdx/document-word';
import { parseDocumentSyntax, printDocumentSyntax } from '../src/mdx/document-syntax';
import { type DocumentNode } from '../src/mdx/document-types';

const native = () => undefined;
function open(bytes: Uint8Array): OoxmlPackage {
  const result = readOoxmlPackage(bytes);
  if (!result.ok) throw new Error(result.reason);
  expect(validatePackageInvariants(result.package).ok).toBe(true);
  return result.package;
}
function write(source: string) {
  return writeWordDocument(parseDocumentSyntax(source), native);
}
function ofKind(node: OoxmlNode, kind: OoxmlNode['kind']): OoxmlNode[] {
  return [
    ...(node.kind === kind ? [node] : []),
    ...(node.kind === 'textValue' ? [] : node.children.flatMap((n) => ofKind(n, kind))),
  ];
}

const handbook = [
  '# Assembly guide',
  '',
  'Use **firm pressure**, *light oil*, and ~~old parts~~. Read the [procedure](https://example.org/a?x=1&y=2 "Manual").',
  '',
  '*Read **both steps** in [the notes](#notes).* Keep ~~*old* parts~~ separate.',
  '',
  '3. Inspect the housing.',
  '   - [x] Measure the shaft.',
  '   - [ ] Check the seal.',
  '4. Record the result.',
  '',
  '| Part | Limit |',
  '| :--- | ---: |',
  '| Pin | `2.4 mm` |',
  '| Bolt | **4 mm** |',
].join('\n');
const fieldNotes = [
  '## Site notes',
  '',
  '> First reading.',
  '>',
  '> Second reading with [notes][guide].',
  '',
  '[guide]: ../guide.md "Guide"',
  '',
  'Collect an image: *![Site A](https://example.org/remote.png "Overview")*.',
  '',
  '---',
  '',
  '```text sample="B"',
  'Station\t12',
  '',
  'End & <$&',
  '```',
  '',
  'A soft',
  'break and a hard\\',
  'break.',
  '',
  '- One paragraph.',
  '',
  '  A second paragraph.',
  '',
  '- Final entry.',
].join('\n');

test.each([handbook, fieldNotes])('native content survives save and reopen %#', (source) => {
  const files = write(source);
  expect([...files.keys()].every((key) => key.startsWith('/'))).toBe(true);
  const pkg = open(writeZip(files));
  const restored = readWordDocument(open(writeOoxmlPackage(pkg)), native);
  expect(parseDocumentSyntax(printDocumentSyntax(restored))).toEqual(parseDocumentSyntax(source));
  const document = new TextDecoder().decode(files.get('/word/document.xml'));
  expect(document).toContain('<w:sectPr>');
  expect(document).toContain('<w:numPr>');
  if (source === handbook) {
    expect(document).toContain('<w:hyperlink r:id=');
    expect(ofKind(pkg.parts.get(pkg.mainDocumentPart)!.root, 'tableCell')).toHaveLength(6);
    expect(document).toContain('w:val="Heading1"');
    expect(document).toContain('<w14:checkbox>');
    expect(document).toContain('<w:b/>');
    expect(document).toContain('<w:i/>');
    expect(document).toContain('<w:strike/>');
    expect(new TextDecoder().decode(files.get('/word/numbering.xml'))).toContain(
      '<w:start w:val="3"/>'
    );
  } else {
    expect(pkg.externalTargets).toHaveLength(0);
    expect(document).toContain('md:relative="../guide.md"');
    expect(document).not.toContain('<w:drawing');
  }
});

test('exports current canonical text and rows before saving and after reopening', () => {
  const original = open(writeZip(write(handbook)));
  const store = new TreeDocumentStore(original.parts.get(original.mainDocumentPart)!);
  const table = ofKind(store.part.root, 'table')[0]!;
  const rows = ofKind(table, 'tableRow');
  const first = ofKind(store.part.root, 'paragraph')[0]!;
  const edit = store.transact((ctx) => {
    ctx.apply({ op: 'insertText', paragraphId: first.id, offset: 0, text: 'Updated ' });
    ctx.apply({ op: 'insertTableRow', tableId: table.id, rowId: rows[2].id, where: 'below' });
    ctx.apply({ op: 'deleteTableRow', tableId: table.id, rowId: rows[1].id });
  });
  expect(edit.ok).toBe(true);
  const newRow = ofKind(store.part.root, 'tableRow').at(-1)!;
  const p = ofKind(newRow, 'paragraph')[0]!;
  expect(
    store.transact((ctx) =>
      ctx.apply({ op: 'insertText', paragraphId: p.id, offset: 0, text: 'Nut' })
    ).ok
  ).toBe(true);
  const current = withPart(original, store.part);
  const expected = parseDocumentSyntax(
    handbook
      .replace('Assembly guide', 'Updated Assembly guide')
      .replace('| Pin | `2.4 mm` |\n', '')
      .replace('| Bolt | **4 mm** |', '| Bolt | **4 mm** |\n| Nut | |')
  );
  for (const pkg of [current, open(writeOoxmlPackage(current))]) {
    expect(parseDocumentSyntax(printDocumentSyntax(readWordDocument(pkg, native)))).toEqual(
      expected
    );
  }
});

test('nested native controls preserve paragraph boundaries and use current inline text', () => {
  const source = [
    '<Section>',
    '',
    '## Scope',
    '',
    '<Note>',
    '',
    'First <Term>word</Term>.',
    '',
    'Second paragraph.',
    '',
    '</Note>',
    '',
    '</Section>',
  ].join('\n');
  const tree = parseDocumentSyntax(source);
  const calls: string[] = [];
  const files = writeWordDocument(tree, (node, inline, children) => {
    calls.push(node.type);
    if (!node.name || !['Section', 'Note', 'Term'].includes(node.name)) return undefined;
    expect(inline).toBe(node.name === 'Term');
    return (
      '<w:sdt><w:sdtPr><w:tag w:val="' +
      node.name +
      '"/></w:sdtPr><w:sdtContent>' +
      children() +
      '</w:sdtContent></w:sdt>'
    );
  });
  expect(calls).toContain('root');
  expect(calls).toContain('text');
  const pkg = open(writeZip(files));
  const store = new TreeDocumentStore(pkg.parts.get(pkg.mainDocumentPart)!);
  const p = ofKind(store.part.root, 'paragraph').find(
    (node) => paragraphTextOf(store.part, node.id) === 'First word.'
  )!;
  expect(
    store.transact((ctx) =>
      ctx.apply({ op: 'insertText', paragraphId: p.id, offset: 8, text: 'L' })
    ).ok
  ).toBe(true);
  const nodes = readWordDocument(withPart(pkg, store.part), (node, inline, children) => {
    const tag = contentControlPropertiesOf(node).tag;
    if (tag === 'Note') return undefined;
    return [
      {
        type: inline ? 'mdxJsxTextElement' : 'mdxJsxFlowElement',
        name: tag,
        attributes: [],
        children: children(),
      },
    ];
  });
  const expected = parseDocumentSyntax(
    source.replace('<Note>\n\n', '').replace('\n\n</Note>', '').replace('word', 'woLrd')
  );
  expect(parseDocumentSyntax(printDocumentSyntax(nodes))).toEqual(expected);
});

test('a flow control with inline children contains an editable Word paragraph', () => {
  const flow: DocumentNode = {
    type: 'exampleFlow',
    children: [{ type: 'text', value: 'Local text' }],
  };
  const entries = writeWordDocument(
    { type: 'root', children: [flow] },
    (node, inline, children) => {
      if (node !== flow) return undefined;
      expect(inline).toBe(false);
      return '<w:sdt><w:sdtPr/><w:sdtContent>' + children() + '</w:sdtContent></w:sdt>';
    }
  );
  const pkg = open(writeZip(entries));
  expect(ofKind(pkg.parts.get(pkg.mainDocumentPart)!.root, 'paragraph')).toHaveLength(1);
  expect(readWordDocument(pkg, native)).toEqual(parseDocumentSyntax('Local text'));
});

test('unsafe link targets, malformed text, and unsupported shapes fail explicitly', () => {
  const one = (node: DocumentNode) =>
    writeWordDocument(
      { type: 'root', children: [{ type: 'paragraph', children: [node] }] },
      native
    );
  expect(() => one({ type: 'link', url: 'java\nscript:alert(1)', children: [] })).toThrow(
    'link destination'
  );
  expect(() => one({ type: 'text', value: '\0' })).toThrow('Invalid Markdown text');
  expect(() => one({ type: 'unknownInline' })).toThrow('Unsupported inline');
  const loop: DocumentNode = { type: 'emphasis' };
  loop.children = [loop];
  expect(() => one(loop)).toThrow('oversized Markdown tree');
});

test('links to duplicate headings use real bookmarks and follow native heading edits', () => {
  const source = '[See the next test](#test-result-1)\n\n## Test result\n\n## Test result';
  const entries = write(source);
  const document = new TextDecoder().decode(entries.get('/word/document.xml'));
  expect(document).toContain('w:anchor="MarkdownHeading1"');
  expect(document).toContain('w:name="MarkdownHeading1"');
  const pkg = open(writeZip(entries));
  expect(parseDocumentSyntax(printDocumentSyntax(readWordDocument(pkg, native)))).toEqual(
    parseDocumentSyntax(source)
  );
  const store = new TreeDocumentStore(pkg.parts.get(pkg.mainDocumentPart)!);
  const target = ofKind(store.part.root, 'paragraph').at(-1)!;
  expect(
    store.transact((ctx) =>
      ctx.apply({ op: 'insertText', paragraphId: target.id, offset: 0, text: 'Final ' })
    ).ok
  ).toBe(true);
  const exported = printDocumentSyntax(readWordDocument(withPart(pkg, store.part), native));
  expect(exported).toContain('[See the next test](#final-test-result)');
  expect(exported).toContain('## Final Test result');
});
