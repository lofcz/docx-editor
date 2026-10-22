import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  canonicalOoxmlFingerprint,
  escapeXml,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkMdx from 'remark-mdx';
import remarkGfm from 'remark-gfm';
import { docxFixture } from '../../test/support/document';
import fixture from '../__tests__/fixtures/mdx-tags.json';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload, JsonObject } from '../blocks/payload-schema';
import { blockAttributeAt } from './blocks';
import { markdownPointAt } from './markdown';
import {
  CUSTOM_ALLOWED_TAGS,
  CUSTOM_DATA_ONLY_TAGS,
  CUSTOM_LITERAL_TAGS,
  MDX_BLOCK,
  MDX_INLINE,
  projectMdx,
  serializeOpaqueMdx,
  type OpaqueMdxAst,
} from './mdx';

interface CapturedCase {
  id: string;
  markdown: string;
  type: typeof MDX_BLOCK | typeof MDX_INLINE;
  ast: OpaqueMdxAst;
  serialized: string;
}
const cases = fixture.cases as unknown as CapturedCase[];
const parser = unified().use(remarkParse).use(remarkGfm).use(remarkMdx);
const run = (text: string) => `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
const p = (text: string) => `<w:p>${run(text)}</w:p>`;
const control = (id: string, body: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
function payload(value: CapturedCase): BlockPayload {
  return { id: value.id, type: value.type, version: 1, attributes: {}, mdast: value.ast };
}
function captured(id: string): CapturedCase {
  const value = cases.find((entry) => entry.id === id);
  if (!value) throw new Error(`Missing reference case: ${id}`);
  return value;
}
function jsxNames(node: unknown): string[] {
  if (!node || typeof node !== 'object') return [];
  const value = node as { type?: string; name?: string; children?: unknown[] };
  return [
    ...(value.type?.startsWith('mdxJsx') && value.name ? [value.name] : []),
    ...(value.children ?? []).flatMap(jsxNames),
  ];
}

async function importedProjection(body: string, payloads: BlockPayload[]) {
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const value of payloads) {
    const written = withPayloadRecord(pkg, value, 'Readable Word fallback');
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  const surface = editor.surface!;
  const session = surface.session;
  const opened = openDocumentForExport(session, { displayMode: 'proposed' });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const before = [...session.currentPackage().parts].map(([name, part]) => [
      name,
      canonicalOoxmlFingerprint(part),
    ]);
    const firstParagraph = session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: firstParagraph, offset: 0 },
      head: { paragraphId: firstParagraph, offset: 0 },
    });
    const selection = surface.state().selection;
    const revision = { documentId: 'mdx', epoch: 'load-1', revision: session.packageRevision() };
    const result = projectMdx(session.currentPackage(), await opened.session.layout(), revision);
    expect(
      [...session.currentPackage().parts].map(([name, part]) => [
        name,
        canonicalOoxmlFingerprint(part),
      ])
    ).toEqual(before);
    expect(session.packageRevision()).toBe(revision.revision);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
    return {
      result,
      paragraphs: session
        .paragraphIds()
        .map((id) => ({ id, text: paragraphTextOf(session.part(), id) })),
      bytes: writeOoxmlPackage(session.currentPackage()),
    };
  } finally {
    opened.session.dispose();
    editor.destroy();
  }
}

test('runtime tags match captured application exports and cover both MDX identities', () => {
  expect(CUSTOM_ALLOWED_TAGS).toEqual(fixture.allowedTags);
  expect(Object.keys(CUSTOM_ALLOWED_TAGS)).toHaveLength(21);
  expect(CUSTOM_LITERAL_TAGS).toEqual(fixture.literalTags);
  expect(CUSTOM_DATA_ONLY_TAGS).toEqual(['suggestions', 'listening']);
  for (const tag of Object.keys(CUSTOM_ALLOWED_TAGS)) {
    expect(captured(`${tag}-block`).type).toBe(MDX_BLOCK);
    expect(captured(`${tag}-inline`).type).toBe(MDX_INLINE);
  }
});

test.each(cases)('opaque MDX matches the application serializer: $id', (value) => {
  const original = JSON.stringify(value.ast);
  const serialized = serializeOpaqueMdx(payload(value));
  expect(serialized.text).toBe(value.serialized);
  expect(serialized.ast).toEqual(value.ast);
  expect(serialized.ast === value.ast).toBe(false);
  expect(serialized.inline).toBe(value.type === MDX_INLINE);
  expect(Object.isFrozen(serialized.ast)).toBe(true);
  expect(Object.isFrozen(serialized.ast.attributes)).toBe(true);
  expect(JSON.stringify(value.ast)).toBe(original);
  const tree = parser.parse(
    serialized.inline ? `Before ${serialized.text} after.` : serialized.text
  );
  expect(jsxNames(tree)).toEqual(jsxNames(value.ast));
  const allowed = CUSTOM_ALLOWED_TAGS[value.ast.name ?? ''] ?? [];
  const attributes = serialized.ast.attributes.filter(
    (attribute) => attribute.type === 'mdxJsxAttribute'
  );
  for (const name of allowed) {
    // Every allowed attribute retains the captured value, including quotes, entities, and emoji.
    expect(attributes.find((attribute) => attribute.name === name)).toEqual(
      value.ast.attributes.find((attribute) => attribute.name === name)
    );
  }
});

test('native projection retains every stored AST and surrounding UTF-16 positions', async () => {
  const body =
    cases
      .map((value) =>
        value.type === MDX_INLINE
          ? `<w:p>${run('Before 😀 ')}${control(value.id, run(`fallback ${value.id}`))}${run(' after.')}</w:p>`
          : control(value.id, p(`fallback ${value.id}`))
      )
      .join('') + p('Final native text.');
  const { result, paragraphs, bytes } = await importedProjection(body, cases.map(payload));
  expect(result.success).toBe(true);
  if (!result.success) throw new Error(result.error);
  const projection = result.projection;
  expect(projection.opaqueMdx).toHaveLength(cases.length);
  for (const value of cases) {
    const span = projection.opaqueMdx.find((entry) => entry.blockId === value.id)!;
    expect(span.ast).toEqual(value.ast);
    expect(span.inline).toBe(value.type === MDX_INLINE);
    expect(projection.text.slice(span.start, span.end)).toBe(value.serialized);
    expect(blockAttributeAt(projection, span.start + 1)).toMatchObject({
      blockId: value.id,
      property: 'mdast',
      path: ['mdast'],
      start: span.start,
      end: span.end,
    });
    expect(markdownPointAt(projection, span.start + 1)).toBeNull();
    if (value.type === MDX_INLINE) {
      const native = paragraphs.find((paragraph) =>
        paragraph.text?.includes(`fallback ${value.id}`)
      )!;
      expect(projection.text.slice(span.end, span.end + ' after\\.'.length)).toBe(' after\\.');
      expect(markdownPointAt(projection, span.end + ' after\\.'.length, 'backward')).toEqual({
        sourceScope: 'body',
        paragraphId: native.id,
        offset: native.text!.length,
      });
    }
  }
  expect(projection.text).not.toContain('fallback');
  expect(jsxNames(parser.parse(projection.text))).toEqual(
    cases.flatMap((value) => jsxNames(value.ast))
  );
  const reopened = readOoxmlPackage(bytes);
  if (!reopened.ok) throw new Error(reopened.reason);
  const index = buildBlockIndex(reopened.package);
  for (const value of cases) expect(index.byId.get(value.id)?.payload.mdast).toEqual(value.ast);
}, 30_000);

test('rich MDX inside a Word cell uses HTML boundaries and retains native cell content', async () => {
  const listening = captured('nested-listening');
  const { result, paragraphs } = await importedProjection(
    p('Before table.') +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc>' +
      control(listening.id, p('Audio fallback')) +
      p('Same cell tail.') +
      '</w:tc><w:tc>' +
      p('Other cell.') +
      '</w:tc></w:tr></w:tbl>' +
      p('After table.'),
    [payload(listening)]
  );
  expect(result.success).toBe(true);
  if (!result.success) throw new Error(result.error);
  const projection = result.projection;
  expect(projection.text).toContain('<table>');
  expect(jsxNames(parser.parse(projection.text))).toEqual([
    'table',
    'tr',
    'td',
    'listening',
    'speaker',
    'line',
    'speaker',
    'line',
    'td',
  ]);
  expect(projection.opaqueMdx[0]?.ast).toEqual(listening.ast);
  expect(projection.text).not.toContain('Audio fallback');
  expect(projection.text).toContain('Same cell tail\\.');
  expect(projection.sourceMap.tables).toHaveLength(1);
  const tail = paragraphs.find((paragraph) => paragraph.text === 'Same cell tail.')!;
  expect(markdownPointAt(projection, projection.text.indexOf('Same cell tail') + 5)).toEqual({
    sourceScope: 'body',
    paragraphId: tail.id,
    offset: 5,
  });
});

test('unknown AST metadata, future tags, and prototype-shaped JSON stay inert and intact', () => {
  const original = captured('competence-block');
  const ast = JSON.parse(JSON.stringify(original.ast)) as OpaqueMdxAst;
  const extra = JSON.parse(
    '{"__proto__":{"polluted":"no"},"constructor":{"prototype":{"polluted":"no"}},"unknown":[1,true,null]}'
  );
  const changed = { ...payload(original), mdast: { ...ast, name: 'future-widget', data: extra } };
  const serialized = serializeOpaqueMdx(changed);
  expect(serialized.ast).toEqual(changed.mdast);
  expect(serialized.text).toContain('<future-widget');
  expect(serialized.ast.data).toEqual(extra);
  expect(Object.isFrozen((serialized.ast.data as JsonObject).unknown)).toBe(true);
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
});

test.each(['<Žák code="ABC">Goal</Žák>', '<>Fragment 😀</>'])(
  'valid future Unicode names and fragments remain opaque: %s',
  (markdown) => {
    const root = parser.parse(`Before ${markdown} after.`);
    const paragraph = root.children[0] as unknown as { children: JsonObject[] };
    const ast = paragraph.children.find((child) => child.type === 'mdxJsxTextElement')!;
    const serialized = serializeOpaqueMdx({
      id: 'future',
      type: MDX_INLINE,
      version: 1,
      attributes: {},
      mdast: JSON.parse(JSON.stringify(ast)) as JsonObject,
    });
    expect(serialized.text).toBe(markdown);
    expect(serialized.ast.name).toBe(ast.name as string | null);
  }
);

test('expressions and spreads serialize as source without executing code', () => {
  const markdown =
    '<listening {...globalThis.__sciobotMdxProbe()} title={globalThis.__sciobotMdxProbe()}>\n\n{globalThis.__sciobotMdxProbe()}\n\n</listening>';
  const ast = JSON.parse(JSON.stringify(parser.parse(markdown).children[0])) as OpaqueMdxAst;
  const scope = globalThis as typeof globalThis & { __sciobotMdxProbe?: () => void };
  let calls = 0;
  scope.__sciobotMdxProbe = () => {
    calls += 1;
  };
  try {
    const serialized = serializeOpaqueMdx({
      id: 'inert',
      type: MDX_BLOCK,
      version: 1,
      attributes: {},
      mdast: ast,
    });
    expect(serialized.text).toContain('{...globalThis.__sciobotMdxProbe()}');
    expect(serialized.text).toContain('title={globalThis.__sciobotMdxProbe()}');
    expect(serialized.ast).toEqual(ast);
    expect(calls).toBe(0);
  } finally {
    delete scope.__sciobotMdxProbe;
  }
});

test('invalid ASTs fail without changing the document or invoking getters', async () => {
  const value = captured('competence-block');
  let calls = 0;
  const dangerous = { ...payload(value) };
  Object.defineProperty(dangerous, 'type', {
    enumerable: true,
    get() {
      calls += 1;
      return MDX_BLOCK;
    },
  });
  expect(() => serializeOpaqueMdx(dangerous)).toThrow('non-json-property');
  expect(calls).toBe(0);
  expect(() =>
    serializeOpaqueMdx({ ...payload(value), mdast: { ...value.ast, type: 'root' } })
  ).toThrow('Invalid opaque MDX root');
  expect(() => serializeOpaqueMdx({ ...payload(value), type: MDX_INLINE })).toThrow(
    'inline identity'
  );
  const malformed = {
    ...payload(value),
    mdast: { ...value.ast, children: [{ type: 'unknown-node', value: 'Keep fallback' }] },
  };
  const { result } = await importedProjection(control(value.id, p('Keep Word fallback.')), [
    malformed,
  ]);
  expect(result).toMatchObject({ success: false, code: 'invalid-document' });
});
