import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { formatRow, formatTable } from '../../test/support/table-format-fixture';
import {
  projectSourcePackage,
  sourceRevision,
  targetSourceFixture,
} from '../../test/support/target-source-fixture';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import reference from '../__tests__/fixtures/target-parsing.json';
import { indexTargetSource } from '../diff/target-source';
import { normalizeLogicalBlocks } from '../diff/logical-blocks';
import { markdownPointAt, markdownRangeAt } from './markdown';
import { normalizeTargetSource } from './normalize-target';
import {
  parseTargetMarkdown,
  targetSourceOffsetAt,
  targetSourceRangeAt,
  type ParsedTargetMarkdown,
  type TargetAstNode,
} from './parse-target';

function parsed(source: string): ParsedTargetMarkdown {
  const result = parseTargetMarkdown(source);
  if (!result.success) throw new Error(result.error);
  return result.target;
}
const unchanged = (start: number, end: number) => ({ start, end, exact: true, generated: false });
function nodes(ast: TargetAstNode): TargetAstNode[] {
  return [
    ast,
    ...(Array.isArray(ast.children) ? (ast.children as TargetAstNode[]).flatMap(nodes) : []),
  ];
}
function frozen(value: unknown): boolean {
  return (
    value === null ||
    typeof value !== 'object' ||
    (Object.isFrozen(value) && Object.values(value).every(frozen))
  );
}
function nativeTableBytes() {
  return docxFixture(
    formatTable(formatRow('A', 'B', 'C')).replace('>A</w:t>', '>A</w:t><w:br/><w:t>After</w:t>')
  );
}

test('unchanged UTF-16 strings and empty targets keep identity coordinates', () => {
  for (const source of [
    '',
    'Plain text',
    'A **bold** link [here](https://example.org).',
    '<br />',
  ]) {
    const target = parsed(source);
    expect(target.markdown).toBe(source);
    expect(target.normalization).toEqual({
      unit: 'utf16-code-unit',
      steps: [
        { sourceLength: source.length, targetLength: source.length, changes: [] },
        { sourceLength: source.length, targetLength: source.length, changes: [] },
      ],
    });
    for (let i = 0; i <= source.length; i++) expect(targetSourceOffsetAt(target, i)).toBe(i);
    expect(targetSourceRangeAt(target, 0, source.length)).toEqual(unchanged(0, source.length));
  }
});

test('length-changing HTML breaks retain exact tag edges and reject every rewritten interior boundary', () => {
  const target = parsed('A<br>B');
  expect(target.markdown).toBe('A<br />B');
  expect(target.normalization.steps).toEqual([
    {
      sourceLength: 6,
      targetLength: 8,
      changes: [{ sourceStart: 1, sourceEnd: 5, targetStart: 1, targetEnd: 7 }],
    },
    { sourceLength: 8, targetLength: 8, changes: [] },
  ]);
  expect(Array.from({ length: 9 }, (_, index) => targetSourceOffsetAt(target, index))).toEqual([
    0,
    1,
    null,
    null,
    null,
    null,
    null,
    5,
    6,
  ]);
  expect(targetSourceRangeAt(target, 1, 7)).toEqual({
    start: 1,
    end: 5,
    exact: false,
    generated: false,
  });
  expect(targetSourceRangeAt(target, 7, 8)).toEqual(unchanged(5, 6));
  expect(targetSourceRangeAt(target, 0, 8)).toEqual({
    start: 0,
    end: 6,
    exact: false,
    generated: false,
  });
  expect(targetSourceRangeAt(target, 2, 8)).toBeNull();
});

test('equal-length rewrites still mark semantic boundaries instead of pretending to be unchanged', () => {
  const target = parsed('A<BR  >B');
  expect(target.markdown).toBe('A<br />B');
  expect(target.sourceMarkdown.length).toBe(target.markdown.length);
  expect(target.normalization.steps[0]!.changes).toEqual([
    { sourceStart: 1, sourceEnd: 7, targetStart: 1, targetEnd: 7 },
  ]);
  expect(targetSourceRangeAt(target, 1, 7)).toEqual({
    start: 1,
    end: 7,
    exact: false,
    generated: false,
  });
  expect(targetSourceOffsetAt(target, 4)).toBeNull();
  expect(targetSourceRangeAt(target, 7, 8)).toEqual(unchanged(7, 8));
});

test('multiline HTML block breaks retain source line offsets and shrinking tag boundaries', () => {
  const source = '<div>\n<br\n >B\n</div>',
    target = parsed(source);
  expect(target.markdown).toBe('<div>\n<br />B\n</div>');
  expect(target.normalization.steps[0]!.changes).toEqual([
    { sourceStart: 6, sourceEnd: 12, targetStart: 6, targetEnd: 12 },
  ]);
  expect(targetSourceRangeAt(target, 6, 12)).toEqual({
    start: 6,
    end: 12,
    exact: false,
    generated: false,
  });
  expect(targetSourceRangeAt(target, 12, 13)).toEqual(unchanged(12, 13));
  const shrinking = parsed('<div>\n<br   >B\n</div>');
  expect(shrinking.normalization.steps[0]!.changes).toEqual([
    { sourceStart: 6, sourceEnd: 13, targetStart: 6, targetEnd: 12 },
  ]);
  expect(targetSourceRangeAt(shrinking, 12, 13)).toEqual(unchanged(13, 14));
  // The production normalizer leaves a multiline inline tag outside an HTML block unchanged.
  expect(normalizeTargetSource('A<br\n >B').normalization.steps[0]!.changes).toEqual([]);
});

test('multiple adjacent and separated breaks retain original owners without ambiguous text alignment', () => {
  const target = parsed('X<br><br >X<BR>X');
  expect(target.markdown).toBe('X<br /><br />X<br />X');
  expect(target.normalization.steps[0]!.changes).toEqual([
    { sourceStart: 1, sourceEnd: 5, targetStart: 1, targetEnd: 7 },
    { sourceStart: 5, sourceEnd: 10, targetStart: 7, targetEnd: 13 },
    { sourceStart: 11, sourceEnd: 15, targetStart: 14, targetEnd: 20 },
  ]);
  expect([0, 1, 7, 13, 14, 20, 21].map((offset) => targetSourceOffsetAt(target, offset))).toEqual([
    0, 1, 5, 10, 11, 15, 16,
  ]);
  expect(targetSourceRangeAt(target, 13, 14)).toEqual(unchanged(10, 11));
  expect(targetSourceRangeAt(target, 20, 21)).toEqual(unchanged(15, 16));
});

test('inserted ordered-list separators map to explicit generated zero-length source spans', () => {
  const target = parsed('Label:\n6. First\n7. Second');
  expect(target.markdown).toBe('Label:\n\n6. First\n7. Second');
  expect(target.normalization.steps[1]).toEqual({
    sourceLength: 25,
    targetLength: 26,
    changes: [{ sourceStart: 7, sourceEnd: 7, targetStart: 7, targetEnd: 8 }],
  });
  expect(targetSourceOffsetAt(target, 7)).toBe(7);
  expect(targetSourceOffsetAt(target, 8)).toBe(7);
  expect(targetSourceRangeAt(target, 7, 8)).toEqual({
    start: 7,
    end: 7,
    exact: false,
    generated: true,
  });
  expect(targetSourceRangeAt(target, 8, 16)).toEqual(unchanged(7, 15));
  expect(targetSourceRangeAt(target, 0, 26)).toEqual({
    start: 0,
    end: 25,
    exact: false,
    generated: true,
  });
  const list = nodes(target.ast).find((node) => node.type === 'list')!;
  const position = list.position as { start: { offset: number }; end: { offset: number } };
  expect(targetSourceRangeAt(target, position.start.offset, position.end.offset)).toEqual(
    unchanged(7, 25)
  );
});

test('several separated list insertions retain cumulative offsets in original input order', () => {
  const source = 'One:\n6. A\n7. B\n\nTwo:\n8. C\n9. D',
    target = parsed(source);
  expect(target.markdown).toBe('One:\n\n6. A\n7. B\n\nTwo:\n\n8. C\n9. D');
  const second = source.indexOf('8. C');
  expect(target.normalization.steps[1]!.changes).toEqual([
    { sourceStart: 5, sourceEnd: 5, targetStart: 5, targetEnd: 6 },
    { sourceStart: second, sourceEnd: second, targetStart: second + 1, targetEnd: second + 2 },
  ]);
  expect(targetSourceRangeAt(target, second + 2, target.markdown.length)).toEqual(
    unchanged(second, source.length)
  );
});

test('HTML break and list separator passes compose their real intermediate coordinates', () => {
  const target = parsed('A<BR>B\n6. First\n7. Second');
  expect(target.markdown).toBe('A<br />B\n\n6. First\n7. Second');
  expect(target.normalization.steps).toEqual([
    {
      sourceLength: 25,
      targetLength: 27,
      changes: [{ sourceStart: 1, sourceEnd: 5, targetStart: 1, targetEnd: 7 }],
    },
    {
      sourceLength: 27,
      targetLength: 28,
      changes: [{ sourceStart: 9, sourceEnd: 9, targetStart: 9, targetEnd: 10 }],
    },
  ]);
  expect(targetSourceRangeAt(target, 10, 28)).toEqual(unchanged(7, 25));
  expect(targetSourceRangeAt(target, 9, 10)).toEqual({
    start: 7,
    end: 7,
    exact: false,
    generated: true,
  });
  expect(targetSourceRangeAt(target, 1, 10)).toEqual({
    start: 1,
    end: 7,
    exact: false,
    generated: true,
  });
  expect(targetSourceOffsetAt(target, 2)).toBeNull();
});

test('code examples, escaped tags, fences, inline math, and non-interrupting prose keep existing normalization behavior', () => {
  const sources = [
    '`<br>`',
    '```html\n<br>\n6. One\n7. Two\n```',
    '~~~\n<br>\n6. One\n7. Two\n~~~',
    '    <br>',
    '\\<br>',
    '$$\n6. One\n7. Two\n$$',
    'Label:\n6. Only',
    'začala\n2. světová válka',
    '1. First\n2. Second',
    'Label:\n1. First\n2. Second',
    '<activity>\n6. One\n7. Two\n</activity>',
  ];
  for (const source of sources) {
    const result = normalizeTargetSource(source);
    expect(result.markdown).toBe(source);
    expect(result.normalization.steps.map((step) => step.changes)).toEqual([[], []]);
  }
  // The trace follows the existing string normalizer even in an expression-like text span.
  const expression = parsed('Before {"<br>"} after');
  expect(expression.markdown).toBe('Before {"<br />"} after');
  expect(expression.normalization.steps[0]!.changes).toHaveLength(1);
  const inlineMath = parsed('$A<br>B$');
  expect(inlineMath.markdown).toBe('$A<br />B$');
  expect(inlineMath.normalization.steps[0]!.changes).toHaveLength(1);
});

test('CRLF and Unicode separators retain exact source offsets while inserted LF remains generated', () => {
  const source = 'Label:\r\n6. First\r\n7. Second',
    target = parsed(source);
  expect(target.markdown).toBe('Label:\r\n\n6. First\r\n7. Second');
  expect(targetSourceRangeAt(target, 8, 9)).toEqual({
    start: 8,
    end: 8,
    exact: false,
    generated: true,
  });
  expect(targetSourceRangeAt(target, 9, target.markdown.length)).toEqual(
    unchanged(8, source.length)
  );
  const unicode = parsed('<div><br\u00a0>B</div>');
  expect(unicode.markdown).toBe('<div><br />B</div>');
  expect(targetSourceRangeAt(unicode, 11, 12)).toEqual(unchanged(10, 11));
  expect(normalizeTargetSource('A<br\u00a0>B').normalization.steps[0]!.changes).toEqual([]);
  expect(parseTargetMarkdown('A<br\u00a0>B').success).toBe(false);
});

test('surrogate interiors refuse in normalized and original strings while combining boundaries remain coordinate-only', () => {
  const target = parsed('😀<br>é🧑');
  expect(target.markdown).toBe('😀<br />é🧑');
  expect(targetSourceOffsetAt(target, 1)).toBeNull();
  expect(targetSourceRangeAt(target, 0, 2)).toEqual(unchanged(0, 2));
  const last = target.markdown.lastIndexOf('🧑');
  expect(targetSourceOffsetAt(target, last + 1)).toBeNull();
  expect(targetSourceRangeAt(target, last, last + 2)).toEqual(unchanged(8, 10));
  const combining = target.markdown.indexOf('́');
  expect(targetSourceOffsetAt(target, combining)).toBe(7);
  expect(targetSourceRangeAt(target, combining - 1, combining + 1)).toEqual(unchanged(6, 8));
});

test('invalid, inverted, fractional, nonfinite, and out-of-bounds coordinates refuse without partial ranges', () => {
  const target = parsed('A<br>B');
  for (const offset of [-1, 0.5, NaN, Infinity, -Infinity, 9, Number.MAX_SAFE_INTEGER]) {
    expect(targetSourceOffsetAt(target, offset)).toBeNull();
    expect(targetSourceRangeAt(target, 0, offset)).toBeNull();
    expect(targetSourceRangeAt(target, offset, 8)).toBeNull();
  }
  expect(targetSourceRangeAt(target, 8, 7)).toBeNull();
  expect(targetSourceRangeAt(target, 8, 8)).toEqual(unchanged(6, 6));
});

test('every independent parser reference and C01 stream retains output/AST behavior and exact unchanged source ranges', () => {
  expect(reference.cases).toHaveLength(24);
  for (const entry of reference.cases) {
    const target = parsed(entry.input);
    expect(target.markdown).toBe(entry.markdown);
    expect<unknown>(target.normalizedAst).toEqual(entry.normalizedAst);
    expect(targetSourceRangeAt(target, 0, target.markdown.length)).toMatchObject({
      start: 0,
      end: entry.input.length,
    });
    for (const node of nodes(target.ast)) {
      const position = node.position as
        | { start: { offset: number }; end: { offset: number } }
        | undefined;
      if (!position) continue;
      const start = position.start.offset,
        end = position.end.offset,
        range = targetSourceRangeAt(target, start, end);
      if (range?.exact)
        expect(entry.input.slice(range.start, range.end)).toBe(target.markdown.slice(start, end));
    }
  }
  for (const entry of pairs.cases)
    for (const source of [entry.beforeMarkdown, entry.afterMarkdown]) {
      const target = parsed(source);
      expect(targetSourceRangeAt(target, 0, target.markdown.length)).toEqual(
        unchanged(0, source.length)
      );
      const logical = normalizeLogicalBlocks(target);
      expect(logical.success).toBe(true);
      if (logical.success) expect(logical.document.ast).toEqual(target.ast);
    }
});

test('trace outputs are detached, deeply frozen, deterministic JSON without native editor references', () => {
  const first = parsed('A<br>B\n6. First\n7. Second'),
    second = parsed(first.sourceMarkdown);
  expect(first.normalization).toEqual(second.normalization);
  expect(first.normalization).not.toBe(second.normalization);
  expect(frozen(first.normalization)).toBe(true);
  expect(JSON.parse(JSON.stringify(first.normalization))).toEqual(first.normalization);
  expect(Object.hasOwn(first.normalization, 'editor')).toBe(false);
  expect(Object.hasOwn(first.normalization, 'paragraphId')).toBe(false);
  expect(Object.isFrozen(targetSourceRangeAt(first, 10, first.markdown.length))).toBe(true);
});

test('real native table text after a normalized break resolves to its original paragraph and UTF-16 offsets', async () => {
  const loaded = readOoxmlPackage(nativeTableBytes());
  if (!loaded.ok) throw new Error(loaded.reason);
  const projection = await projectSourcePackage(loaded.package),
    target = parsed(projection.text),
    part = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
  expect(projection.text).toContain('**A**<br>**After**');
  expect(target.markdown).toContain('**A**<br />**After**');
  const start = target.markdown.indexOf('After'),
    range = targetSourceRangeAt(target, start, start + 5)!;
  expect(range).toEqual(
    unchanged(projection.text.indexOf('After'), projection.text.indexOf('After') + 5)
  );
  const native = markdownRangeAt(projection, range.start, range.end);
  expect(native?.start.paragraphId).toBe(native?.end.paragraphId);
  expect(native?.start.offset).toBe(2);
  expect(native?.end.offset).toBe(7);
  expect(paragraphTextOf(part, native!.start.paragraphId)).toBe('A\nAfter');
  const indexed = indexTargetSource({ revision: sourceRevision, pkg: loaded.package, projection });
  expect(indexed.success).toBe(true);
  if (indexed.success)
    expect(
      indexed.index.owners.find((owner) => owner.id === native!.start.paragraphId)?.cell
    ).not.toBeNull();
  // Exact string coordinates alone do not grant ownership of the generated table delimiter.
  const delimiter = targetSourceOffsetAt(target, 0)!;
  expect(delimiter).toBe(0);
  expect(markdownPointAt(projection, delimiter)).toBeNull();
});

test('nested table text keeps its own native paragraph after a generated break and prior string rewrites', async () => {
  const input = await targetSourceFixture(),
    target = parsed(input.projection.text),
    start = target.markdown.indexOf('Nested B'),
    range = targetSourceRangeAt(target, start, start + 8)!;
  const native = markdownRangeAt(input.projection, range.start, range.end)!;
  expect(native.start.paragraphId).toBe(native.end.paragraphId);
  expect([native.start.offset, native.end.offset]).toEqual([0, 8]);
  expect(
    paragraphTextOf(input.pkg.parts.get(input.pkg.mainDocumentPart)!, native.start.paragraphId)
  ).toBe('Nested B');
  const indexed = indexTargetSource(input);
  if (!indexed.success) throw new Error(indexed.error);
  const owner = indexed.index.owners.find((owner) => owner.id === native.start.paragraphId)!;
  expect(owner.cell).not.toBeNull();
  expect(owner.cell?.tableId).not.toBe(
    indexed.index.owners.find((candidate) => candidate.kind === 'table' && !candidate.cell)?.id
  );
  expect(range.exact).toBe(true);
});

test('parsing, source mapping, and invalid targets preserve live native bytes, imported XML, selection, and undo/redo history', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: nativeTableBytes(),
  });
  try {
    const surface = editor.surface!,
      session = surface.session;
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: session.paragraphIds()[0]!, offset: 0, text: 'User ' },
    ]);
    const bytes = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision(),
      xml = canonicalOoxmlFingerprint(session.part()),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    const projection = await projectSourcePackage(session.currentPackage()),
      target = parsed(projection.text),
      start = target.markdown.indexOf('After'),
      range = targetSourceRangeAt(target, start, start + 5)!;
    expect(markdownRangeAt(projection, range.start, range.end)?.start.offset).toBe(7);
    expect(targetSourceOffsetAt(target, -1)).toBeNull();
    expect(parseTargetMarkdown('<activity>').success).toBe(false);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(revision);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(xml);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(session.canRedo()).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
  } finally {
    editor.destroy();
  }
});
