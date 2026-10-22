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
import { docxFixture } from '../../test/support/document';
import mdxFixture from '../__tests__/fixtures/mdx-tags.json';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload, JsonObject } from '../blocks/payload-schema';
import { findAndReplace } from '../tools/find-replace';
import { projectMdx, MDX_INLINE } from './mdx';
import { projectTables } from './tables';
import {
  resolvePatchSpan,
  type MarkdownPatchSpan,
  type PatchResolutionContext,
} from './resolve-patch';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
const p = (text: string) => `<w:p>${run(text)}</w:p>`;
const control = (id: string, body: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const payload = (id: string, type: string, attributes: JsonObject = {}): BlockPayload => ({
  id,
  type,
  attributes,
  version: 1,
});

async function imported(body: string, payloads: BlockPayload[] = []) {
  const parsed = readOoxmlPackage(docxFixture(body));
  if (!parsed.ok) throw new Error(parsed.reason);
  let pkg = parsed.package;
  for (const value of payloads) {
    const written = withPayloadRecord(pkg, value, 'Word fallback');
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  const surface = editor.surface!;
  const session = surface.session;
  const opened = openDocumentForExport(session);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout();
    const revision = { documentId: 'lesson', epoch: 'load-1', revision: session.packageRevision() };
    const projected = projectMdx(session.currentPackage(), layout, revision);
    const tables = projectTables(layout, revision);
    if (!projected.success) throw new Error(projected.error);
    if (!tables.success) throw new Error(tables.error);
    const projection = projected.projection;
    const fingerprints = () =>
      [...session.currentPackage().parts].map(([name, part]) => [
        name,
        canonicalOoxmlFingerprint(part),
      ]);
    const before = fingerprints();
    const selection = surface.state().selection;
    const bytes = writeOoxmlPackage(session.currentPackage());
    const paragraphs = session
      .paragraphIds()
      .map((id) => ({ id, text: paragraphTextOf(session.part(), id) }));
    const context: PatchResolutionContext = { revision, tableCells: tables.projection.cells };
    return {
      projection,
      paragraphs,
      context,
      span(text: string): MarkdownPatchSpan {
        const start = projection.text.indexOf(text);
        if (start < 0) throw new Error(`Missing projected text: ${text}`);
        return { start, end: start + text.length };
      },
      resolve(span: MarkdownPatchSpan, overrides: Partial<PatchResolutionContext> = {}) {
        return resolvePatchSpan(projection, span, { ...context, ...overrides });
      },
      dispose() {
        try {
          expect(fingerprints()).toEqual(before);
          expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
          expect(session.packageRevision()).toBe(revision.revision);
          expect(surface.state().selection).toEqual(selection);
          expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
        } finally {
          opened.session.dispose();
          editor.destroy();
        }
      },
    };
  } catch (error) {
    opened.session.dispose();
    editor.destroy();
    throw error;
  }
}

test('a unique copied match resolves exact native UTF-16 offsets across marked runs', async () => {
  const value = await imported(
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>' + run(' and 😀 plain') + '</w:p>'
  );
  try {
    const match = value.span('plain');
    expect(findAndReplace(value.projection.text, 'plain', 'edited')).toMatchObject({
      success: true,
    });
    const result = value.resolve(match);
    expect(result).toMatchObject({
      success: true,
      target: {
        kind: 'text',
        start: { paragraphId: value.paragraphs[0]!.id, offset: 'Bold and 😀 '.length },
        end: { paragraphId: value.paragraphs[0]!.id, offset: 'Bold and 😀 plain'.length },
        paragraphIds: [value.paragraphs[0]!.id],
        cellId: null,
      },
    });
    if (!result.success) throw new Error(result.error);
    expect(Object.isFrozen(result.target)).toBe(true);
    expect(JSON.parse(JSON.stringify(result.target))).toEqual(result.target);
    expect(Object.keys(result)).toEqual(['success', 'target']);
  } finally {
    value.dispose();
  }
});

test('adjacent native paragraph ranges retain every canonical identity', async () => {
  const value = await imported(p('First sentence') + p('Middle sentence') + p('Last sentence'));
  try {
    const start = value.span('sentence').start;
    const end = value.span('Last').end;
    expect(value.resolve({ start, end })).toMatchObject({
      success: true,
      target: {
        kind: 'text',
        start: { paragraphId: value.paragraphs[0]!.id, offset: 6 },
        end: { paragraphId: value.paragraphs[2]!.id, offset: 4 },
        paragraphIds: value.paragraphs.map((paragraph) => paragraph.id),
      },
    });
  } finally {
    value.dispose();
  }
});

test('complete headings, empty paragraphs, and tables resolve as syntax-aware constructs', async () => {
  const value = await imported(
    '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>' +
      run('Heading') +
      '</w:p>' +
      p('') +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>' +
      p('Cell') +
      '</w:tc></w:tr></w:tbl>'
  );
  try {
    const heading = value.projection.sourceMap.paragraphs.find(
      (paragraph) => paragraph.paragraphId === value.paragraphs[0]!.id
    )!;
    expect(
      value.resolve({ start: heading.markdownStart!, end: heading.markdownEnd! })
    ).toMatchObject({
      success: true,
      target: { kind: 'paragraph', paragraphId: value.paragraphs[0]!.id },
    });
    expect(value.resolve(value.span('##'))).toMatchObject({
      success: false,
      code: 'ambiguous-target',
    });
    const empty = value.projection.sourceMap.paragraphs.find(
      (paragraph) => paragraph.paragraphId === value.paragraphs[1]!.id
    )!;
    expect(value.resolve({ start: empty.markdownStart!, end: empty.markdownEnd! })).toMatchObject({
      success: true,
      target: { kind: 'paragraph', paragraphId: value.paragraphs[1]!.id },
    });
    const table = value.projection.sourceMap.tables[0]!;
    expect(value.resolve({ start: table.markdownStart, end: table.markdownEnd })).toMatchObject({
      success: true,
      target: { kind: 'table', tableId: table.tableId },
    });
  } finally {
    value.dispose();
  }
});

test('domain attributes target persistent properties while complete containers keep block identity', async () => {
  const value = await imported(control('activity', p('Question text')), [
    payload('activity', 'lesson_activity', { name: 'Warmup', duration: '15' }),
  ]);
  try {
    const property = value.projection.attributes.find(
      (attribute) => attribute.property === 'duration'
    )!;
    expect(value.resolve({ start: property.valueStart!, end: property.valueEnd! })).toMatchObject({
      success: true,
      target: {
        kind: 'property',
        blockId: 'activity',
        controlId: value.projection.blocks[0]!.controlId,
        property: 'duration',
        path: ['attributes', 'duration'],
      },
    });
    expect(value.resolve(value.span('Question'))).toMatchObject({
      success: true,
      target: { kind: 'text', start: { offset: 0 }, end: { offset: 8 } },
    });
    const block = value.projection.blocks[0]!;
    expect(value.resolve({ start: block.start, end: block.end })).toMatchObject({
      success: true,
      target: { kind: 'block', blockId: 'activity', type: 'lesson_activity' },
    });
    expect(value.resolve({ start: property.start, end: value.span('Question').end })).toMatchObject(
      { success: false, code: 'unsupported-operation' }
    );
  } finally {
    value.dispose();
  }
});

test('complete opaque trees resolve to their stored AST property without touching Word fallback text', async () => {
  const reference = mdxFixture.cases.find((entry) => entry.id === 'competence-inline')!;
  const value = await imported(
    '<w:p>' + run('Before ') + control('opaque', run('Word fallback')) + run(' after') + '</w:p>',
    [{ ...payload('opaque', MDX_INLINE), mdast: reference.ast as unknown as JsonObject }]
  );
  try {
    const block = value.projection.blocks[0]!;
    expect(value.resolve({ start: block.start, end: block.end })).toMatchObject({
      success: true,
      target: { kind: 'property', blockId: 'opaque', property: 'mdast', path: ['mdast'] },
    });
    expect(value.resolve(value.span('Stored'))).toMatchObject({
      success: true,
      target: { kind: 'property', property: 'mdast' },
    });
    expect(
      value.resolve({ start: value.span('Before').start, end: value.span('Stored').end })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
  } finally {
    value.dispose();
  }
});

test('same-cell paragraphs use P03 identities and cross-cell spans require a table plan', async () => {
  const value = await imported(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc>' +
      p('First cell start') +
      p('First cell end') +
      '</w:tc><w:tc>' +
      p('Second cell') +
      '</w:tc></w:tr></w:tbl>'
  );
  try {
    const span = {
      start: value.span('cell start').start,
      end: value.span('First cell end').start + 5,
    };
    expect(value.resolve(span)).toMatchObject({
      success: true,
      target: {
        kind: 'text',
        cellId: value.context.tableCells![0]!.id,
        paragraphIds: value.paragraphs.slice(0, 2).map((paragraph) => paragraph.id),
      },
    });
    expect(value.resolve(span, { tableCells: [] })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(value.resolve({ start: span.start, end: value.span('Second cell').end })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
  } finally {
    value.dispose();
  }
});

test('protected interiors fail before returning targets, including an internal field boundary', async () => {
  const value = await imported(p('Before field after'));
  try {
    const field = value.span('field');
    const protectedContext = { protected: [{ ...field, nodeId: 'field-1', reason: 'Word field' }] };
    expect(
      value.resolve({ start: 0, end: value.projection.text.length }, protectedContext)
    ).toEqual({
      success: false,
      code: 'unsupported-operation',
      error: 'The matched Markdown touches protected content (field-1): Word field',
    });
    expect(
      value.resolve({ start: field.start + 1, end: field.start + 1 }, protectedContext)
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(value.resolve({ start: 0, end: field.start }, protectedContext)).toMatchObject({
      success: true,
      target: { kind: 'text' },
    });
    expect(
      value.resolve(
        { start: field.start, end: field.start },
        {
          protected: [{ start: field.start, end: field.start, reason: 'zero-width native marker' }],
        }
      )
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(
      value.resolve(value.span('Before'), {
        protected: [{ start: NaN, end: field.end, reason: 'invalid field map' }],
      })
    ).toMatchObject({ success: false, code: 'invalid-document' });
  } finally {
    value.dispose();
  }
});

test('stale document, load epoch, and package revision fail without a target', async () => {
  const value = await imported(p('Stable native text'));
  try {
    for (const revision of [
      { ...value.context.revision, documentId: 'other' },
      { ...value.context.revision, epoch: 'load-2' },
      { ...value.context.revision, revision: value.context.revision.revision + 1 },
    ]) {
      const result = value.resolve(value.span('native'), { revision });
      expect(result).toMatchObject({
        success: false,
        code: 'stale-revision',
        expectedRevision: value.context.revision,
        actualRevision: revision,
      });
      expect(Object.hasOwn(result, 'target')).toBe(false);
    }
  } finally {
    value.dispose();
  }
});

test('escape interiors, surrogate splits, and invalid spans fail atomically', async () => {
  const value = await imported(p('A * star 😀 tail'));
  try {
    const star = value.projection.text.indexOf('\\*');
    expect(value.resolve({ start: star + 1, end: star + 2 })).toMatchObject({
      success: false,
      code: 'ambiguous-target',
    });
    const emoji = value.span('😀');
    expect(value.resolve({ start: emoji.start + 1, end: emoji.end })).toMatchObject({
      success: false,
      code: 'ambiguous-target',
    });
    for (const span of [
      { start: -1, end: 0 },
      { start: 2, end: 1 },
      { start: 0.5, end: 1 },
      { start: 0, end: Infinity },
      { start: 0, end: value.projection.text.length + 1 },
    ])
      expect(value.resolve(span)).toMatchObject({ success: false, code: 'invalid-input' });
  } finally {
    value.dispose();
  }
});

test('overlapping property maps and duplicate construct identities refuse ambiguity', async () => {
  const value = await imported(control('activity', p('Question')), [
    payload('activity', 'lesson_activity', { duration: '15' }),
  ]);
  try {
    const property = value.projection.attributes[0]!;
    const duplicate = {
      ...value.projection,
      attributes: [...value.projection.attributes, { ...property, blockId: 'different' }],
    };
    expect(
      resolvePatchSpan(duplicate, { start: property.start, end: property.end }, value.context)
    ).toMatchObject({ success: false, code: 'ambiguous-target' });
    const paragraph = value.projection.sourceMap.paragraphs[0]!;
    const repeated = {
      ...value.projection,
      sourceMap: {
        ...value.projection.sourceMap,
        paragraphs: [paragraph, { ...paragraph, paragraphId: 'different' }],
      },
    };
    expect(
      resolvePatchSpan(
        repeated,
        { start: paragraph.markdownStart!, end: paragraph.markdownEnd! },
        value.context
      )
    ).toMatchObject({ success: false, code: 'ambiguous-target' });
  } finally {
    value.dispose();
  }
});
