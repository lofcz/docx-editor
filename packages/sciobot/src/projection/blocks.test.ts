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
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload } from '../blocks/payload-schema';
import { DOMAIN_BLOCK_TAGS, GENERATED_LABEL_TAG, type BlockTagDefinition } from './block-tags';
import { blockAttributeAt, projectBlocks } from './blocks';
import { markdownPointAt, markdownRangeAt } from './markdown';

const p = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
const control = (id: string, content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const label = (text: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${GENERATED_LABEL_TAG}"/></w:sdtPr><w:sdtContent>${p(text)}</w:sdtContent></w:sdt>`;
const payload = (
  id: string,
  type: string,
  attributes: BlockPayload['attributes'] = {}
): BlockPayload => ({ id, type, version: 1, attributes });

async function project(
  body: string,
  payloads: BlockPayload[],
  definitions?: ReadonlyMap<string, BlockTagDefinition>,
  mode: 'original' | 'proposed' = 'proposed'
) {
  const parsed = readOoxmlPackage(docxFixture(body));
  if (!parsed.ok) throw new Error(parsed.reason);
  let pkg = parsed.package;
  for (const value of payloads) {
    const written = withPayloadRecord(pkg, value, 'Readable fallback');
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  const session = editor.surface!.session;
  const opened = openDocumentForExport(session, { displayMode: mode });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const before = [...session.currentPackage().parts].map(([name, part]) => [
      name,
      canonicalOoxmlFingerprint(part),
    ]);
    const revision = { documentId: 'lesson', epoch: 'load-1', revision: session.packageRevision() };
    const paragraphs = session
      .paragraphIds()
      .map((id) => ({ id, text: paragraphTextOf(session.part(), id) }));
    const result = projectBlocks(
      session.currentPackage(),
      await opened.session.layout(),
      revision,
      definitions
    );
    if (!result.success) throw new Error(result.error);
    expect(
      [...session.currentPackage().parts].map(([name, part]) => [
        name,
        canonicalOoxmlFingerprint(part),
      ])
    ).toEqual(before);
    expect(session.packageRevision()).toBe(revision.revision);
    expect(session.canUndo()).toBe(false);
    return { projection: result.projection, paragraphs };
  } finally {
    opened.session.dispose();
    editor.destroy();
  }
}

test('nested lesson attributes target payload properties while sentences keep their native identities', async () => {
  const { projection, paragraphs } = await project(
    control('title', p('Lesson')) +
      control(
        'phase',
        label('Phase heading') +
          control(
            'activity',
            label('Activity heading') +
              p('First question.') +
              `<w:p>${control('source', '<w:r><w:t>Citation</w:t></w:r>')}</w:p>` +
              p('Last question.')
          ) +
          p('After activity')
      ) +
      p('Tail'),
    [
      payload('title', 'lesson_title', { grade: '7' }),
      payload('phase', 'lesson_phase', { phaseType: 'evocation' }),
      payload('activity', 'lesson_activity', { name: 'Warmup', duration: '10' }),
      payload('source', 'source', {
        trust: 'verified',
        url: 'https://example.com',
        title: 'Source',
        sourceType: 'web',
      }),
    ]
  );
  expect(projection.text).toBe(
    '<lesson_title grade="7">\n\nLesson\n\n</lesson_title>\n\n<phase type="evocation">\n\n<activity name="Warmup" duration="10">\n\nFirst question\\.\n\n<source type="verified" url="https://example.com" title="Source" />\n\nLast question\\.\n\n</activity>\n\nAfter activity\n\n</phase>\n\nTail'
  );
  const duration = projection.text.indexOf('10');
  expect(blockAttributeAt(projection, duration)).toMatchObject({
    blockId: 'activity',
    property: 'duration',
    name: 'duration',
    valueStart: duration,
    valueEnd: duration + 2,
  });
  expect(markdownPointAt(projection, duration)).toBeNull();
  const question = projection.text.indexOf('First question');
  expect(
    markdownRangeAt(projection, question, question + 'First question\\.'.length)
  ).toMatchObject({
    start: {
      paragraphId: paragraphs.find((value) => value.text === 'First question.')!.id,
      offset: 0,
    },
    end: { offset: 15 },
  });
  expect(projection.blocks.map((value) => value.id)).toEqual([
    'title',
    'phase',
    'activity',
    'source',
  ]);
  expect(projection.fallbackParagraphIds).toHaveLength(2);
  expect(
    projection.sourceMap.paragraphs.some((value) =>
      projection.fallbackParagraphIds.includes(value.paragraphId)
    )
  ).toBe(false);
  expect(projection.text).not.toContain('heading');
  expect(JSON.parse(JSON.stringify(projection))).toEqual(projection);
});

test('native tables and unmarked first paragraphs remain inside their registered container', async () => {
  const { projection, paragraphs } = await project(
    control(
      'activity',
      p('Keep first') +
        '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>' +
        p('A|B') +
        '</w:tc><w:tc>' +
        p('😀') +
        '</w:tc></w:tr></w:tbl>' +
        p('Last')
    ) + p('Outside'),
    [payload('activity', 'lesson_activity', { name: 'A "quote" & <name>', duration: 10 })]
  );
  expect(projection.text).toBe(
    '<activity name="A &quot;quote&quot; &amp; &lt;name&gt;" duration="10">\n\nKeep first\n\n| A\\|B | 😀 |\n| --- | --- |\n\nLast\n\n</activity>\n\nOutside'
  );
  expect(projection.fallbackParagraphIds).toEqual([]);
  const table = projection.sourceMap.tables[0]!;
  expect(projection.text.slice(table.markdownStart, table.markdownEnd)).toBe(
    '| A\\|B | 😀 |\n| --- | --- |'
  );
  expect(markdownPointAt(projection, projection.text.indexOf('Keep first'))).toMatchObject({
    paragraphId: paragraphs[0]!.id,
    offset: 0,
  });
  const emoji = projection.text.indexOf('😀');
  expect(markdownPointAt(projection, emoji)).toMatchObject({ offset: 0 });
  expect(markdownPointAt(projection, emoji + 1)).toBeNull();
});

test('inline source replacement preserves native text offsets on both sides', async () => {
  const { projection, paragraphs } = await project(
    '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
      control('source', '<w:r><w:t>Label*</w:t></w:r>') +
      '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>',
    [
      payload('source', 'source', {
        trust: 'verified',
        url: 'https://example.com?a=1&b=2',
        title: 'Title',
        bookId: 42,
      }),
    ]
  );
  expect(projection.text).toBe(
    'Before <source type="verified" url="https://example.com?a=1&amp;b=2" title="Title" bookId="42" /> after'
  );
  expect(markdownPointAt(projection, 2)).toMatchObject({
    paragraphId: paragraphs[0]!.id,
    offset: 2,
  });
  expect(markdownPointAt(projection, projection.text.lastIndexOf('after'))).toMatchObject({
    paragraphId: paragraphs[0]!.id,
    offset: 14,
  });
  expect(markdownPointAt(projection, projection.text.indexOf('https'))).toBeNull();
  expect(blockAttributeAt(projection, projection.text.indexOf('https'))?.property).toBe('url');
  const paragraph = projection.sourceMap.paragraphs[0]!;
  expect(projection.text.slice(paragraph.markdownStart, paragraph.markdownEnd)).toBe(
    projection.text
  );
});

test('empty native containers retain their identity and a native empty paragraph construct', async () => {
  const { projection } = await project(control('recap', '<w:p/>'), [
    payload('recap', 'preparation_recap'),
  ]);
  expect(projection.text).toBe('<preparation_recap>\n\n\n\n</preparation_recap>');
  expect(projection.blocks[0]?.id).toBe('recap');
  expect(projection.sourceMap.paragraphs).toHaveLength(1);
  expect(projection.sourceMap.paragraphs[0]?.markdownStart).toBe(
    projection.blocks[0]?.contentStart
  );
});

test('an application can register arbitrary container tags while unknown types keep native fallback', async () => {
  const body = control('custom', p('Native child'));
  const values = [payload('custom', 'plugin.Widget', { category: 'experiment' })];
  const unknown = await project(body, values);
  expect(unknown.projection.text).toBe('Native child');
  expect(unknown.projection.blocks).toEqual([]);
  const definitions = new Map(DOMAIN_BLOCK_TAGS);
  definitions.set('plugin.Widget', {
    tag: 'custom-widget',
    kind: 'container',
    attributes: (value) => [
      { name: 'category', property: 'category', value: value.attributes.category! },
    ],
  });
  const known = await project(body, values, definitions);
  expect(known.projection.text).toBe(
    '<custom-widget category="experiment">\n\nNative child\n\n</custom-widget>'
  );
  expect(
    blockAttributeAt(known.projection, known.projection.text.indexOf('experiment'))?.blockId
  ).toBe('custom');
});

test('native review projections retain original source coordinates inside domain tags', async () => {
  const body = control(
    'phase',
    '<w:p><w:del w:id="1" w:author="A"><w:r><w:delText>Old</w:delText></w:r></w:del>' +
      '<w:ins w:id="2" w:author="A"><w:r><w:t>New</w:t></w:r></w:ins></w:p>'
  );
  const values = [payload('phase', 'lesson_phase', { phaseType: 'reflection' })];
  const original = (await project(body, values, undefined, 'original')).projection;
  const working = (await project(body, values)).projection;
  expect(original.text).toBe('<phase type="reflection">\n\nOld\n\n</phase>');
  expect(working.text).toBe('<phase type="reflection">\n\nNew\n\n</phase>');
  expect(markdownPointAt(original, original.text.indexOf('Old'))?.offset).toBe(0);
  expect(markdownPointAt(working, working.text.indexOf('New'))?.offset).toBe(3);
});

test('empty paragraphs beside registered containers keep exact construct positions', async () => {
  const { projection } = await project(control('first', '<w:p/>') + control('second', p('Text')), [
    payload('first', 'preparation_recap'),
    payload('second', 'preparation_recap'),
  ]);
  const first = projection.sourceMap.paragraphs[0]!;
  expect(first.markdownStart).toBe(projection.blocks[0]!.contentStart);
  expect(first.markdownEnd).toBe(first.markdownStart);
  expect(
    projection.text.slice(
      projection.sourceMap.paragraphs[1]!.markdownStart,
      projection.sourceMap.paragraphs[1]!.markdownEnd
    )
  ).toBe('Text');
});

test('domain property remapping matches title, recap, visual, and excerpt rules', async () => {
  const { projection } = await project(
    control('title', p('Title')) +
      control('recap', p('Summary')) +
      control(
        'visual',
        label('Picture') + '<w:tbl><w:tr><w:tc>' + p('Word fallback') + '</w:tc></w:tr></w:tbl>'
      ) +
      control('excerpt', p('Quoted text')),
    [
      payload('title', 'lesson_title', {
        grade: '',
        organizationalMinutes: 3,
        lessonMinutes: 45,
        extra: true,
        _id: 'internal',
      }),
      payload('recap', 'recap_phase', { phaseType: 'reflection', duration: 99, extra: 'kept' }),
      payload('visual', 'lesson_visual', {
        kind: 'summary',
        phase: 'evocation',
        src: 'media/chart.png',
        alt: 'Chart',
        prompt: 'Draw',
        extra: 'kept',
      }),
      payload('excerpt', 'source_excerpt', {
        trust: 'generated',
        title: 'Book',
        sourceType: 'web',
        bookId: 42,
        extra: 'kept',
      }),
    ]
  );
  expect(projection.text).toContain(
    '<lesson_title organizational_minutes="3" lesson_minutes="45" extra>'
  );
  expect(projection.text).not.toContain('_id');
  expect(projection.text).toContain('<recap_phase type="reflection" extra="kept">');
  expect(projection.text).not.toContain('duration="99"');
  expect(projection.text).toContain(
    '<lesson_visual kind="summary" src="media/chart.png" alt="Chart" prompt="Draw" extra="kept" />'
  );
  expect(projection.text).not.toContain('Word fallback');
  expect(projection.sourceMap.tables).toEqual([]);
  expect(
    projection.sourceMap.paragraphs.some((paragraph) =>
      projection.fallbackParagraphIds.includes(paragraph.paragraphId)
    )
  ).toBe(false);
  expect(projection.text).toContain(
    '<source type="generated" title="Book" bookId="42" extra="kept">\n\nQuoted text\n\n</source>'
  );
  expect(
    blockAttributeAt(projection, projection.text.indexOf('organizational_minutes'))?.property
  ).toBe('organizationalMinutes');
  expect(blockAttributeAt(projection, projection.text.indexOf('extra>'))?.valueStart).toBeNull();
});
