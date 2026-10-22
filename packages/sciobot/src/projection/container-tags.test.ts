import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  escapeXml,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import reference from '../__tests__/fixtures/container-tags.json';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload } from '../blocks/payload-schema';
import { readWorkingView } from '../review/working-view';
import { DOMAIN_BLOCK_TAGS, attributeLiteral } from './block-tags';
import { blockAttributeAt } from './blocks';
import { markdownRangeAt } from './markdown';

const p = (text: string) => `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
const control = (id: string, content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const payload = (
  id: string,
  type: string,
  attributes: BlockPayload['attributes']
): BlockPayload => ({ id, type, attributes, version: 1 });

for (const item of reference.cases) {
  test(`container attributes match captured application output: ${item.id}`, () => {
    const bindings = DOMAIN_BLOCK_TAGS.get(item.type)!.attributes(
      payload(item.id, item.type, item.attributes as unknown as BlockPayload['attributes'])
    );
    const serialized = bindings
      .filter((binding) => binding.value !== null)
      .map((binding) => ({
        type: 'mdxJsxAttribute',
        name: binding.name,
        value:
          binding.value === true
            ? null
            : typeof binding.value === 'string'
              ? binding.value
              : JSON.stringify(binding.value),
      }));
    expect(serialized).toEqual(item.serializedAttributes);
  });
}

test('captured modern and legacy callout aliases project from variant back to type', () => {
  for (const item of reference.aliases) {
    const { type, children: _children, ...attributes } = item.node;
    const bindings = DOMAIN_BLOCK_TAGS.get(type)!.attributes(
      payload('alias', type, attributes as unknown as BlockPayload['attributes'])
    );
    expect(bindings.at(-1)).toEqual({
      name: 'type',
      property: 'variant',
      value: attributes.variant,
    });
    expect(bindings.some((binding) => binding.name === 'variant')).toBe(false);
  }
  expect(reference.aliases.map((item) => item.node.variant)).toEqual(['tip', 'warning', 'note']);
});

test('working container Markdown keeps native children, ordered inert properties, and reopened payloads', async () => {
  const body =
    control('tip', p('Tip 😀 text')) +
    control('story', p('Story child')) +
    control('solution', p('Answer child'));
  const parsed = readOoxmlPackage(docxFixture(body));
  if (!parsed.ok) throw new Error(parsed.reason);
  let pkg = parsed.package;
  for (const item of reference.cases.slice(0, 3)) {
    const written = withPayloadRecord(
      pkg,
      payload(item.id, item.type, item.attributes as unknown as BlockPayload['attributes']),
      'Fallback'
    );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  let bytes = writeOoxmlPackage(pkg);
  for (let pass = 0; pass < 2; pass++) {
    const editor = createDocxEditor({ container: document.createElement('div'), document: bytes });
    try {
      const session = editor.surface!.session;
      const before = writeOoxmlPackage(session.currentPackage());
      const token = {
        documentId: 'containers',
        epoch: `load-${pass}`,
        revision: session.packageRevision(),
      };
      const result = await readWorkingView(session, token);
      if (!result.success) throw new Error(result.error);
      const projection = result.view.projection;
      expect(projection.text).toBe(
        '<callout icon="💡" extra="unknown" type="tip">\n\nTip 😀 text\n\n</callout>\n\n<story zeta="first" audience="class" count="2" enabled hidden="false" _custom="kept" suggestion="inert">\n\nStory child\n\n</story>\n\n<solution answer="42" details="{&quot;nested&quot;:[&quot;data&quot;,2]}" constructor="inert" onClick="globalThis.X37 = true">\n\nAnswer child\n\n</solution>'
      );
      for (const text of ['Tip 😀 text', 'Story child', 'Answer child']) {
        const paragraphId = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === text)!;
        const start = projection.text.indexOf(text);
        expect(markdownRangeAt(projection, start, start + text.length)).toMatchObject({
          start: { paragraphId, offset: 0 },
          end: { paragraphId, offset: text.length },
        });
      }
      for (const item of reference.cases.slice(0, 3)) {
        for (const binding of DOMAIN_BLOCK_TAGS.get(item.type)!.attributes(
          payload(item.id, item.type, item.attributes as unknown as BlockPayload['attributes'])
        )) {
          if (binding.value === null) continue;
          const literal = attributeLiteral(binding.value);
          const attribute = binding.value === true ? binding.name : `${binding.name}="${literal}"`;
          const start = projection.text.indexOf(attribute);
          expect(start).toBeGreaterThanOrEqual(0);
          expect(blockAttributeAt(projection, start)).toMatchObject({
            blockId: item.id,
            name: binding.name,
            property: binding.property,
            path: ['attributes', binding.property],
          });
        }
      }
      expect(projection.blocks.map((block) => block.id)).toEqual(['tip', 'story', 'solution']);
      expect((globalThis as { X37?: unknown }).X37).toBeUndefined();
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
      expect(session.packageRevision()).toBe(token.revision);
      expect(session.canUndo()).toBe(false);
      bytes = writeOoxmlPackage(session.currentPackage());
    } finally {
      editor.destroy();
    }
  }
});

test('nested containers retain native boundaries and empty children', async () => {
  const parsed = readOoxmlPackage(
    docxFixture(control('story', control('solution', '<w:p/>') + control('callout', p('Nested'))))
  );
  if (!parsed.ok) throw new Error(parsed.reason);
  let pkg = parsed.package;
  for (const type of ['story', 'solution', 'callout']) {
    const written = withPayloadRecord(pkg, payload(type, type, {}), 'Fallback');
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  try {
    const session = editor.surface!.session;
    const result = await readWorkingView(session, {
      documentId: 'nested',
      epoch: '1',
      revision: session.packageRevision(),
    });
    if (!result.success) throw new Error(result.error);
    expect(result.view.projection.text).toBe(
      '<story>\n\n<solution>\n\n\n\n</solution>\n\n<callout>\n\nNested\n\n</callout>\n\n</story>'
    );
    expect(result.view.projection.sourceMap.paragraphs).toHaveLength(2);
  } finally {
    editor.destroy();
  }
});
