import { expect, test } from 'bun:test';
import {
  ATOMIC_CONTENT_NAMESPACE,
  withAtomicContentMarker,
  isAtomicContentControl,
} from '../package/atomic-content-control';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  contentControlsIn,
  readOoxmlPart,
  replaceNode,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
} from '../index';

function fixture(extra = '') {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:sdt ${extra}><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!read.ok) throw Error(read.reason);
  return read.part;
}

test('atomic fallback markers survive markup-compatible save and reopen and refuse generic text edits', () => {
  const part = fixture(
    'xmlns:docxAtomic="urn:other" xmlns:other="urn:existing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="other" mc:PreserveAttributes="other:flag" other:flag="retained"'
  );
  const control = contentControlsIn(part.root)[0]!.node;
  const marked = withAtomicContentMarker(control);
  const next = replaceNode(part, control.id, marked);
  if (!next.ok) throw Error('The marker was refused.');
  const xml = serializeOoxmlPart(next.part);
  expect(xml).toContain('Ignorable="docxAtomic2 other"');
  expect(xml).toContain('PreserveAttributes="docxAtomic2:atomicContent other:flag"');
  const read = readOoxmlPart(xml, { name: part.name, contentType: part.contentType });
  if (!read.ok) throw Error(read.reason);
  const restored = contentControlsIn(read.part.root)[0]!.node;
  expect(isAtomicContentControl(restored)).toBe(true);
  expect(restored.namespaceBindings).toContainEqual({
    prefix: 'docxAtomic',
    namespaceUri: 'urn:other',
  });
  expect(
    restored.attributes.find(
      (attribute) => attribute.namespaceUri === 'urn:existing' && attribute.localName === 'flag'
    )?.value
  ).toBe('retained');
  expect(contentControlPropertiesOf(restored).lock).toBe('contentLocked');
  const content = restored.children.find((node) => node.kind === 'contentControlContent')!;
  const paragraph = content.children.find((node) => node.kind === 'paragraph')!;
  expect(
    applyTreeOp(read.part, { op: 'insertText', paragraphId: paragraph.id, offset: 0, text: 'peer' })
      .ok
  ).toBe(false);
});

for (const [namespace, value] of [
  [ATOMIC_CONTENT_NAMESPACE, '2'],
  [ATOMIC_CONTENT_NAMESPACE, 'true'],
  ['urn:other', '1'],
])
  test(`unknown atomic marker stays ordinary: ${namespace}/${value}`, () => {
    const part = fixture(`xmlns:a="${namespace}" a:atomicContent="${value}"`);
    const control = contentControlsIn(part.root)[0]!.node;
    expect(isAtomicContentControl(control)).toBe(false);
    expect(contentControlPropertiesOf(control).lock).toBe('unlocked');
    const content = control.children.find((node) => node.kind === 'contentControlContent')!;
    const paragraph = content.children.find((node) => node.kind === 'paragraph')!;
    expect(
      applyTreeOp(part, { op: 'insertText', paragraphId: paragraph.id, offset: 0, text: 'peer' }).ok
    ).toBe(true);
  });
