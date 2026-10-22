import { expect, test } from 'bun:test';
import { blankDocumentBytes } from '../blank-document.ts';
import {
  readOoxmlPackage,
  readOoxmlPart,
  replaceNode,
  withPart,
  type OoxmlPart,
} from '../../store/index.ts';
import { partOfNodeId } from '../surface-scope.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function part(name: string, localName: string): OoxmlPart {
  const content =
    '<w:p><w:ins w:id="1" w:author="ScioBot"><w:r><w:t>Pending</w:t></w:r></w:ins></w:p>';
  const loaded = readOoxmlPart(
    `<w:${localName} xmlns:w="${W}">${localName === 'document' ? `<w:body>${content}</w:body>` : content}</w:${localName}>`,
    { name, contentType: 'application/xml' }
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const holder = localName === 'document' ? loaded.part.root.children[0]! : loaded.part.root;
  if (holder.kind === 'textValue') throw new Error('Missing native container');
  const paragraph = holder.children[0]!;
  if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
  const wrapper = paragraph.children.find((node) => node.kind === 'revisionInsert')!;
  const changed = replaceNode(loaded.part, wrapper.id, { ...wrapper, id: 'lid:history:1' });
  if (!changed.ok) throw new Error('Cannot install local fixture ID');
  return changed.part;
}
test('local collaboration IDs resolve to body nodes without opening a story store', () => {
  const loaded = readOoxmlPackage(blankDocumentBytes());
  if (!loaded.ok) throw new Error(loaded.reason);
  const body = part(loaded.package.mainDocumentPart, 'document');
  const pkg = withPart(loaded.package, body);
  const session = { currentPackage: () => pkg, part: () => body };
  expect(partOfNodeId(session, 'lid:history:1') === body).toBe(true);
  expect(partOfNodeId(session, body.root.id) === body).toBe(true);
  expect(partOfNodeId(session, 'missing')).toBeNull();
  expect(partOfNodeId(session, undefined)).toBeNull();
});
test('local IDs resolve in related parts through package indexes without partFor', () => {
  const loaded = readOoxmlPackage(blankDocumentBytes());
  if (!loaded.ok) throw new Error(loaded.reason);
  const header = part('/word/header1.xml', 'hdr');
  const pkg = withPart(loaded.package, header),
    body = pkg.parts.get(pkg.mainDocumentPart)!;
  const session = { currentPackage: () => pkg, part: () => body };
  expect(partOfNodeId(session, 'lid:history:1') === header).toBe(true);
  expect(partOfNodeId(session, '/word/header1.xml#0') === header).toBe(true);
  expect(partOfNodeId(session, '/word/missing.xml#0')).toBeNull();
});
