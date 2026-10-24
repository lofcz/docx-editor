import { expect } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { table } from './target-table-fixture';
export const context = {
  documentId: 'tracked-cell-atoms',
  epoch: 'native',
  author: 'AI',
  date: '2026-10-24T12:00:00Z',
};
export const atoms =
  '$x$ ![Cat](asset:cat "Picture") [Ada](mention:user) <source title="Paper" url="https://example.test" />';
export const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
export const decodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
export const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => name.includes('/media/'));
export const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
export const allocate = () => {
  let id = 0;
  return () => `cell-atom-${++id}`;
};
// Bind only fresh Word control identities. Keep all other direct-execution native structure exact.
export function bindControls(
  expected: OoxmlPackage,
  actual: OoxmlPackage,
  before: OoxmlPackage
): OoxmlPackage {
  const left = [...buildBlockIndex(expected).byId.values()],
    right = [...buildBlockIndex(actual).byId.values()];
  expect(right).toHaveLength(left.length);
  let part = expected.parts.get(expected.mainDocumentPart)!;
  for (let i = 0; i < left.length; i++) {
    expect({ ...left[i]!.payload, id: right[i]!.payload.id }).toEqual(right[i]!.payload);
    const a = findNode(actual.parts.get(actual.mainDocumentPart)!, right[i]!.controlId)!;
    const e = findNode(part, left[i]!.controlId)!;
    if (a.kind !== 'contentControl' || e.kind !== 'contentControl') throw Error('Missing control.');
    const property = (control: typeof a, name: string) =>
      control.children
        .find((n) => n.kind === 'contentControlProperties')!
        .children.find((n) => n.localName === name)!;
    for (const name of ['id', 'tag']) {
      const source = property(a, name),
        target = property(e, name);
      if (source?.kind !== 'generic' || target?.kind !== 'generic')
        throw Error('Missing Word identity.');
      const value = source.attributes.find(
        (attr) => attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === 'val'
      )!.value;
      if (name === 'id') expect(value).toMatch(/^[1-9][0-9]*$/);
      const changed = replaceNode(part, target.id, {
        ...target,
        attributes: target.attributes.map((attr) =>
          attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === 'val'
            ? { ...attr, value }
            : attr
        ),
      });
      if (!changed.ok) throw Error('Cannot bind fresh identity.');
      part = changed.part;
    }
  }
  const paragraphs = (node: OoxmlNode): OoxmlNode[] =>
    node.kind === 'textValue'
      ? []
      : [...(node.kind === 'paragraph' ? [node] : []), ...node.children.flatMap(paragraphs)];
  const expectedParagraphs = paragraphs(part.root),
    actualParagraphs = paragraphs(actual.parts.get(actual.mainDocumentPart)!.root);
  expect(actualParagraphs).toHaveLength(expectedParagraphs.length);
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  for (let i = 0; i < expectedParagraphs.length; i++) {
    const target = expectedParagraphs[i]!,
      source = actualParagraphs[i]!;
    if (target.kind !== 'paragraph' || source.kind !== 'paragraph')
      throw Error('Missing paragraph.');
    const identity = (node: typeof target) =>
      node.attributes.filter(
        (a) => a.namespaceUri === W14 && ['paraId', 'textId'].includes(a.localName)
      );
    if (findNode(before.parts.get(before.mainDocumentPart)!, target.id)) {
      expect(identity(source)).toEqual(identity(target));
      continue;
    }
    const attrs = identity(source);
    expect(attrs).toHaveLength(2);
    for (const attr of attrs) expect(attr.value).toMatch(/^[0-9A-F]{8}$/);
    const changed = replaceNode(part, target.id, {
      ...target,
      attributes: [...target.attributes.filter((a) => !identity(target).includes(a)), ...attrs],
    });
    if (!changed.ok) throw Error('Cannot bind fresh paragraph identities.');
    part = changed.part;
  }
  return withPart(expected, part);
}
export const base = table([
  ['A', 'B'],
  ['Keep', 'Peer'],
]);
export const merged =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Peer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
export const cases = [
  { name: 'row', body: base, target: `| A | B |\n| - | - |\n| Keep | Peer |\n| ${atoms} | Tail |` },
  {
    name: 'column',
    body: base,
    target: `| A | New | B |\n| - | - | - |\n| Keep | ${atoms} | Peer |`,
  },
  { name: 'split', body: merged, target: `| A | ${atoms} |\n| - | - |\n| Keep | Peer |` },
];
