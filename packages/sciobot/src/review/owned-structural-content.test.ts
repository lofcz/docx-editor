import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  replaceNode,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readReviewGroups, persistReviewGroups } from './groups';
import {
  captureOwnedProposalContentProofs,
  verifyOwnedProposalContentProofs,
} from './owned-proposal-content';
const revision = { documentId: 'structural-proof', epoch: 'load', revision: 0 };
const insertion = (text: string) =>
  `<w:ins w:id="7" w:author="Bot"><w:r><w:t>${text}</w:t></w:r></w:ins>`;
function setup(tail: string) {
  const loaded = readOoxmlPackage(
    docxFixture(
      '<w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="Bot"/></w:rPr></w:pPr>' +
        insertion('Owned') +
        tail +
        '</w:p>'
    )
  );
  if (!loaded.ok) throw Error(loaded.reason);
  const baseline = readOoxmlPackage(docxFixture('<w:p/>'));
  if (!baseline.ok) throw Error(baseline.reason);
  const pkg = persistReviewGroups(
    baseline.package,
    loaded.package,
    [{ partName: loaded.package.mainDocumentPart, addresses: [{ id: '7', author: 'Bot' }] }],
    true
  );
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const read = readReviewGroups(pkg);
  if (!read.success) throw Error(read.error);
  const ownership = read.groups
    .filter((g) => g.native.some((m) => m.address.author === 'Bot'))
    .map((g) => ({
      groupId: g.id,
      native: g.native.map((m) => ({ partName: m.partName, addresses: [m.address] })),
      properties: g.properties,
    }));
  const proof = captureOwnedProposalContentProofs(pkg, revision, ownership);
  if (!proof.success) throw Error(proof.error);
  const all: OoxmlNode[] = [];
  const visit = (n: OoxmlNode) => {
    all.push(n);
    if (n.kind !== 'textValue') n.children.forEach(visit);
  };
  visit(part.root);
  return {
    pkg,
    part,
    all,
    ownership,
    proofs: proof.contentProofs,
    verify: (changed: OoxmlPackage) =>
      verifyOwnedProposalContentProofs(changed, revision, ownership, proof.contentProofs),
  };
}
const peer = '<w:r><w:t>Peer</w:t></w:r>';

test('structural receipts ignore committed trailing text but still detect changed owned content', () => {
  const f = setup(peer),
    paragraph = f.all.find((n) => n.kind === 'paragraph')!;
  const appended = applyTreeOp(f.part, {
    op: 'insertText',
    paragraphId: paragraph.id,
    offset: paragraphTextOf(f.part, paragraph.id)!.length,
    text: 'More',
    bias: 'right',
  });
  if (!appended.ok) throw Error(appended.reason);
  expect(f.verify(withPart(f.pkg, appended.part))).toBeNull();
  const owned = f.all.find((n) => n.kind === 'textValue' && n.value === 'Owned');
  if (owned?.kind !== 'textValue') throw Error('Missing owned text.');
  const changed = replaceNode(f.part, owned.id, { ...owned, value: 'Changed' });
  if (!changed.ok) throw Error('Invalid changed owned text.');
  expect(f.verify(withPart(f.pkg, changed.part))).toMatchObject({ code: 'ambiguous-target' });
});

test('paragraph properties and interior committed text remain protected by structural receipts', () => {
  const f = setup(peer + insertion('Tail'));
  const text = f.all.find((n) => n.kind === 'textValue' && n.value === 'Peer');
  if (text?.kind !== 'textValue') throw Error('Missing interior text.');
  const changed = replaceNode(f.part, text.id, { ...text, value: 'Changed' });
  if (!changed.ok) throw Error('Invalid interior edit.');
  expect(f.verify(withPart(f.pkg, changed.part))).toMatchObject({ code: 'ambiguous-target' });
  const paragraph = f.all.find((n) => n.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw Error('Missing paragraph.');
  const properties = replaceNode(f.part, paragraph.id, {
    ...paragraph,
    attributes: [
      ...paragraph.attributes,
      { kind: 'genericExtension', namespaceUri: '', localName: 'proof-test', value: 'changed' },
    ],
  });
  if (!properties.ok) throw Error('Invalid paragraph metadata.');
  expect(f.verify(withPart(f.pkg, properties.part))).toMatchObject({ code: 'ambiguous-target' });
});

test('a foreign property revision inside a trailing text run remains in the structural proof', () => {
  const f = setup(
    '<w:r><w:rPr><w:rPrChange w:id="9" w:author="Peer"><w:rPr/></w:rPrChange></w:rPr><w:t>Peer</w:t></w:r>'
  );
  const record = f.all.find((n) => n.kind !== 'textValue' && n.localName === 'rPrChange');
  if (record?.kind !== 'generic') throw Error('Missing foreign property revision.');
  const changed = replaceNode(f.part, record.id, {
    ...record,
    attributes: record.attributes.map((a) =>
      a.localName === 'author' ? { ...a, value: 'Other' } : a
    ),
  });
  if (!changed.ok) throw Error('Invalid foreign property edit.');
  const current = withPart(f.pkg, changed.part);
  expect(findNode(current.parts.get(current.mainDocumentPart)!, record.id)).toBeDefined();
  expect(f.verify(current)).toMatchObject({ code: 'ambiguous-target' });
});
