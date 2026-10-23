import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  readOoxmlPart,
  replaceChildren,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

/** Author the exact joined suffix independently, retaining the fixture's imported Word IDs. */
export function rejectedBornPeerBody(pkg: OoxmlPackage): string {
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const body = bodyStoryRoot(part);
  if (!body || body.kind === 'textValue') throw Error('Missing imported body');
  const after = body.children.find(
    (node) => node.kind === 'paragraph' && paragraphTextOf(part, node.id) === 'After'
  );
  if (after?.kind !== 'paragraph') throw Error('Missing imported suffix paragraph');
  const seed = readOoxmlPart(
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t xml:space="preserve"> Peer</w:t></w:r></w:p>',
    { name: '/r19-peer-oracle.xml', contentType: part.contentType }
  );
  if (!seed.ok) throw Error(seed.reason);
  const peerRun = seed.part.root.children[0];
  if (peerRun?.kind !== 'run') throw Error('Invalid independent suffix seed');
  const patched = replaceChildren(part, after.id, [peerRun, ...after.children]);
  if (!patched.ok) throw Error(JSON.stringify(patched.issues));
  const reopened = readOoxmlPackage(writeOoxmlPackage(withPart(pkg, patched.part)));
  if (!reopened.ok) throw Error(reopened.reason);
  return canonicalOoxmlFingerprint(reopened.package.parts.get(pkg.mainDocumentPart)!);
}
