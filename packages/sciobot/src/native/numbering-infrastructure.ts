import {
  partNameKey,
  readOoxmlPart,
  relationshipsOf,
  resolveContentTypeOf,
  resolveInternalTarget,
  withNewPart,
  withRelationship,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
const PART = '/word/numbering.xml';
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';

/** Existing imported parts and relationships remain untouched, including unsupported numbering. */
export function withNativeNumberingInfrastructure(pkg: OoxmlPackage): OoxmlPackage | null {
  if (
    [...pkg.parts.keys(), ...pkg.partBytes.keys()].some((name) => partNameKey(name) === PART) ||
    relationshipsOf(pkg, pkg.mainDocumentPart).some((rel) => rel.type === REL)
  )
    return pkg;
  const parsed = readOoxmlPart(`<w:numbering xmlns:w="${WML_NAMESPACE_URI}"/>`, {
    name: PART,
    contentType: TYPE,
  });
  if (!parsed.ok) return null;
  const next = withNewPart(pkg, PART, parsed.part.root, TYPE);
  const from = pkg.mainDocumentPart.split('/').slice(1, -1),
    to = PART.split('/').slice(1);
  while (from.length && to.length && from[0] === to[0]) {
    from.shift();
    to.shift();
  }
  const related = withRelationship(
    next,
    pkg.mainDocumentPart,
    REL,
    '../'.repeat(from.length) + to.join('/')
  );
  return related.ok && nativeNumberingInfrastructureReady(related.pkg) ? related.pkg : null;
}

/** An initialized empty root has a document lifetime independent from any actor's list edits. */
export function nativeNumberingInfrastructureReady(pkg: OoxmlPackage): boolean {
  const part = pkg.parts.get(PART);
  if (
    !part ||
    part.root.namespaceUri !== WML_NAMESPACE_URI ||
    part.root.localName !== 'numbering' ||
    resolveContentTypeOf(pkg, PART) !== TYPE
  )
    return false;
  return relationshipsOf(pkg, pkg.mainDocumentPart).some((rel) => {
    if (rel.type !== REL || rel.targetMode === 'External') return false;
    const target = resolveInternalTarget(pkg.mainDocumentPart, rel.rawTarget);
    return target.ok && target.partName === PART;
  });
}
