import { expect } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlElement,
} from '@docx-editor.dev/core/store';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export function wordPart(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return read.package.parts.get(read.package.mainDocumentPart)!;
}
export function paraIds(node: OoxmlNode): string[] {
  if (node.kind === 'textValue') return [];
  return [
    ...(node.kind === 'paragraph'
      ? node.attributes
          .filter((a) => a.namespaceUri === W14 && a.localName === 'paraId')
          .map((a) => a.value)
      : []),
    ...node.children.flatMap(paraIds),
  ];
}
/** Compare born identities by their validated meanings; imported Word identities stay exact. */
export function bornWordBody(pkg: OoxmlPackage, imported: ReadonlySet<string>) {
  const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw Error(reopened.reason);
  const saved = reopened.package,
    part = saved.parts.get(saved.mainDocumentPart)!,
    ids = paraIds(part.root);
  expect(new Set(ids).size).toBe(ids.length);
  expect(ids.every((id) => /^[0-9A-F]{8}$/.test(id))).toBe(true);
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const value = (node: OoxmlNode, name: string) =>
    node.kind !== 'textValue'
      ? node.attributes.find((a) => a.namespaceUri === W && a.localName === name)?.value
      : undefined;
  const numberKey = (id: string): string => {
    if (id === '0') return 'none';
    const parts = [...saved.parts.values()].filter(
      (candidate) =>
        candidate.contentType ===
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
    );
    expect(parts).toHaveLength(1);
    const numbering = parts[0]!;
    const instances = numbering.root.children.filter(
      (node) => node.kind !== 'textValue' && node.localName === 'num' && value(node, 'numId') === id
    );
    expect(instances).toHaveLength(1);
    const instance = instances[0]!;
    if (instance.kind === 'textValue') throw Error('Invalid numbering instance');
    const references = instance.children.filter(
      (node) => node.kind !== 'textValue' && node.localName === 'abstractNumId'
    );
    expect(references).toHaveLength(1);
    const abstractId = value(references[0]!, 'val');
    const definitions = numbering.root.children.filter(
      (node) =>
        node.kind !== 'textValue' &&
        node.localName === 'abstractNum' &&
        value(node, 'abstractNumId') === abstractId
    );
    expect(definitions).toHaveLength(1);
    const definition = definitions[0]!;
    if (definition.kind === 'textValue') throw Error('Invalid abstract numbering');
    const abstractKey = canonicalOoxmlFingerprint({
      ...numbering,
      root: {
        ...definition,
        attributes: definition.attributes.filter(
          (a) => !(a.namespaceUri === W && a.localName === 'abstractNumId')
        ),
      } as OoxmlElement,
    });
    return canonicalOoxmlFingerprint({
      ...numbering,
      root: {
        ...instance,
        attributes: instance.attributes.filter(
          (a) => !(a.namespaceUri === W && a.localName === 'numId')
        ),
        children: instance.children.map((node) =>
          node === references[0] && node.kind !== 'textValue'
            ? ({
                ...node,
                attributes: node.attributes.map((a) =>
                  a.namespaceUri === W && a.localName === 'val' ? { ...a, value: abstractKey } : a
                ),
              } as OoxmlElement)
            : node
        ),
      } as OoxmlElement,
    });
  };
  let born = 0;
  const visit = (node: OoxmlNode, parentBorn = false): OoxmlNode => {
    if (node.kind === 'textValue') return node;
    const identity = node.attributes.find(
      (a) => a.namespaceUri === W14 && a.localName === 'paraId'
    )?.value;
    const owned =
      parentBorn || (node.kind === 'paragraph' && !!identity && !imported.has(identity));
    const marker = node.kind === 'paragraph' && owned ? `born-${born++}` : null;
    const attributes = node.attributes.map((a) => {
      if (marker && a.namespaceUri === W14 && ['paraId', 'textId'].includes(a.localName))
        return { ...a, value: marker };
      if (owned && node.localName === 'numId' && a.namespaceUri === W && a.localName === 'val')
        return { ...a, value: numberKey(a.value) };
      if (owned && a.namespaceUri === R && a.localName === 'id') {
        const relationships = (saved.relationships.get(saved.mainDocumentPart) ?? []).filter(
          (entry) => entry.id === a.value
        );
        expect(relationships).toHaveLength(1);
        const relationship = relationships[0]!;
        return {
          ...a,
          value: JSON.stringify([
            relationship.type,
            relationship.rawTarget,
            relationship.targetMode,
          ]),
        };
      }
      return a;
    });
    return {
      ...node,
      attributes,
      children: node.children.map((child) => visit(child, owned)),
    } as OoxmlElement;
  };
  return canonicalOoxmlFingerprint({ ...part, root: visit(part.root) as OoxmlElement });
}
