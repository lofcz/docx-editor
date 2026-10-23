// Package history can retain typed parts and their declarations independently from opaque ZIP bytes.
// Derive export bytes from retained declaration metadata without writing canonical operations.
import { partNameKey } from './opc-names.ts';
import {
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlAttribute,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlGenericElementNode,
} from './ooxml-tree.ts';
import type { OoxmlPackage } from './ooxml-package.ts';
const NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NAME = '/[Content_Types].xml';
const cached = new WeakMap<OoxmlPackage, OoxmlPackage>();

export function projectContentTypeBytes(pkg: OoxmlPackage): OoxmlPackage {
  const previous = cached.get(pkg);
  if (previous) return previous;
  const entry = [...pkg.partBytes].find(([name]) => partNameKey(name) === partNameKey(NAME));
  if (!entry) return pkg;
  const parsed = readOoxmlPart(new TextDecoder().decode(entry[1]), {
    name: NAME,
    contentType: 'application/xml',
  });
  if (
    !parsed.ok ||
    parsed.part.root.kind !== 'generic' ||
    parsed.part.root.namespaceUri !== NS ||
    parsed.part.root.localName !== 'Types'
  )
    return pkg;
  const root = parsed.part.root;
  const declared = new Map(
    [...pkg.parts.values()].flatMap((part) => {
      const key = partNameKey(part.name),
        type = pkg.contentTypes.overrides.get(key);
      return type ? [[key, { name: part.name, type }] as const] : [];
    })
  );
  const defaults = new Map<string, string>();
  for (const node of root.children) {
    if (node.kind !== 'generic' || node.namespaceUri !== NS || node.localName !== 'Default')
      continue;
    const extension = node.attributes.find(
      (attr) => attr.namespaceUri === '' && attr.localName === 'Extension'
    )?.value;
    const type = node.attributes.find(
      (attr) => attr.namespaceUri === '' && attr.localName === 'ContentType'
    )?.value;
    if (extension && type) defaults.set(extension.toLowerCase(), type);
  }
  let changed = false;
  const children = root.children.map((node): OoxmlNode => {
    if (node.kind !== 'generic' || node.namespaceUri !== NS || node.localName !== 'Override')
      return node;
    const name = node.attributes.find(
      (attr) => attr.namespaceUri === '' && attr.localName === 'PartName'
    )?.value;
    const wanted = name && declared.get(partNameKey(name));
    if (!wanted) return node;
    declared.delete(partNameKey(name!));
    const current = node.attributes.find(
      (attr) => attr.namespaceUri === '' && attr.localName === 'ContentType'
    );
    if (!current || current.value === wanted.type) return node;
    changed = true;
    return {
      ...node,
      attributes: node.attributes.map((attr) =>
        attr === current ? { ...attr, value: wanted.type } : attr
      ),
    } as OoxmlGenericElementNode;
  });
  for (const [key, wanted] of declared) {
    const extension = wanted.name.slice(wanted.name.lastIndexOf('.') + 1).toLowerCase();
    if (defaults.get(extension) === wanted.type) continue;
    changed = true;
    children.push({
      kind: 'generic',
      id: `${root.id}:export:${key}`,
      namespaceUri: NS,
      localName: 'Override',
      prefix: root.prefix,
      namespaceBindings: [],
      attributes: [
        ['PartName', wanted.name],
        ['ContentType', wanted.type],
      ].map(
        ([localName, value]) =>
          ({
            kind: 'genericExtension',
            namespaceUri: '',
            prefix: '',
            localName: localName!,
            value: value!,
          }) as OoxmlAttribute
      ),
      children: [],
    } as OoxmlGenericElementNode);
  }
  const result = changed
    ? Object.freeze({
        ...pkg,
        partBytes: new Map(pkg.partBytes).set(
          entry[0],
          new TextEncoder().encode(
            serializeOoxmlPart({ ...parsed.part, root: { ...root, children } } as OoxmlPart)
          )
        ),
      })
    : pkg;
  cached.set(pkg, result);
  return result;
}
