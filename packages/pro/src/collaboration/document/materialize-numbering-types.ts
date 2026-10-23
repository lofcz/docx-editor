/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import {
  createNodeIdAllocator,
  partNameKey,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlGenericElementNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

const NUMBERING = '/word/numbering.xml';
const TYPES = '/[Content_Types].xml';
const NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const cached = new WeakMap<Uint8Array, Uint8Array>();

/** Reconcile retained numbering metadata with historical content-type bytes, without a journal. */
export function withNumberingTypeBytes(pkg: OoxmlPackage): OoxmlPackage {
  if (!pkg.parts.has(NUMBERING) || pkg.contentTypes.overrides.get(NUMBERING) !== TYPE) return pkg;
  const before = pkg.partBytes.get(TYPES);
  if (!before) return pkg;
  let after = cached.get(before);
  if (!after) {
    const read = readOoxmlPart(new TextDecoder().decode(before), {
      name: TYPES,
      contentType: 'application/xml',
    });
    if (!read.ok || read.part.root.namespaceUri !== NS || read.part.root.localName !== 'Types')
      return pkg;
    const root = read.part.root;
    const matches = root.children.filter(
      (node) =>
        node.kind !== 'textValue' &&
        node.namespaceUri === NS &&
        node.localName === 'Override' &&
        node.attributes.some(
          (attribute) =>
            attribute.localName === 'PartName' &&
            attribute.namespaceUri === '' &&
            partNameKey(attribute.value) === NUMBERING
        )
    );
    const current = matches[0];
    if (
      matches.length === 1 &&
      current?.kind !== 'textValue' &&
      current?.attributes.some(
        (attribute) =>
          attribute.localName === 'ContentType' &&
          attribute.namespaceUri === '' &&
          attribute.value === TYPE
      )
    )
      after = before;
    else {
      const override: OoxmlGenericElementNode = {
        kind: 'generic',
        id: current?.id ?? createNodeIdAllocator(read.part)(),
        namespaceUri: NS,
        localName: 'Override',
        prefix: current?.kind === 'generic' ? current.prefix : root.prefix,
        namespaceBindings: current?.kind === 'generic' ? current.namespaceBindings : [],
        attributes: [
          ...(current?.kind === 'generic'
            ? current.attributes.filter(
                (attribute) =>
                  !(
                    attribute.namespaceUri === '' &&
                    ['PartName', 'ContentType'].includes(attribute.localName)
                  )
              )
            : []),
          { kind: 'genericExtension', namespaceUri: '', localName: 'PartName', value: NUMBERING },
          { kind: 'genericExtension', namespaceUri: '', localName: 'ContentType', value: TYPE },
        ],
        children: current?.kind === 'generic' ? current.children : [],
      };
      const children = root.children.filter((node) => !matches.includes(node));
      children.push(override);
      after = new TextEncoder().encode(
        serializeOoxmlPart({ ...read.part, root: { ...root, children } })
      );
    }
    cached.set(before, after);
  }
  return after === before
    ? pkg
    : Object.freeze({ ...pkg, partBytes: new Map(pkg.partBytes).set(TYPES, after) });
}
